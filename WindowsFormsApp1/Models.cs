using System;
using System.Collections.Generic;
using System.Linq;

namespace NmapInventory
{
    // ── Gerätetypen ──────────────────────────────────────────
    public enum DeviceType
    {
        Unbekannt = 0,
        WindowsPC = 1,
        WindowsServer = 2,
        Linux = 3,
        Drucker = 4,
        Smartphone = 5,
        Tablet = 6,
        NetzwerkGeraet = 7,   // Router, Switch, AP
        NAS = 8,
        SmartTV = 9,
        IoT = 10,
        MacOS = 11,
        Laptop = 12
    }

    public class DeviceInfo
    {
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
        public string MacAddress { get; set; }

        // Erweiterte Nmap-Felder
        public string OS { get; set; }
        public string OSDetails { get; set; }
        public string Vendor { get; set; }
        public List<NmapPort> OpenPorts { get; set; } = new List<NmapPort>();

        // Gerätekategorie — wird automatisch beim Scan gesetzt
        public DeviceType DeviceType { get; set; } = DeviceType.Unbekannt;
        public string DeviceTypeIcon => DeviceTypeHelper.GetIcon(DeviceType);

        // SNMP-Abfrageergebnis — optional, nur wenn SNMP aktiv war
        public SnmpResult SnmpData { get; set; }
    }

    // Einzelner Port mit Service-Info
    public class NmapPort
    {
        public int Port { get; set; }
        public string Protocol { get; set; }   // tcp / udp
        public string State { get; set; }   // open / closed / filtered
        public string Service { get; set; }   // http, ssh, smb ...
        public string Version { get; set; }   // Service-Version
        public string Banner { get; set; }   // Banner-Text
    }

    // Vollständige Nmap-Scan-Details pro Gerät — wird in DB gespeichert
    public class NmapScanDetail
    {
        public int ID { get; set; }
        public int DeviceID { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Vendor { get; set; }
        public string OS { get; set; }
        public string OSDetails { get; set; }
        public string Ports { get; set; }   // Kompakter String: "80/tcp open http"
        public string PortsJson { get; set; }   // JSON für detaillierte Ansicht
        public string RawOutput { get; set; }   // Roher Nmap-Block für dieses Gerät
        public DateTime ScanTime { get; set; }
    }

    public class DatabaseDevice
    {
        public int ID { get; set; }
        public string Zeitstempel { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Vendor { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
        public string OS { get; set; }
        public string Comment { get; set; }
        public DeviceType DeviceType { get; set; } = DeviceType.Unbekannt;
        public string DeviceTypeIcon => DeviceTypeHelper.GetIcon(DeviceType);
    }

    // ── Gerätetyp-Erkennung ──────────────────────────────────
    public static class DeviceTypeHelper
    {
        public static string GetIcon(DeviceType type)
        {
            switch (type)
            {
                case DeviceType.WindowsPC: return "🖥";
                case DeviceType.WindowsServer: return "🗄";
                case DeviceType.Laptop: return "💻";
                case DeviceType.Linux: return "🐧";
                case DeviceType.MacOS: return "🍎";
                case DeviceType.Drucker: return "🖨";
                case DeviceType.Smartphone: return "📱";
                case DeviceType.Tablet: return "📟";
                case DeviceType.NetzwerkGeraet: return "🔀";
                case DeviceType.NAS: return "💾";
                case DeviceType.SmartTV: return "📺";
                case DeviceType.IoT: return "💡";
                default: return "❓";
            }
        }

        public static string GetLabel(DeviceType type)
        {
            switch (type)
            {
                case DeviceType.WindowsPC: return "Windows PC";
                case DeviceType.WindowsServer: return "Windows Server";
                case DeviceType.Laptop: return "Laptop";
                case DeviceType.Linux: return "Linux";
                case DeviceType.MacOS: return "macOS";
                case DeviceType.Drucker: return "Drucker";
                case DeviceType.Smartphone: return "Smartphone";
                case DeviceType.Tablet: return "Tablet";
                case DeviceType.NetzwerkGeraet: return "Netzwerkgerät";
                case DeviceType.NAS: return "NAS";
                case DeviceType.SmartTV: return "Smart TV";
                case DeviceType.IoT: return "IoT-Gerät";
                default: return "Unbekannt";
            }
        }

        /// <summary>
        /// Ermittelt den Gerätetyp anhand von Vendor, OS, Ports und Hostname.
        /// Reihenfolge: Ports (sicherste Erkennung) → OS → Vendor → Hostname
        /// </summary>
        /// <summary>
        /// Prüft ob eine MAC-Adresse "locally administered" ist (zufällig generiert, z.B. Android-Datenschutz).
        /// Bit 1 des ersten Oktetts gesetzt → kein echter OUI → wahrscheinlich Smartphone.
        /// </summary>
        private static bool IsRandomizedMac(string mac)
        {
            if (string.IsNullOrEmpty(mac)) return false;
            try
            {
                byte first = Convert.ToByte(mac.Split(':')[0], 16);
                return (first & 0x02) != 0; // locally administered bit
            }
            catch { return false; }
        }

        public static DeviceType Detect(DeviceInfo device)
        {
            string vendor = (device.Vendor ?? "").ToLower();
            string os = (device.OS ?? device.OSDetails ?? "").ToLower();
            string hostname = (device.Hostname ?? "").ToLower();
            var ports = device.OpenPorts ?? new List<NmapPort>();
            string mac = device.MacAddress ?? "";

            // ── 1. Ports — zuverlässigste Erkennung ──────────
            bool hasPort(int p) => ports.Any(x => x.Port == p);

            // Drucker-Ports: 9100 (RAW), 631 (IPP), 515 (LPD)
            if (hasPort(9100) || hasPort(631) || hasPort(515))
                return DeviceType.Drucker;

            // NAS/Server: Samba + NFS oder viele Ports
            if (hasPort(2049) || (hasPort(445) && hasPort(111)))
                return DeviceType.NAS;

            // Router/Switch: SNMP, Telnet, typische Management-Ports
            if (hasPort(161) || hasPort(162) || (hasPort(23) && !hasPort(3389)))
                return DeviceType.NetzwerkGeraet;

            // RDP → Windows PC oder Server
            if (hasPort(3389))
            {
                if (os.Contains("server")) return DeviceType.WindowsServer;
                if (hostname.Contains("laptop") || hostname.Contains("nb") || hostname.Contains("book"))
                    return DeviceType.Laptop;
                return DeviceType.WindowsPC;
            }

            // SSH ohne Windows → Linux oder NAS
            if (hasPort(22) && !hasPort(3389))
            {
                if (hasPort(80) || hasPort(443)) return DeviceType.Linux;
                return DeviceType.Linux;
            }

            // ── 2. OS-String ──────────────────────────────────
            if (os.Contains("windows"))
            {
                if (os.Contains("server")) return DeviceType.WindowsServer;
                if (os.Contains("phone")) return DeviceType.Smartphone;
                if (hostname.Contains("laptop") || hostname.Contains("nb"))
                    return DeviceType.Laptop;
                return DeviceType.WindowsPC;
            }
            if (os.Contains("linux") || os.Contains("ubuntu") || os.Contains("debian") ||
                os.Contains("centos") || os.Contains("fedora"))
                return DeviceType.Linux;
            if (os.Contains("macos") || os.Contains("mac os") || os.Contains("darwin"))
                return DeviceType.MacOS;
            if (os.Contains("ios") || os.Contains("iphone"))
                return DeviceType.Smartphone;
            if (os.Contains("ipad"))
                return DeviceType.Tablet;
            if (os.Contains("android"))
                return DeviceType.Smartphone;

            // ── 3. MAC-Hersteller ─────────────────────────────
            if (VendorMatches(vendor, "apple"))
            {
                // Apple kann iPhone, iPad, Mac oder AppleTV sein
                if (hostname.Contains("iphone") || hostname.Contains("phone"))
                    return DeviceType.Smartphone;
                if (hostname.Contains("ipad"))
                    return DeviceType.Tablet;
                if (hostname.Contains("apple-tv") || hostname.Contains("appletv"))
                    return DeviceType.SmartTV;
                return DeviceType.MacOS;
            }
            if (VendorMatches(vendor, "samsung", "huawei", "xiaomi", "oneplus", "oppo",
                "motorola", "sony mobile", "lg electronics"))
                return DeviceType.Smartphone;

            // Amazon: Echo/Alexa → IoT, Fire TV → SmartTV, Kindle → Tablet
            if (VendorMatches(vendor, "amazon"))
            {
                if (hostname.Contains("fire") || hostname.Contains("aftm") || hostname.Contains("aftb"))
                    return DeviceType.SmartTV;
                if (hostname.Contains("kindle") || hostname.Contains("kftt") || hostname.Contains("kfjwi"))
                    return DeviceType.Tablet;
                return DeviceType.IoT;   // Echo, Alexa, Ring, etc.
            }

            if (VendorMatches(vendor, "hp", "hewlett", "canon", "epson", "brother",
                "xerox", "lexmark", "kyocera", "ricoh", "konica", "oki data", "zebra"))
                return DeviceType.Drucker;

            if (VendorMatches(vendor, "cisco", "juniper", "mikrotik", "ubiquiti",
                "zyxel", "netgear", "tp-link", "d-link", "asus", "linksys",
                "arcadyan", "avm", "fritzbox", "speedport", "fritz"))
                return DeviceType.NetzwerkGeraet;

            if (VendorMatches(vendor, "synology", "qnap", "buffalo", "western digital",
                "drobo", "promise"))
                return DeviceType.NAS;

            if (VendorMatches(vendor, "nintendo", "sony interactive", "microsoft xbox", "valve"))
                return DeviceType.IoT; // Spielkonsole → IoT als nächstpassender Typ

            if (VendorMatches(vendor, "microsoft"))
            {
                if (hostname.Contains("xbox")) return DeviceType.IoT;
                return DeviceType.WindowsPC;
            }

            // ── 4. Hostname-Muster ────────────────────────────
            if (hostname.Contains("desktop") || hostname.Contains("desktop-") || hostname.Contains("pc-"))
                return DeviceType.WindowsPC;
            if (hostname.Contains("laptop") || hostname.Contains("notebook") ||
                hostname.Contains("-nb-") || hostname.Contains("book"))
                return DeviceType.Laptop;
            if (hostname.Contains("server") || hostname.Contains("srv") ||
                hostname.Contains("dc-") || hostname.Contains("-srv"))
                return DeviceType.WindowsServer;
            if (hostname.Contains("print") || hostname.Contains("drucker") ||
                hostname.Contains("printer") || hostname.Contains("-pr-"))
                return DeviceType.Drucker;
            if (hostname.Contains("phone") || hostname.Contains("iphone") ||
                hostname.Contains("android") || hostname.Contains("mobile"))
                return DeviceType.Smartphone;
            if (hostname.Contains("router") || hostname.Contains("switch") ||
                hostname.Contains("ap-") || hostname.Contains("gateway"))
                return DeviceType.NetzwerkGeraet;
            if (hostname.Contains("nas") || hostname.Contains("diskstation"))
                return DeviceType.NAS;
            if (hostname.Contains("tv") || hostname.Contains("firetv") ||
                hostname.Contains("chromecast") || hostname.Contains("roku"))
                return DeviceType.SmartTV;

            return DeviceType.Unbekannt;
        }

        private static bool VendorMatches(string vendor, params string[] keywords)
            => keywords.Any(k => vendor.Contains(k));
    }

    public class SoftwareInfo
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string InstallLocation { get; set; }
        public string Source { get; set; }
        public string PCName { get; set; }
        public DateTime Timestamp { get; set; }
        public string LastUpdate { get; set; }
    }

    public class DatabaseSoftware
    {
        public int ID { get; set; }
        public int DeviceID { get; set; }
        public string Zeitstempel { get; set; }
        public string PCName { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string LastUpdate { get; set; }
    }

    public class Customer
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }

    public class Location
    {
        public int ID { get; set; }
        public int CustomerID { get; set; }
        public int ParentID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public int Level { get; set; }
    }

    public class LocationIP
    {
        public int ID { get; set; }
        public int LocationID { get; set; }
        public string IPAddress { get; set; }
        public string WorkstationName { get; set; }
    }

    public class NodeData
    {
        public string Type { get; set; }
        public int ID { get; set; }
        public object Data { get; set; }
    }

    public class ScanResult
    {
        public List<DeviceInfo> Devices { get; set; }
        public string RawOutput { get; set; }
    }

    // ── SNMP-Konfiguration ────────────────────────────────────
    public class SnmpSettings
    {
        public int Version { get; set; } = 2;           // 1, 2 oder 3
        public string Community { get; set; } = "public"; // v1/v2c Community-String
        public int Port { get; set; } = 161;
        public int TimeoutMs { get; set; } = 2000;

        // SNMPv3-Felder
        public string Username { get; set; } = "";
        public string AuthPassword { get; set; } = "";
        public string PrivPassword { get; set; } = "";
        public string AuthProtocol { get; set; } = "SHA256";  // SHA256, SHA384 oder SHA512
        public string PrivProtocol { get; set; } = "AES128"; // AES128, AES192 oder AES256
        public string SecurityLevel { get; set; } = "AuthPriv"; // NoAuth, AuthNoPriv, AuthPriv
    }

    // ── SNMP-Abfrageergebnis ──────────────────────────────────
    public class SnmpResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string SysDescr { get; set; }      // 1.3.6.1.2.1.1.1.0
        public string SysObjectID { get; set; }   // 1.3.6.1.2.1.1.2.0
        public string SysUpTime { get; set; }     // 1.3.6.1.2.1.1.3.0
        public string SysContact { get; set; }    // 1.3.6.1.2.1.1.4.0
        public string SysName { get; set; }       // 1.3.6.1.2.1.1.5.0
        public string SysLocation { get; set; }   // 1.3.6.1.2.1.1.6.0
        public string SnmpVersion { get; set; }   // "v1", "v2c" oder "v3"
        public DateTime QueryTime { get; set; } = DateTime.Now;
    }
}