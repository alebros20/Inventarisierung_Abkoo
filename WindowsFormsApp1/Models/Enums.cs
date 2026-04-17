using System.Collections.Generic;
using System.Linq;

namespace NmapInventory
{
    public enum DeviceType
    {
        Unbekannt = 0,
        WindowsPC = 1,
        WindowsServer = 2,
        Linux = 3,
        Drucker = 4,
        Smartphone = 5,
        Tablet = 6,
        NetzwerkGeraet = 7,
        NAS = 8,
        SmartTV = 9,
        IoT = 10,
        MacOS = 11,
        Laptop = 12
    }

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

        private static bool IsRandomizedMac(string mac)
        {
            if (string.IsNullOrEmpty(mac)) return false;
            try
            {
                byte first = System.Convert.ToByte(mac.Split(':')[0], 16);
                return (first & 0x02) != 0;
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

            bool hasPort(int p) => ports.Any(x => x.Port == p);

            if (hasPort(9100) || hasPort(631) || hasPort(515))
                return DeviceType.Drucker;
            if (hasPort(2049) || (hasPort(445) && hasPort(111)))
                return DeviceType.NAS;
            if (hasPort(161) || hasPort(162) || (hasPort(23) && !hasPort(3389)))
                return DeviceType.NetzwerkGeraet;
            if (hasPort(3389))
            {
                if (os.Contains("server")) return DeviceType.WindowsServer;
                if (hostname.Contains("laptop") || hostname.Contains("nb") || hostname.Contains("book"))
                    return DeviceType.Laptop;
                return DeviceType.WindowsPC;
            }
            if (hasPort(22) && !hasPort(3389))
            {
                if (hasPort(80) || hasPort(443)) return DeviceType.Linux;
                return DeviceType.Linux;
            }

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

            if (VendorMatches(vendor, "apple"))
            {
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

            if (VendorMatches(vendor, "amazon"))
            {
                if (hostname.Contains("fire") || hostname.Contains("aftm") || hostname.Contains("aftb"))
                    return DeviceType.SmartTV;
                if (hostname.Contains("kindle") || hostname.Contains("kftt") || hostname.Contains("kfjwi"))
                    return DeviceType.Tablet;
                return DeviceType.IoT;
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
                return DeviceType.IoT;
            if (VendorMatches(vendor, "microsoft"))
            {
                if (hostname.Contains("xbox")) return DeviceType.IoT;
                return DeviceType.WindowsPC;
            }

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
}
