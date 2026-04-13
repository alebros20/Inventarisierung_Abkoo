using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Renci.SshNet;
using System.Windows.Forms;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using Lextm.SharpSnmpLib.Security;
using Microsoft.Win32;

namespace NmapInventory
{
    public class NmapScanner
    {
        private readonly string nmapPath;

        public NmapScanner()
        {
            nmapPath = FindNmap();
        }

        // Sucht nmap.exe in dieser Reihenfolge:
        // 1. Neben der eigenen .exe (für ClickOnce / portable)
        // 2. Standard-Installationspfade (32/64-bit)
        // 3. PATH-Variable (falls nmap global installiert)
        private static string FindNmap()
        {
            // 1. Neben der eigenen .exe
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string local = Path.Combine(appDir, "nmap.exe");
            if (File.Exists(local)) return local;

            // Auch in einem nmap-Unterordner neben der .exe
            string subDir = Path.Combine(appDir, "nmap", "nmap.exe");
            if (File.Exists(subDir)) return subDir;

            // 2. Typische Windows-Installationspfade
            string[] candidates = {
                @"C:\Program Files (x86)\Nmap\nmap.exe",
                @"C:\Program Files\Nmap\nmap.exe",
                Path.Combine(Environment.GetFolderPath(
                    Environment.SpecialFolder.ProgramFilesX86), "Nmap", "nmap.exe"),
                Path.Combine(Environment.GetFolderPath(
                    Environment.SpecialFolder.ProgramFiles), "Nmap", "nmap.exe"),
            };
            foreach (var c in candidates)
                if (File.Exists(c)) return c;

            // 3. Fallback: PATH — Windows sucht selbst
            return "nmap";
        }

        // Gibt den gefundenen Pfad zurück — für Diagnose im UI
        public string NmapPath => nmapPath;
        public bool IsNmapAvailable => nmapPath == "nmap"
            ? IsInPath("nmap")
            : File.Exists(nmapPath);

        private static bool IsInPath(string exe)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = exe,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                    return !string.IsNullOrWhiteSpace(output);
                }
            }
            catch { return false; }
        }

        // ─────────────────────────────────────────────────────────
        // 🔍 DISCOVERY SCAN — ARP, schnell, findet ALLE Geräte + MACs
        // Funktioniert nur im lokalen Subnetz (Layer 2)
        // ─────────────────────────────────────────────────────────
        public ScanResult DiscoveryScan(string target)
        {
            // -sn       kein Port-Scan
            // -PR       ARP-Ping (Layer 2, findet alle Geräte inkl. stiller Hosts)
            // --send-eth Ethernet-Frames direkt senden (sichert MAC-Erkennung)
            // -T4       schneller Timing
            return RunScan(target, "-sn -PR --send-eth -T4", ParseDiscoveryOutput);
        }

        // ─────────────────────────────────────────────────────────
        // 🔬 FERNWARTUNGS-SCAN — optimiert für gemischte RDP+SSH /24
        //
        // Strategie: Geräte in Gruppen à 10 scannen → Fortschritt sichtbar
        // Ports: RDP, SSH, VNC, WinRM, SMB, HTTP/S + Telnet
        // Timing: --min-parallelism 10 --max-retries 1 für Geschwindigkeit
        // ─────────────────────────────────────────────────────────
        public ScanResult DetailScan(string targets)
        {
            const string REMOTE_PORTS =
                "22,23,80,443,445,3389,5900,5985,5986,8080,8443";

            const string SCRIPTS =
                "rdp-info,ssh-hostkey,vnc-info,smb-os-discovery,http-title";

            // --min-parallelism 10  gleichzeitig 10 Hosts scannen
            // --max-retries 1       kein langes Warten auf nicht-antwortende Ports
            // --host-timeout 60s    nach 60s zum nächsten Host
            // -sV --version-light   schnelle Versions-Erkennung (kein Full-Probe)
            string args = $"-sV --version-light -O -T4 --open " +
                          $"-p {REMOTE_PORTS} --script={SCRIPTS} " +
                          $"--min-parallelism 10 --max-retries 1 --host-timeout 60s";

            return RunScan(targets, args, ParseDetailOutput);
        }

        // Scan pro einzelnem Gerät (für Fortschrittsanzeige in MainForm)
        public ScanResult DetailScanSingle(string ip)
        {
            const string REMOTE_PORTS =
                "22,23,80,443,445,3389,5900,5985,5986,8080,8443";
            const string SCRIPTS =
                "rdp-info,ssh-hostkey,vnc-info,smb-os-discovery,http-title";
            string args = $"-sV --version-light -O -T4 --open " +
                          $"-p {REMOTE_PORTS} --script={SCRIPTS} " +
                          $"--max-retries 1 --host-timeout 30s";
            return RunScan(ip, args, ParseDetailOutput);
        }

        // Rückwärtskompatibilität — ruft DiscoveryScan auf
        public ScanResult Scan(string target) => DiscoveryScan(target);

        // ─────────────────────────────────────────────────────────
        // Interner Runner
        // ─────────────────────────────────────────────────────────
        private ScanResult RunScan(string targets, string args, Func<string, List<DeviceInfo>> parser)
        {
            string rawOut = "";
            var devices = new List<DeviceInfo>();
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = nmapPath,
                    Arguments = $"{args} {targets}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                using (var proc = Process.Start(psi))
                {
                    rawOut = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit();
                }
                devices = parser(rawOut);
            }
            catch (Exception ex) { rawOut = $"Fehler: {ex.Message}"; }
            return new ScanResult { Devices = devices, RawOutput = rawOut };
        }

        // ─────────────────────────────────────────────────────────
        // Parser für Discovery-Scan (schnell, nur IP/MAC/Vendor)
        // ─────────────────────────────────────────────────────────
        private List<DeviceInfo> ParseDiscoveryOutput(string output)
        {
            var devices = new List<DeviceInfo>();
            var lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            string ip = "", dnsName = "", mac = "", vendor = "";
            bool isUp = false;

            void Flush()
            {
                if (!isUp || string.IsNullOrEmpty(ip)) return;
                var dev = new DeviceInfo
                {
                    IP         = ip,
                    Hostname   = !string.IsNullOrEmpty(dnsName) ? dnsName : "",   // nur echter DNS-Name
                    MacAddress = mac,
                    Vendor     = vendor,
                    Status     = "up",
                    Ports      = "-",
                    OpenPorts  = new List<NmapPort>()
                };
                dev.DeviceType = DeviceTypeHelper.Detect(dev);
                devices.Add(dev);
            }

            foreach (var line in lines)
            {
                if (line.Contains("Nmap scan report for"))
                {
                    Flush();
                    ip = ""; dnsName = ""; mac = ""; vendor = ""; isUp = false;

                    var m = Regex.Match(line, @"for\s+(\S+)\s+\((\d+\.\d+\.\d+\.\d+)\)");
                    if (m.Success) { dnsName = m.Groups[1].Value; ip = m.Groups[2].Value; }
                    else { var m2 = Regex.Match(line, @"(\d+\.\d+\.\d+\.\d+)"); if (m2.Success) ip = m2.Groups[1].Value; }
                }
                else if (line.Contains("Host is up"))
                    isUp = true;
                else if (line.Contains("MAC Address:"))
                {
                    var m = Regex.Match(line,
                        @"MAC Address:\s+([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})\s*(?:\(([^)]*)\))?");
                    if (m.Success) { mac = m.Groups[1].Value; vendor = m.Groups[2].Value; }
                }
            }
            Flush(); // letztes Gerät
            return devices;
        }

        // ─────────────────────────────────────────────────────────
        // Parser für Detail-Scan (Ports, Services, OS, Skript-Ergebnisse)
        // ─────────────────────────────────────────────────────────
        private List<DeviceInfo> ParseDetailOutput(string output)
        {
            var devices = new List<DeviceInfo>();
            var blocks = Regex.Split(output, @"(?=Nmap scan report for )");

            foreach (var block in blocks)
            {
                if (!block.Contains("Nmap scan report for")) continue;
                var device = new DeviceInfo { OpenPorts = new List<NmapPort>(), Status = "up" };
                var lines = block.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                NmapPort lastPort = null;  // für Skript-Ausgaben dem richtigen Port zuordnen

                foreach (var line in lines)
                {
                    // IP / Hostname
                    if (line.StartsWith("Nmap scan report for"))
                    {
                        var m = Regex.Match(line, @"for\s+(\S+)\s+\((\d+\.\d+\.\d+\.\d+)\)");
                        if (m.Success) { device.Hostname = m.Groups[1].Value; device.IP = m.Groups[2].Value; }
                        else { var m2 = Regex.Match(line, @"(\d+\.\d+\.\d+\.\d+)"); if (m2.Success) device.IP = m2.Groups[1].Value; }
                    }

                    // MAC + Vendor
                    else if (line.Contains("MAC Address:"))
                    {
                        var m = Regex.Match(line,
                            @"MAC Address:\s+([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})\s*(?:\(([^)]*)\))?");
                        if (m.Success) { device.MacAddress = m.Groups[1].Value; device.Vendor = m.Groups[2].Value; }
                    }

                    // Port-Zeile: 3389/tcp open ms-wbt-server Microsoft Terminal Services
                    else if (Regex.IsMatch(line, @"^\d+/(tcp|udp)\s+\w+"))
                    {
                        var m = Regex.Match(line, @"^(\d+)/(tcp|udp)\s+(\S+)\s+(\S+)\s*(.*)$");
                        if (m.Success)
                        {
                            lastPort = new NmapPort
                            {
                                Port = int.Parse(m.Groups[1].Value),
                                Protocol = m.Groups[2].Value,
                                State = m.Groups[3].Value,
                                Service = m.Groups[4].Value,
                                Version = m.Groups[5].Value.Trim()
                            };
                            device.OpenPorts.Add(lastPort);
                        }
                    }

                    // Skript-Ausgaben — dem letzten Port zuordnen
                    else if (line.TrimStart().StartsWith("|") && lastPort != null)
                    {
                        string scriptLine = line.TrimStart('|', '_', ' ', '-');

                        // RDP-Info
                        if (scriptLine.StartsWith("rdp-info:") || scriptLine.Contains("Remote Desktop Protocol"))
                            AppendBanner(lastPort, scriptLine);

                        // SSH Host Key
                        else if (scriptLine.StartsWith("ssh-hostkey:") || scriptLine.Contains("ssh-rsa") || scriptLine.Contains("ecdsa"))
                            AppendBanner(lastPort, scriptLine.Trim());

                        // VNC Info
                        else if (scriptLine.Contains("VNC") || scriptLine.StartsWith("vnc-info:"))
                            AppendBanner(lastPort, scriptLine);

                        // SMB OS Discovery
                        else if (scriptLine.Contains("OS:") || scriptLine.Contains("Computer name:") ||
                                 scriptLine.Contains("Domain name:") || scriptLine.Contains("Workgroup:"))
                        {
                            AppendBanner(lastPort, scriptLine.Trim());
                            // OS aus SMB extrahieren
                            var osMatch = Regex.Match(scriptLine, @"OS:\s*(.+)");
                            if (osMatch.Success && string.IsNullOrEmpty(device.OS))
                                device.OS = osMatch.Groups[1].Value.Trim();
                        }

                        // HTTP Title
                        else if (scriptLine.StartsWith("http-title:") || scriptLine.Contains("Title:"))
                            AppendBanner(lastPort, scriptLine.Trim());

                        // Banner (generisch)
                        else if (scriptLine.StartsWith("banner:"))
                            AppendBanner(lastPort, scriptLine.Substring(7).Trim());
                    }

                    // OS-Erkennung
                    else if (line.StartsWith("OS details:"))
                        device.OSDetails = line.Substring(11).Trim();
                    else if (line.StartsWith("Running:"))
                        device.OS = line.Substring(8).Trim();
                    else if (line.StartsWith("Aggressive OS guesses:") && string.IsNullOrEmpty(device.OS))
                        device.OS = Regex.Match(line.Substring(22), @"^([^(]+)").Groups[1].Value.Trim();
                }

                if (!string.IsNullOrEmpty(device.IP))
                {
                    if (string.IsNullOrEmpty(device.Hostname)) device.Hostname = device.IP;
                    device.Ports = device.OpenPorts.Count > 0
                        ? string.Join(", ", device.OpenPorts.Select(p =>
                            $"{p.Port}/{p.Protocol} {GetFernwartungsLabel(p.Port, p.Service)}"))
                        : "Keine Fernwartungs-Ports offen";
                    device.DeviceType = DeviceTypeHelper.Detect(device); // nach Ports + OS setzen
                    devices.Add(device);
                }
            }
            return devices;
        }

        // Kurzlabel für bekannte Fernwartungs-Ports
        private string GetFernwartungsLabel(int port, string service)
        {
            switch (port)
            {
                case 22: return "SSH";
                case 23: return "Telnet ⚠️";
                case 80: return "HTTP";
                case 443: return "HTTPS";
                case 445: return "SMB";
                case 3389: return "RDP 🖥";
                case 5900: return "VNC";
                case 5985: return "WinRM";
                case 5986: return "WinRM-SSL";
                default: return service;
            }
        }

        private void AppendBanner(NmapPort port, string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            port.Banner = string.IsNullOrEmpty(port.Banner)
                ? text
                : port.Banner + " | " + text;
        }
    }


    public class HardwareManager
    {
        // ── Lokal: WMI-Kurzübersicht ──────────────────────────
        public string GetHardwareInfo()
        {
            var sb = new System.Text.StringBuilder();
            try
            {
                ManagementObjectCollection Get(string wql) =>
                    new ManagementObjectSearcher(wql).Get();
                string Val(ManagementBaseObject o, string k)
                {
                    try { return o[k]?.ToString() ?? ""; } catch { return ""; }
                }

                sb.AppendLine("=== Systemübersicht ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_ComputerSystem"))
                    sb.AppendLine($"Computer: {Val(o, "Name")}  Hersteller: {Val(o, "Manufacturer")}  Modell: {Val(o, "Model")}");

                sb.AppendLine("\n=== Betriebssystem ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_OperatingSystem"))
                    sb.AppendLine($"{Val(o, "Caption")}  Version {Val(o, "Version")}  Build {Val(o, "BuildNumber")}");

                sb.AppendLine("\n=== Prozessor ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_Processor"))
                    sb.AppendLine($"{Val(o, "Name")}  {Val(o, "NumberOfCores")} Kerne / {Val(o, "NumberOfLogicalProcessors")} Threads");

                long ram = 0;
                foreach (ManagementObject o in Get("SELECT * FROM Win32_PhysicalMemory"))
                { long.TryParse(Val(o, "Capacity"), out long c); ram += c; }
                sb.AppendLine($"\n=== RAM ===\n{ram / (1024 * 1024 * 1024)} GB gesamt");

                sb.AppendLine("\n=== Laufwerke ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3"))
                {
                    long.TryParse(Val(o, "Size"), out long s);
                    long.TryParse(Val(o, "FreeSpace"), out long f);
                    sb.AppendLine($"{Val(o, "Name")} ({Val(o, "VolumeName")})  {s / (1024 * 1024 * 1024)} GB gesamt, {f / (1024 * 1024 * 1024)} GB frei");
                }
            }
            catch (Exception ex) { sb.AppendLine($"Fehler: {ex.Message}"); }
            return sb.ToString();
        }

        // ── Remote: msinfo32 /computer — exakt wie msinfo32 ──
        public string GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            try
            {
                string tempFile = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), $"msinfo_{Guid.NewGuid():N}.txt");

                // msinfo32 /computer verbindet sich direkt zum Zielrechner
                // und exportiert exakt dieselben Daten wie die msinfo32-Oberfläche
                var psi = new ProcessStartInfo
                {
                    FileName = "msinfo32.exe",
                    Arguments = $"/computer {computerIP} /report \"{tempFile}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(psi))
                {
                    // msinfo32 remote kann 30-60 Sek dauern
                    bool finished = proc.WaitForExit(90000);
                    if (!finished)
                    {
                        proc.Kill();
                        return "Zeitüberschreitung — msinfo32 hat nicht innerhalb von 90 Sekunden geantwortet.\n" +
                               "Bitte prüfe ob der Zielrechner erreichbar ist.";
                    }
                }

                if (!System.IO.File.Exists(tempFile))
                    return FallbackRemoteWmi(computerIP, username, password);

                // msinfo32 schreibt Unicode (UTF-16)
                string content = System.IO.File.ReadAllText(tempFile, System.Text.Encoding.Unicode);
                System.IO.File.Delete(tempFile);

                if (string.IsNullOrWhiteSpace(content))
                    return FallbackRemoteWmi(computerIP, username, password);

                return content;
            }
            catch (Exception)
            {
                // msinfo32 nicht verfügbar oder Fehler → WMI-Fallback
                return FallbackRemoteWmi(computerIP, username, password);
            }
        }

        // Fallback wenn msinfo32 /computer nicht klappt: vollständige WMI-Abfrage
        private string FallbackRemoteWmi(string computerIP, string username, string password)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("(msinfo32 nicht verfügbar — WMI-Abfrage)\n");
            try
            {
                var options = BuildConnectionOptions(username, password);
                var scope = new ManagementScope($"\\\\{computerIP}\\root\\cimv2", options);
                scope.Connect();

                ManagementObjectCollection Get(string wql) =>
                    new ManagementObjectSearcher(scope, new ObjectQuery(wql)).Get();
                string Val(ManagementBaseObject o, string k)
                {
                    try { return o[k]?.ToString() ?? ""; } catch { return ""; }
                }

                sb.AppendLine("=== Systemübersicht ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_ComputerSystem"))
                {
                    sb.AppendLine($"Computername     : {Val(o, "Name")}");
                    sb.AppendLine($"Hersteller       : {Val(o, "Manufacturer")}");
                    sb.AppendLine($"Modell           : {Val(o, "Model")}");
                    sb.AppendLine($"Systemtyp        : {Val(o, "SystemType")}");
                    long.TryParse(Val(o, "TotalPhysicalMemory"), out long ram);
                    sb.AppendLine($"Physischer RAM   : {ram / (1024 * 1024 * 1024)} GB");
                }

                sb.AppendLine("\n=== Betriebssystem ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_OperatingSystem"))
                {
                    sb.AppendLine($"BS-Name          : {Val(o, "Caption")}");
                    sb.AppendLine($"Version          : {Val(o, "Version")}");
                    sb.AppendLine($"Build            : {Val(o, "BuildNumber")}");
                    sb.AppendLine($"Architektur      : {Val(o, "OSArchitecture")}");
                    sb.AppendLine($"Letzter Start    : {Val(o, "LastBootUpTime")}");
                }

                sb.AppendLine("\n=== BIOS ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_BIOS"))
                {
                    sb.AppendLine($"Version          : {Val(o, "SMBIOSBIOSVersion")}");
                    sb.AppendLine($"Hersteller       : {Val(o, "Manufacturer")}");
                    sb.AppendLine($"Datum            : {Val(o, "ReleaseDate")}");
                    sb.AppendLine($"Seriennummer     : {Val(o, "SerialNumber")}");
                }

                sb.AppendLine("\n=== Prozessor ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_Processor"))
                {
                    sb.AppendLine($"Name             : {Val(o, "Name")}");
                    sb.AppendLine($"Kerne            : {Val(o, "NumberOfCores")}");
                    sb.AppendLine($"Logische Proz.   : {Val(o, "NumberOfLogicalProcessors")}");
                    sb.AppendLine($"Max. Takt        : {Val(o, "MaxClockSpeed")} MHz");
                }

                sb.AppendLine("\n=== Arbeitsspeicher (Module) ===");
                int slot = 0;
                foreach (ManagementObject o in Get("SELECT * FROM Win32_PhysicalMemory"))
                {
                    slot++;
                    long.TryParse(Val(o, "Capacity"), out long cap);
                    sb.AppendLine($"Modul {slot}          : {Val(o, "DeviceLocator")}  {cap / (1024 * 1024 * 1024)} GB  {Val(o, "Speed")} MHz  {Val(o, "Manufacturer")}");
                }

                sb.AppendLine("\n=== Laufwerke ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_DiskDrive"))
                {
                    long.TryParse(Val(o, "Size"), out long s);
                    sb.AppendLine($"Modell           : {Val(o, "Model")}  {s / (1024 * 1024 * 1024)} GB  {Val(o, "InterfaceType")}");
                }

                sb.AppendLine("\n=== Logische Datenträger ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3"))
                {
                    long.TryParse(Val(o, "Size"), out long s);
                    long.TryParse(Val(o, "FreeSpace"), out long f);
                    sb.AppendLine($"{Val(o, "Name")} ({Val(o, "VolumeName")})  {s / (1024 * 1024 * 1024)} GB gesamt, {f / (1024 * 1024 * 1024)} GB frei  {Val(o, "FileSystem")}");
                }

                sb.AppendLine("\n=== Grafikkarte ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_VideoController"))
                {
                    long.TryParse(Val(o, "AdapterRAM"), out long vram);
                    sb.AppendLine($"{Val(o, "Name")}  {vram / (1024 * 1024)} MB  Treiber: {Val(o, "DriverVersion")}");
                }

                sb.AppendLine("\n=== Netzwerkadapter ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true"))
                {
                    var ips = o["IPAddress"] as string[];
                    sb.AppendLine($"{Val(o, "Description")}");
                    sb.AppendLine($"  MAC: {Val(o, "MACAddress")}  IP: {(ips?.Length > 0 ? string.Join(", ", ips) : "-")}");
                }

                sb.AppendLine("\n=== Drucker ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_Printer"))
                    sb.AppendLine($"{Val(o, "Name")}  Standard: {Val(o, "Default")}");

                sb.AppendLine("\n=== Autostart ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_StartupCommand"))
                    sb.AppendLine($"{Val(o, "Name")} — {Val(o, "Command")}  [{Val(o, "Location")}]");

                sb.AppendLine("\n=== Laufende Dienste ===");
                foreach (ManagementObject o in Get("SELECT * FROM Win32_Service WHERE State='Running'"))
                    sb.AppendLine($"{Val(o, "DisplayName"),-45} [{Val(o, "Name")}]");
            }
            catch (UnauthorizedAccessException)
            {
                sb.AppendLine("\nFEHLER: Zugriff verweigert.");
                sb.AppendLine("→ UAC-Fix auf Zielgerät: reg add \"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f");
            }
            catch (Exception ex) { sb.AppendLine($"\nFEHLER: {ex.Message}"); }
            return sb.ToString();
        }

        internal static ConnectionOptions BuildConnectionOptions(string username, string password)
        {
            var options = new ConnectionOptions
            {
                Authentication = AuthenticationLevel.PacketPrivacy,
                Impersonation = ImpersonationLevel.Impersonate,
                EnablePrivileges = true
            };
            if (!string.IsNullOrEmpty(username))
            {
                options.Username = username;
                options.Password = password ?? "";
            }
            return options;
        }
    }

    /// <summary>
    /// Ermittelt eine eindeutige Geräte-ID per WMI (Windows) oder SSH (Linux/macOS).
    /// Fallback-Kette: SMBIOS UUID → MachineGuid (Win) / machine-id (Linux) → HW-UUID (Mac) → MAC
    /// </summary>
    public class DeviceIdentifier
    {
        public class IdentifyResult
        {
            public string UniqueID { get; set; }
            public string Source { get; set; }   // "SMBIOS", "MachineGuid", "machine-id", "HW-UUID", "MAC"
            public List<(string Mac, string IP, string Type)> Interfaces { get; set; } = new List<(string, string, string)>();
        }

        private static readonly string[] INVALID_UUIDS = {
            "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF",
            "00000000-0000-0000-0000-000000000000",
            "03000200-0400-0500-0006-000700080009"  // VMware-Dummy
        };

        /// <summary>
        /// Ermittelt die UniqueID eines Windows-Geräts via WMI.
        /// Versucht: 1) SMBIOS UUID  2) MachineGuid aus Registry
        /// Sammelt zusätzlich alle Netzwerk-Interfaces (MAC + IP).
        /// </summary>
        public IdentifyResult IdentifyWindows(string ip, string username, string password)
        {
            var result = new IdentifyResult();
            try
            {
                var options = HardwareManager.BuildConnectionOptions(username, password);
                var scope = new ManagementScope($@"\\{ip}\root\cimv2", options)
                {
                    Options = { Timeout = TimeSpan.FromSeconds(15) }
                };
                scope.Connect();

                // 1) SMBIOS UUID von Win32_ComputerSystemProduct
                try
                {
                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT UUID FROM Win32_ComputerSystemProduct")))
                    foreach (ManagementObject o in searcher.Get())
                    {
                        string uuid = o["UUID"]?.ToString()?.Trim().ToUpperInvariant();
                        if (!string.IsNullOrEmpty(uuid) && !IsInvalidUUID(uuid))
                        {
                            result.UniqueID = uuid;
                            result.Source = "SMBIOS";
                            break;
                        }
                    }
                }
                catch { }

                // 2) Fallback: MachineGuid aus Registry (via StdRegProv)
                if (string.IsNullOrEmpty(result.UniqueID))
                {
                    try
                    {
                        var regScope = new ManagementScope($@"\\{ip}\root\default", options)
                        {
                            Options = { Timeout = TimeSpan.FromSeconds(10) }
                        };
                        regScope.Connect();
                        var regClass = new ManagementClass(regScope, new ManagementPath("StdRegProv"), null);
                        var inParams = regClass.GetMethodParameters("GetStringValue");
                        inParams["hDefKey"] = 0x80000002; // HKLM
                        inParams["sSubKeyName"] = @"SOFTWARE\Microsoft\Cryptography";
                        inParams["sValueName"] = "MachineGuid";
                        var outParams = regClass.InvokeMethod("GetStringValue", inParams, null);
                        string guid = outParams["sValue"]?.ToString();
                        if (!string.IsNullOrEmpty(guid))
                        {
                            result.UniqueID = guid.Trim().ToUpperInvariant();
                            result.Source = "MachineGuid";
                        }
                    }
                    catch { }
                }

                // Netzwerk-Interfaces sammeln
                try
                {
                    using (var searcher = new ManagementObjectSearcher(scope,
                        new ObjectQuery("SELECT Description, MACAddress, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true")))
                    foreach (ManagementObject o in searcher.Get())
                    {
                        string mac = o["MACAddress"]?.ToString();
                        if (string.IsNullOrEmpty(mac)) continue;
                        var ips = o["IPAddress"] as string[];
                        string ifIp = ips?.Length > 0 ? ips[0] : "";
                        string desc = o["Description"]?.ToString() ?? "";
                        string ifType = desc.ToLower().Contains("wi-fi") || desc.ToLower().Contains("wireless") ? "WiFi" : "Ethernet";
                        result.Interfaces.Add((mac.Trim().ToUpperInvariant().Replace('-', ':'), ifIp, ifType));
                    }
                }
                catch { }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Ermittelt die UniqueID eines Linux/macOS-Geräts via SSH.
        /// Linux: 1) /etc/machine-id  2) /sys/class/dmi/id/product_uuid
        /// macOS: system_profiler Hardware UUID
        /// Sammelt zusätzlich alle Netzwerk-Interfaces.
        /// </summary>
        public IdentifyResult IdentifyViaSsh(string ip, string username, string password, int port = 22)
        {
            var result = new IdentifyResult();
            try
            {
                using (var client = new SshClient(
                    new Renci.SshNet.ConnectionInfo(ip, port, username,
                        new Renci.SshNet.PasswordAuthenticationMethod(username, password))
                    { Timeout = TimeSpan.FromSeconds(15) }))
                {
                    client.Connect();
                    if (!client.IsConnected) return result;

                    // Betriebssystem erkennen
                    string uname = RunSshCommand(client, "uname -s")?.Trim().ToLower() ?? "";

                    if (uname == "darwin")
                    {
                        // macOS: Hardware UUID
                        string hwUuid = RunSshCommand(client, "system_profiler SPHardwareDataType 2>/dev/null | grep 'Hardware UUID'");
                        if (!string.IsNullOrEmpty(hwUuid))
                        {
                            string uuid = hwUuid.Split(':').LastOrDefault()?.Trim().ToUpperInvariant();
                            if (!string.IsNullOrEmpty(uuid) && !IsInvalidUUID(uuid))
                            {
                                result.UniqueID = uuid;
                                result.Source = "HW-UUID";
                            }
                        }
                    }
                    else
                    {
                        // Linux: /etc/machine-id (kein root nötig)
                        string machineId = RunSshCommand(client, "cat /etc/machine-id 2>/dev/null")?.Trim();
                        if (!string.IsNullOrEmpty(machineId) && machineId.Length >= 32)
                        {
                            result.UniqueID = machineId.ToUpperInvariant();
                            result.Source = "machine-id";
                        }

                        // Fallback: SMBIOS UUID (braucht root)
                        if (string.IsNullOrEmpty(result.UniqueID))
                        {
                            string smbios = RunSshCommand(client, "cat /sys/class/dmi/id/product_uuid 2>/dev/null || sudo dmidecode -s system-uuid 2>/dev/null")?.Trim();
                            if (!string.IsNullOrEmpty(smbios) && !IsInvalidUUID(smbios.ToUpperInvariant()))
                            {
                                result.UniqueID = smbios.ToUpperInvariant();
                                result.Source = "SMBIOS";
                            }
                        }
                    }

                    // Netzwerk-Interfaces sammeln (funktioniert auf Linux + macOS)
                    string ifOutput = RunSshCommand(client, "ip -o link show 2>/dev/null || ifconfig -a 2>/dev/null");
                    if (!string.IsNullOrEmpty(ifOutput))
                    {
                        // ip -o link: "2: eth0: ... link/ether AA:BB:CC:DD:EE:FF ..."
                        var macRegex = new Regex(@"(?:link/ether|ether)\s+([0-9a-fA-F:]{17})", RegexOptions.IgnoreCase);
                        var ifNameRegex = new Regex(@"^\d+:\s+(\S+?):", RegexOptions.Multiline);
                        var matches = macRegex.Matches(ifOutput);
                        var nameMatches = ifNameRegex.Matches(ifOutput);

                        for (int i = 0; i < matches.Count; i++)
                        {
                            string mac = matches[i].Groups[1].Value.Trim().ToUpperInvariant();
                            if (mac == "00:00:00:00:00:00") continue;
                            string ifName = i < nameMatches.Count ? nameMatches[i].Groups[1].Value : "";
                            string ifType = ifName.StartsWith("wl") || ifName.Contains("wifi") ? "WiFi" : "Ethernet";
                            result.Interfaces.Add((mac, "", ifType));
                        }
                    }

                    client.Disconnect();
                }
            }
            catch { }
            return result;
        }

        private static string RunSshCommand(SshClient client, string command)
        {
            try
            {
                using (var cmd = client.RunCommand(command))
                    return cmd.ExitStatus == 0 ? cmd.Result : null;
            }
            catch { return null; }
        }

        private static bool IsInvalidUUID(string uuid)
        {
            if (string.IsNullOrEmpty(uuid)) return true;
            return INVALID_UUIDS.Any(inv => uuid.Equals(inv, StringComparison.OrdinalIgnoreCase));
        }
    }

    public class SoftwareManager
    {
        // Registry-Pfade wo Windows installierte Software speichert
        private static readonly string[] REG_PATHS = {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        };

        // ── Lokal: Registry direkt auslesen ──────────────────
        public List<SoftwareInfo> GetInstalledSoftware()
        {
            var software = new List<SoftwareInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var regPath in REG_PATHS)
                {
                    using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regPath))
                    {
                        if (key == null) continue;
                        foreach (var subKeyName in key.GetSubKeyNames())
                        {
                            try
                            {
                                using (var sub = key.OpenSubKey(subKeyName))
                                {
                                    if (sub == null) continue;
                                    string name = sub.GetValue("DisplayName")?.ToString();
                                    if (string.IsNullOrWhiteSpace(name)) continue;
                                    // Duplikate überspringen (32/64 bit)
                                    if (!seen.Add(name)) continue;

                                    software.Add(new SoftwareInfo
                                    {
                                        Name = name,
                                        Version = sub.GetValue("DisplayVersion")?.ToString() ?? "",
                                        Publisher = sub.GetValue("Publisher")?.ToString() ?? "",
                                        InstallDate = sub.GetValue("InstallDate")?.ToString() ?? "",
                                        InstallLocation = sub.GetValue("InstallLocation")?.ToString() ?? "",
                                        Source = regPath.Contains("WOW6432") ? "x86" : "x64"
                                    });
                                }
                            }
                            catch { /* einzelner Eintrag fehlerhaft → überspringen */ }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                software.Add(new SoftwareInfo { Name = $"Fehler: {ex.Message}" });
            }

            return software.OrderBy(s => s.Name).ToList();
        }

        // ── Remote: Registry direkt über WMI StdRegProv auslesen ─
        // Kein PowerShell, kein Admin-Share, kein Temp-File nötig.
        // StdRegProv ist eine WMI-Klasse die Remote-Registry-Zugriff erlaubt.
        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            var software = new List<SoftwareInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var options = HardwareManager.BuildConnectionOptions(username, password);
                // StdRegProv liegt in root\default, nicht in root\cimv2
                var scope = new ManagementScope($"\\\\{computerIP}\\root\\default", options);
                scope.Connect();

                const uint HKLM = 0x80000002;
                string[] regPaths = {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };

                var reg = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);

                foreach (var regPath in regPaths)
                {
                    // Alle Unterschlüssel auflisten
                    var inParams = reg.GetMethodParameters("EnumKey");
                    inParams["hDefKey"] = HKLM;
                    inParams["sSubKeyName"] = regPath;
                    var outParams = reg.InvokeMethod("EnumKey", inParams, null);

                    var subKeys = outParams["sNames"] as string[];
                    if (subKeys == null) continue;

                    foreach (var subKey in subKeys)
                    {
                        try
                        {
                            string fullPath = $"{regPath}\\{subKey}";
                            string name = ReadRegString(reg, HKLM, fullPath, "DisplayName");
                            if (string.IsNullOrWhiteSpace(name)) continue;
                            if (!seen.Add(name)) continue; // Duplikat

                            software.Add(new SoftwareInfo
                            {
                                Name = name,
                                Version = ReadRegString(reg, HKLM, fullPath, "DisplayVersion"),
                                Publisher = ReadRegString(reg, HKLM, fullPath, "Publisher"),
                                InstallDate = ReadRegString(reg, HKLM, fullPath, "InstallDate"),
                                InstallLocation = ReadRegString(reg, HKLM, fullPath, "InstallLocation"),
                                Source = regPath.Contains("WOW6432") ? "x86" : "x64"
                            });
                        }
                        catch { /* einzelner Eintrag fehlerhaft → überspringen */ }
                    }
                }
            }
            catch (Exception ex)
            {
                software.Add(new SoftwareInfo { Name = $"Fehler: {ex.Message}" });
            }

            return software.OrderBy(s => s.Name).ToList();
        }

        // Liest einen einzelnen Registry-String-Wert über StdRegProv
        private string ReadRegString(ManagementClass reg, uint hive, string keyPath, string valueName)
        {
            try
            {
                var inParams = reg.GetMethodParameters("GetStringValue");
                inParams["hDefKey"] = hive;
                inParams["sSubKeyName"] = keyPath;
                inParams["sValueName"] = valueName;
                var outParams = reg.InvokeMethod("GetStringValue", inParams, null);
                return outParams["sValue"]?.ToString() ?? "";
            }
            catch { return ""; }
        }

        public void UpdateSoftware(string softwareName, string computerIP, Label statusLabel)
        {
            statusLabel.Text = $"Aktualisiere {softwareName}...";
            MessageBox.Show($"Update-Funktion für '{softwareName}' ist nicht implementiert.", "Info");
            statusLabel.Text = "Bereit";
        }
    }

    // =========================================================
    // =========================================================
    // === LINUX MANAGER — SSH-Abfrage ===
    // =========================================================
    public class LinuxManager
    {
        private readonly int _defaultTimeout = 15;

        public string GetHardwareInfo(string host, string username, string password, string keyFile = "", int port = 22)
        {
            var hw = new System.Text.StringBuilder();
            hw.AppendLine($"=== Linux SSH-Abfrage: {host} ({DateTime.Now:dd.MM.yyyy HH:mm}) ===");

            try
            {
                using (var client = BuildClient(host, port, username, password, keyFile))
                {
                    client.Connect();

                    hw.AppendLine();
                    hw.AppendLine("=== Betriebssystem ===");
                    hw.AppendLine(Run(client, "cat /etc/os-release | grep -E '^(NAME|VERSION|ID)=' | tr -d '\"'"));
                    hw.AppendLine(Run(client, "uname -r"));

                    hw.AppendLine();
                    hw.AppendLine("=== Prozessor ===");
                    hw.AppendLine(Run(client, "lscpu | grep -E 'Model name|CPU\\(s\\)|Thread|Socket|Architecture'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Arbeitsspeicher ===");
                    hw.AppendLine(Run(client, "free -h | grep -E 'Mem|Swap'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Festplatten ===");
                    hw.AppendLine(Run(client, "df -h -x tmpfs -x devtmpfs -x overlay | grep -v 'udev'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Netzwerk ===");
                    hw.AppendLine(Run(client, "ip -br addr show | grep -v '^lo'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Uptime ===");
                    hw.AppendLine(Run(client, "uptime -p 2>/dev/null || uptime"));

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                hw.AppendLine($"Fehler: {ex.Message}");
            }

            return hw.ToString();
        }

        public List<SoftwareInfo> GetInstalledSoftware(string host, string username, string password, string keyFile = "", int port = 22)
        {
            var list = new List<SoftwareInfo>();
            try
            {
                using (var client = BuildClient(host, port, username, password, keyFile))
                {
                    client.Connect();

                    // Paketmanager ermitteln
                    string pm = Run(client, "which dpkg apt rpm pacman 2>/dev/null | head -1");

                    string raw;
                    string source;

                    if (pm.Contains("dpkg") || pm.Contains("apt"))
                    {
                        raw = Run(client, "dpkg-query -W -f='${Package}\\t${Version}\\t${Maintainer}\\n' 2>/dev/null");
                        source = "Linux/dpkg";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: 2);
                    }
                    else if (pm.Contains("rpm"))
                    {
                        raw = Run(client, "rpm -qa --queryformat '%{NAME}\\t%{VERSION}\\t%{VENDOR}\\n' 2>/dev/null");
                        source = "Linux/rpm";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: 2);
                    }
                    else if (pm.Contains("pacman"))
                    {
                        raw = Run(client, "pacman -Q 2>/dev/null");
                        source = "Linux/pacman";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: -1);
                    }
                    else
                    {
                        list.Add(new SoftwareInfo { Name = "Kein unterstützter Paketmanager gefunden", Source = "Linux", PCName = host });
                    }

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                list.Add(new SoftwareInfo { Name = $"SSH-Fehler: {ex.Message}", Source = "Linux", PCName = host });
            }

            return list;
        }

        private SshClient BuildClient(string host, int port, string username, string password, string keyFile)
        {
            if (!string.IsNullOrWhiteSpace(keyFile) && File.Exists(keyFile))
            {
                var keyFiles = new[] { new PrivateKeyFile(keyFile) };
                return new SshClient(new ConnectionInfo(host, port, username,
                    new PrivateKeyAuthenticationMethod(username, keyFiles))
                { Timeout = TimeSpan.FromSeconds(_defaultTimeout) });
            }
            return new SshClient(new ConnectionInfo(host, port, username,
                new PasswordAuthenticationMethod(username, password))
            { Timeout = TimeSpan.FromSeconds(_defaultTimeout) });
        }

        private static string Run(SshClient client, string command)
        {
            using (var cmd = client.CreateCommand(command))
            {
                cmd.CommandTimeout = TimeSpan.FromSeconds(30);
                return (cmd.Execute() ?? "").Trim();
            }
        }

        private static List<SoftwareInfo> ParseTsv(string raw, string host, string source, int nameCol, int versionCol, int publisherCol)
        {
            var list = new List<SoftwareInfo>();
            foreach (var line in raw.Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var cols = line.Split('\t');
                list.Add(new SoftwareInfo
                {
                    Name      = cols.Length > nameCol    && nameCol    >= 0 ? cols[nameCol].Trim()    : line.Trim(),
                    Version   = cols.Length > versionCol && versionCol >= 0 ? cols[versionCol].Trim() : "",
                    Publisher = cols.Length > publisherCol && publisherCol >= 0 ? cols[publisherCol].Trim() : "",
                    Source    = source,
                    PCName    = host,
                    Timestamp = DateTime.Now
                });
            }
            return list;
        }
    }

    // =========================================================
    // === ANDROID MANAGER — ADB over WiFi ===
    // =========================================================
    public class AdbManager
    {
        private readonly string _adbPath;
        public bool IsAdbAvailable => _adbPath != "adb" ? File.Exists(_adbPath) : IsInPath("adb");

        public AdbManager() { _adbPath = FindAdb(); }

        private static string FindAdb()
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates = {
                Path.Combine(appDir, "adb.exe"),
                Path.Combine(appDir, "adb", "adb.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Android", "Sdk", "platform-tools", "adb.exe"),
                @"C:\Program Files\Android\android-sdk\platform-tools\adb.exe",
                @"C:\Android\platform-tools\adb.exe",
            };
            foreach (var c in candidates) if (File.Exists(c)) return c;
            return "adb";
        }

        /// <summary>
        /// Koppelt das Gerät via adb pair (Android 11+ Wireless Debugging).
        /// ipPairPort = "192.168.x.y:PORT" des Koppel-Ports (nicht der Verbindungs-Port!).
        /// </summary>
        public string PairDevice(string ipPairPort, string pairCode)
        {
            // adb pair <ip:pairport> <code>
            string result = RunAdb($"pair {ipPairPort} {pairCode}", out string err);
            // adb pair gibt Erfolg auf stdout aus: "Successfully paired ..."
            if (result.Contains("Successfully") || result.Contains("successfully"))
                return $"Kopplung erfolgreich: {result.Trim()}";
            string combined = (result + "\n" + err).Trim();
            throw new Exception($"Kopplung fehlgeschlagen:\n{combined}");
        }

        /// <summary>
        /// Extrahiert "Hersteller Modell" aus dem Hardware-Info-String (z.B. "Samsung Galaxy A54").
        /// Gibt null zurück wenn nichts gefunden.
        /// </summary>
        public static string ExtractAdbDeviceName(string hwInfo)
        {
            string manufacturer = null, model = null;
            foreach (var line in hwInfo.Split('\n'))
            {
                var t = line.Trim();
                if (t.StartsWith("Hersteller :")) manufacturer = t.Substring(t.IndexOf(':') + 1).Trim();
                else if (t.StartsWith("Modell     :")) model = t.Substring(t.IndexOf(':') + 1).Trim();
            }
            if (!string.IsNullOrEmpty(manufacturer) && !string.IsNullOrEmpty(model))
            {
                // Hersteller nicht doppelt (z.B. "samsung SM-G991B" → "Samsung SM-G991B")
                string mfr = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(manufacturer.ToLower());
                return model.StartsWith(mfr, StringComparison.OrdinalIgnoreCase) ? model : $"{mfr} {model}";
            }
            return model ?? manufacturer;
        }

        public string GetDeviceInfo(string ipPort)
        {
            var hw = new System.Text.StringBuilder();
            hw.AppendLine($"=== Android ADB-Abfrage: {ipPort} ({DateTime.Now:dd.MM.yyyy HH:mm}) ===");
            try
            {
                // Verbinden
                string connectResult = RunAdb($"connect {ipPort}", out string connectErr);
                hw.AppendLine($"  ADB Connect: {connectResult.Trim()}");

                if (connectResult.Contains("failed") || connectResult.Contains("refused") || connectResult.Contains("unable"))
                    throw new Exception($"Verbindung fehlgeschlagen: {connectResult.Trim()}");

                // Gerätestatus prüfen
                string devicesOutput = RunAdb("devices", out _);
                bool unauthorized = devicesOutput.Split('\n')
                    .Any(l => l.Contains(ipPort.Split(':')[0]) && l.Contains("unauthorized"));
                bool offline = devicesOutput.Split('\n')
                    .Any(l => l.Contains(ipPort.Split(':')[0]) && l.Contains("offline"));

                if (unauthorized)
                    throw new Exception(
                        "Gerät nicht autorisiert!\n\n" +
                        "→ Auf dem Android-Gerät erscheint ein Dialog:\n" +
                        "   'ADB-Debugging erlauben?' — bitte ZULASSEN tippen.\n\n" +
                        "Danach erneut 'Android (ADB)' klicken.");

                if (offline)
                    throw new Exception(
                        "Gerät ist offline.\n\n" +
                        "→ Wireless Debugging auf dem Gerät deaktivieren und wieder aktivieren,\n" +
                        "  dann erneut versuchen.");

                string s = $"-s {ipPort}";
                hw.AppendLine();
                hw.AppendLine("=== Gerät ===");
                hw.AppendLine($"  Hersteller : {RunAdb($"{s} shell getprop ro.product.manufacturer", out _).Trim()}");
                hw.AppendLine($"  Modell     : {RunAdb($"{s} shell getprop ro.product.model", out _).Trim()}");
                hw.AppendLine($"  Android    : {RunAdb($"{s} shell getprop ro.build.version.release", out _).Trim()}");
                hw.AppendLine($"  SDK-Level  : {RunAdb($"{s} shell getprop ro.build.version.sdk", out _).Trim()}");
                hw.AppendLine($"  Build      : {RunAdb($"{s} shell getprop ro.build.display.id", out _).Trim()}");
                hw.AppendLine($"  Serial     : {RunAdb($"{s} get-serialno", out _).Trim()}");

                hw.AppendLine();
                hw.AppendLine("=== Arbeitsspeicher ===");
                string memRaw = RunAdb($"{s} shell cat /proc/meminfo", out _);
                foreach (var line in memRaw.Split('\n'))
                    if (line.StartsWith("MemTotal") || line.StartsWith("MemAvailable"))
                        hw.AppendLine($"  {line.Trim()}");

                hw.AppendLine();
                hw.AppendLine("=== Speicher ===");
                hw.AppendLine(RunAdb($"{s} shell df /data | tail -1", out _));
            }
            catch (Exception ex)
            {
                hw.AppendLine();
                hw.AppendLine($"⚠️  {ex.Message}");
            }
            return hw.ToString();
        }

        public List<SoftwareInfo> GetInstalledApps(string ipPort)
        {
            var list = new List<SoftwareInfo>();
            try
            {
                string s = $"-s {ipPort}";
                string raw = RunAdb($"{s} shell pm list packages -i", out _);
                foreach (var line in raw.Split('\n'))
                {
                    // Format: package:com.example.app  installer=com.android.vending
                    var trimmed = line.Trim();
                    if (!trimmed.StartsWith("package:")) continue;

                    string packageName = "";
                    string installer = "";

                    var parts = trimmed.Split(new[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                        packageName = parts[0].Replace("package:", "").Trim();
                    if (parts.Length > 1 && parts[1].StartsWith("installer="))
                        installer = parts[1].Replace("installer=", "").Trim();

                    if (string.IsNullOrEmpty(packageName)) continue;

                    list.Add(new SoftwareInfo
                    {
                        Name      = packageName,
                        Publisher = installer,
                        Source    = "Android/ADB",
                        PCName    = ipPort.Split(':')[0],
                        Timestamp = DateTime.Now
                    });
                }
            }
            catch (Exception ex)
            {
                list.Add(new SoftwareInfo { Name = $"ADB-Fehler: {ex.Message}", Source = "Android/ADB", PCName = ipPort });
            }
            return list;
        }

        private string RunAdb(string args, out string stderr)
        {
            stderr = "";
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName               = _adbPath,
                    Arguments              = args,
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true,
                    CreateNoWindow         = true
                };
                using (var p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit();
                    return output;
                }
            }
            catch (Exception ex) { return $"Fehler: {ex.Message}"; }
        }

        private static bool IsInPath(string exe)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = "where", Arguments = exe, UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
                using (var p = Process.Start(psi)) { string o = p.StandardOutput.ReadToEnd(); p.WaitForExit(); return !string.IsNullOrWhiteSpace(o); }
            }
            catch { return false; }
        }
    }

    // =========================================================
    // === IOS MANAGER — ideviceinfo (libimobiledevice) ===
    // =========================================================
    public class IosManager
    {
        private readonly string _ideviceinfoPath;
        public bool IsAvailable => _ideviceinfoPath != "ideviceinfo" ? File.Exists(_ideviceinfoPath) : IsInPath("ideviceinfo");

        public IosManager() { _ideviceinfoPath = FindTool(); }

        private static string FindTool()
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates = {
                Path.Combine(appDir, "ideviceinfo.exe"),
                Path.Combine(appDir, "libimobiledevice", "ideviceinfo.exe"),
                @"C:\Program Files\libimobiledevice\ideviceinfo.exe",
            };
            foreach (var c in candidates) if (File.Exists(c)) return c;
            return "ideviceinfo";
        }

        /// <summary>
        /// Fragt das verbundene iOS-Gerät ab (USB). udid leer = erstes Gerät.
        /// </summary>
        public string GetDeviceInfo(string udid = "")
        {
            var hw = new System.Text.StringBuilder();
            hw.AppendLine($"=== iOS-Abfrage via USB ({DateTime.Now:dd.MM.yyyy HH:mm}) ===");
            try
            {
                string args = string.IsNullOrEmpty(udid) ? "" : $"-u {udid}";
                string raw  = RunTool("ideviceinfo.exe", args);

                if (string.IsNullOrWhiteSpace(raw))
                    throw new Exception("Kein iOS-Gerät gefunden oder Gerät nicht vertraut.\n→ iPhone: 'Vertrauen' antippen und iTunes / Apple Devices installieren.");

                hw.AppendLine();
                hw.AppendLine("=== Gerät ===");
                hw.AppendLine(ExtractKey(raw, "DeviceName"));
                hw.AppendLine(ExtractKey(raw, "ProductType"));
                hw.AppendLine(ExtractKey(raw, "ProductVersion", "iOS-Version"));
                hw.AppendLine(ExtractKey(raw, "BuildVersion"));
                hw.AppendLine(ExtractKey(raw, "SerialNumber"));
                hw.AppendLine(ExtractKey(raw, "UniqueDeviceID", "UDID"));

                hw.AppendLine();
                hw.AppendLine("=== Speicher ===");
                hw.AppendLine(ExtractKey(raw, "TotalDiskCapacity", "Gesamt"));
                hw.AppendLine(ExtractKey(raw, "TotalDataCapacity", "Daten"));

                hw.AppendLine();
                hw.AppendLine("ℹ️  Softwareliste nicht verfügbar (kein Jailbreak).");
            }
            catch (Exception ex)
            {
                hw.AppendLine($"Fehler: {ex.Message}");
            }
            return hw.ToString();
        }

        public List<string> ListDeviceUdids()
        {
            var list = new List<string>();
            try
            {
                string raw = RunTool("idevice_id.exe", "-l");
                foreach (var line in raw.Split('\n'))
                    if (!string.IsNullOrWhiteSpace(line)) list.Add(line.Trim());
            }
            catch { }
            return list;
        }

        private string RunTool(string exe, string args)
        {
            string fullPath = Path.Combine(Path.GetDirectoryName(_ideviceinfoPath) ?? "", exe);
            if (!File.Exists(fullPath)) fullPath = exe;

            var psi = new ProcessStartInfo
            {
                FileName               = fullPath,
                Arguments              = args,
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true
            };
            using (var p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                return output;
            }
        }

        private static string ExtractKey(string raw, string key, string label = "")
        {
            foreach (var line in raw.Split('\n'))
                if (line.TrimStart().StartsWith(key + ":"))
                    return $"  {(string.IsNullOrEmpty(label) ? key : label),-20}: {line.Split(':')[1].Trim()}";
            return $"  {(string.IsNullOrEmpty(label) ? key : label),-20}: -";
        }

        private static bool IsInPath(string exe)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = "where", Arguments = exe, UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
                using (var p = Process.Start(psi)) { string o = p.StandardOutput.ReadToEnd(); p.WaitForExit(); return !string.IsNullOrWhiteSpace(o); }
            }
            catch { return false; }
        }
    }

    // =========================================================
    // === SNMP MANAGER — v1 / v2c / v3 ===
    // =========================================================
    public class SnmpManager
    {
        // Standard-MIB-II System-OIDs
        private static readonly string OID_SysDescr    = "1.3.6.1.2.1.1.1.0";
        private static readonly string OID_SysObjectID = "1.3.6.1.2.1.1.2.0";
        private static readonly string OID_SysUpTime   = "1.3.6.1.2.1.1.3.0";
        private static readonly string OID_SysContact  = "1.3.6.1.2.1.1.4.0";
        private static readonly string OID_SysName     = "1.3.6.1.2.1.1.5.0";
        private static readonly string OID_SysLocation = "1.3.6.1.2.1.1.6.0";

        /// <summary>
        /// Fragt ein Gerät per SNMP ab. Version wird aus den Settings entnommen.
        /// </summary>
        public SnmpResult Query(string ipAddress, SnmpSettings settings)
        {
            try
            {
                switch (settings.Version)
                {
                    case 1:  return QueryV1(ipAddress, settings);
                    case 2:  return QueryV2c(ipAddress, settings);
                    case 3:  return QueryV3(ipAddress, settings);
                    default: return Error($"Ungültige SNMP-Version: {settings.Version}");
                }
            }
            catch (Exception ex)
            {
                return Error(ex.Message);
            }
        }

        // ── SNMP v1 ──────────────────────────────────────────
        private SnmpResult QueryV1(string ip, SnmpSettings s)
        {
            var community = new OctetString(s.Community);
            var endpoint  = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids      = BuildOidList();

            var result = Messenger.Get(
                VersionCode.V1,
                endpoint,
                community,
                oids,
                s.TimeoutMs);

            return ParseGetResult(result, "v1");
        }

        // ── SNMP v2c ─────────────────────────────────────────
        private SnmpResult QueryV2c(string ip, SnmpSettings s)
        {
            var community = new OctetString(s.Community);
            var endpoint  = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids      = BuildOidList();

            var result = Messenger.Get(
                VersionCode.V2,
                endpoint,
                community,
                oids,
                s.TimeoutMs);

            return ParseGetResult(result, "v2c");
        }

        // ── SNMP v3 ──────────────────────────────────────────
        private SnmpResult QueryV3(string ip, SnmpSettings s)
        {
            var endpoint = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids     = BuildOidList();

            // Auth-Provider bestimmen
            IAuthenticationProvider authProvider = BuildAuthProvider(s);

            // Privacy-Provider bestimmen
            IPrivacyProvider privProvider = BuildPrivacyProvider(s, authProvider);

            // Discovery (Engine-ID ermitteln)
            var discovery = Messenger.GetNextDiscovery(SnmpType.GetRequestPdu);
            var report    = discovery.GetResponse(s.TimeoutMs, endpoint);

            var request = new GetRequestMessage(
                VersionCode.V3,
                Messenger.NextMessageId,
                Messenger.NextRequestId,
                new OctetString(s.Username),
                OctetString.Empty,
                oids,
                privProvider,
                Messenger.MaxMessageSize,
                report);

            var reply = request.GetResponse(s.TimeoutMs, endpoint);

            if (reply.Pdu().ErrorStatus.ToInt32() != 0)
                return Error($"SNMP v3 Fehler: ErrorStatus={reply.Pdu().ErrorStatus}");

            return ParseGetResult(reply.Pdu().Variables.ToList(), "v3");
        }

        // ── Hilfsmethoden ─────────────────────────────────────
        private static List<Variable> BuildOidList() => new List<Variable>
        {
            new Variable(new ObjectIdentifier(OID_SysDescr)),
            new Variable(new ObjectIdentifier(OID_SysObjectID)),
            new Variable(new ObjectIdentifier(OID_SysUpTime)),
            new Variable(new ObjectIdentifier(OID_SysContact)),
            new Variable(new ObjectIdentifier(OID_SysName)),
            new Variable(new ObjectIdentifier(OID_SysLocation)),
        };

        private static SnmpResult ParseGetResult(IList<Variable> vars, string version)
        {
            var r = new SnmpResult { Success = true, SnmpVersion = version };
            foreach (var v in vars)
            {
                string oid = v.Id.ToString();
                string val = v.Data.ToString();

                if (oid == OID_SysDescr)    r.SysDescr    = val;
                else if (oid == OID_SysObjectID) r.SysObjectID = val;
                else if (oid == OID_SysUpTime)   r.SysUpTime   = FormatUptime(val);
                else if (oid == OID_SysContact)  r.SysContact  = val;
                else if (oid == OID_SysName)     r.SysName     = val;
                else if (oid == OID_SysLocation) r.SysLocation = val;
            }
            return r;
        }

        private static string FormatUptime(string raw)
        {
            // SharpSnmpLib liefert TimeTicks als Hundertstelsekunden-String
            if (uint.TryParse(raw, out uint ticks))
            {
                var ts = TimeSpan.FromSeconds(ticks / 100.0);
                return $"{(int)ts.TotalDays}d {ts.Hours:D2}h {ts.Minutes:D2}m";
            }
            return raw;
        }

        private static IAuthenticationProvider BuildAuthProvider(SnmpSettings s)
        {
            if (s.SecurityLevel == "NoAuth")
                return DefaultAuthenticationProvider.Instance;

            switch (s.AuthProtocol?.ToUpper())
            {
                case "SHA384": return new SHA384AuthenticationProvider(new OctetString(s.AuthPassword));
                case "SHA512": return new SHA512AuthenticationProvider(new OctetString(s.AuthPassword));
                default:       return new SHA256AuthenticationProvider(new OctetString(s.AuthPassword));
            }
        }

        private static IPrivacyProvider BuildPrivacyProvider(SnmpSettings s, IAuthenticationProvider auth)
        {
            if (s.SecurityLevel == "NoAuth" || s.SecurityLevel == "AuthNoPriv")
                return new DefaultPrivacyProvider(auth);

            switch (s.PrivProtocol?.ToUpper())
            {
                case "AES256": return new AES256PrivacyProvider(new OctetString(s.PrivPassword), auth);
                case "AES192": return new AES192PrivacyProvider(new OctetString(s.PrivPassword), auth);
                default:       return new AESPrivacyProvider(new OctetString(s.PrivPassword), auth); // AES-128
            }
        }

        private static SnmpResult Error(string msg)
            => new SnmpResult { Success = false, ErrorMessage = msg };

        /// <summary>
        /// Fragt alle übergebenen Geräte per SNMP ab und trägt das Ergebnis in DeviceInfo.SnmpData ein.
        /// Geräte ohne Antwort bekommen SnmpData.Success = false.
        /// </summary>
        public void QueryDevices(IEnumerable<DeviceInfo> devices, SnmpSettings settings,
            Action<string> progress = null)
        {
            foreach (var device in devices)
            {
                progress?.Invoke($"SNMP → {device.IP}");
                try
                {
                    device.SnmpData = Query(device.IP, settings);
                    // Hostname aus SNMP übernehmen falls noch keiner bekannt
                    if (device.SnmpData.Success &&
                        string.IsNullOrEmpty(device.Hostname) &&
                        !string.IsNullOrEmpty(device.SnmpData.SysName))
                        device.Hostname = device.SnmpData.SysName;
                }
                catch
                {
                    device.SnmpData = new SnmpResult { Success = false, ErrorMessage = "Timeout / nicht erreichbar" };
                }
            }
        }
    }

    // =========================================================
    // === CREDENTIAL SCANNER ===
    // =========================================================
    public class CredentialScanner
    {
        public class ScanResult
        {
            public DatabaseDevice Device { get; set; }
            public CredentialTemplate MatchedTemplate { get; set; }
            public bool Success { get; set; }
            public string Message { get; set; }
            public DeviceIdentifier.IdentifyResult Identity { get; set; }
        }

        public event Action<ScanResult> DeviceScanned;
        public event Action<int, int> ProgressChanged; // current, total

        private readonly int _maxParallel;
        private CancellationTokenSource _cts;

        public CredentialScanner(int maxParallel = 5)
        {
            _maxParallel = maxParallel;
        }

        public void Cancel() => _cts?.Cancel();

        public async Task<List<ScanResult>> ScanDevicesAsync(
            List<DatabaseDevice> devices,
            List<CredentialTemplate> templates,
            bool retestExisting)
        {
            _cts = new CancellationTokenSource();
            var results = new List<ScanResult>();
            var semaphore = new SemaphoreSlim(_maxParallel);
            int progress = 0;
            int total = devices.Count;

            var tasks = devices.Select(async device =>
            {
                if (_cts.Token.IsCancellationRequested) return;

                await semaphore.WaitAsync(_cts.Token);
                try
                {
                    var result = await Task.Run(() => TestDevice(device, templates, retestExisting), _cts.Token);
                    lock (results) results.Add(result);
                    DeviceScanned?.Invoke(result);
                    var p = Interlocked.Increment(ref progress);
                    ProgressChanged?.Invoke(p, total);
                }
                finally
                {
                    semaphore.Release();
                }
            });

            try { await Task.WhenAll(tasks); }
            catch (OperationCanceledException) { }

            return results;
        }

        // Alle Protokolle die pro Template durchprobiert werden
        private static readonly (string Protocol, int Port)[] ALL_PROTOCOLS = {
            ("WMI",    135),
            ("SSH",     22),
            ("SNMP",   161),
            ("Telnet",  23),
            ("HTTP",    80),
            ("HTTPS",  443)
        };

        private ScanResult TestDevice(DatabaseDevice device, List<CredentialTemplate> templates, bool retestExisting)
        {
            if (!retestExisting && device.CredentialTemplateID.HasValue)
                return new ScanResult { Device = device, Success = false, Message = "Übersprungen (bereits getestet)" };

            foreach (var template in templates)
            {
                try
                {
                    string password = CredentialStore.Decrypt(template.EncryptedPass);

                    foreach (var (protocol, port) in ALL_PROTOCOLS)
                    {
                        try
                        {
                            bool success = TestProtocol(device.IP, protocol, template.Username, password, port);
                            if (success)
                            {
                                // UUID abrufen nach erfolgreichem Login
                                DeviceIdentifier.IdentifyResult identity = null;
                                try
                                {
                                    var identifier = new DeviceIdentifier();
                                    if (protocol == "WMI")
                                        identity = identifier.IdentifyWindows(device.IP, template.Username, password);
                                    else if (protocol == "SSH")
                                        identity = identifier.IdentifyViaSsh(device.IP, template.Username, password, port);
                                }
                                catch { }

                                string msg = $"{protocol} {template.Username}:*** → Erfolg";
                                if (identity != null && !string.IsNullOrEmpty(identity.UniqueID))
                                    msg += $" [ID: {identity.Source}]";

                                return new ScanResult
                                {
                                    Device = device,
                                    MatchedTemplate = template,
                                    Success = true,
                                    Message = msg,
                                    Identity = identity
                                };
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return new ScanResult { Device = device, Success = false, Message = "Keine Vorlage passte" };
        }

        private bool TestProtocol(string ip, string protocol, string username, string password, int port)
        {
            switch (protocol)
            {
                case "WMI":    return TestWmi(ip, username, password);
                case "SSH":    return TestSsh(ip, username, password, port);
                case "SNMP":   return TestSnmp(ip, password, port);
                case "Telnet": return TestTelnet(ip, username, password, port);
                case "HTTP":   return TestHttp(ip, username, password, port, false);
                case "HTTPS":  return TestHttp(ip, username, password, port, true);
                default:       return false;
            }
        }

        private bool TestWmi(string ip, string username, string password)
        {
            try
            {
                var options = HardwareManager.BuildConnectionOptions(username, password);
                var scope = new System.Management.ManagementScope($@"\\{ip}\root\cimv2", options)
                {
                    Options = { Timeout = TimeSpan.FromSeconds(10) }
                };
                scope.Connect();
                return scope.IsConnected;
            }
            catch { return false; }
        }

        private bool TestSsh(string ip, string username, string password, int port)
        {
            try
            {
                using (var client = new Renci.SshNet.SshClient(
                    new Renci.SshNet.ConnectionInfo(ip, port, username,
                        new Renci.SshNet.PasswordAuthenticationMethod(username, password))
                    { Timeout = TimeSpan.FromSeconds(10) }))
                {
                    client.Connect();
                    bool connected = client.IsConnected;
                    client.Disconnect();
                    return connected;
                }
            }
            catch { return false; }
        }

        private bool TestSnmp(string ip, string community, int port)
        {
            try
            {
                using (var udp = new UdpClient())
                {
                    udp.Client.ReceiveTimeout = 5000;
                    udp.Client.SendTimeout = 5000;

                    // SNMP v2c GET sysDescr.0 (1.3.6.1.2.1.1.1.0)
                    byte[] getRequest = BuildSnmpGetRequest(community, "1.3.6.1.2.1.1.1.0");
                    udp.Send(getRequest, getRequest.Length, new IPEndPoint(IPAddress.Parse(ip), port));

                    var remoteEP = new IPEndPoint(IPAddress.Any, 0);
                    byte[] response = udp.Receive(ref remoteEP);
                    return response != null && response.Length > 0;
                }
            }
            catch { return false; }
        }

        private byte[] BuildSnmpGetRequest(string community, string oid)
        {
            var communityBytes = System.Text.Encoding.ASCII.GetBytes(community);
            var oidParts = oid.Split('.').Select(s => int.Parse(s)).ToArray();

            var oidBytes = new List<byte>();
            oidBytes.Add((byte)(oidParts[0] * 40 + oidParts[1]));
            for (int i = 2; i < oidParts.Length; i++)
            {
                int val = oidParts[i];
                if (val < 128) oidBytes.Add((byte)val);
                else
                {
                    var chunks = new List<byte>();
                    while (val > 0) { chunks.Insert(0, (byte)(val & 0x7F)); val >>= 7; }
                    for (int j = 0; j < chunks.Count - 1; j++) chunks[j] |= 0x80;
                    oidBytes.AddRange(chunks);
                }
            }

            byte[] oidTlv = TlvEncode(0x06, oidBytes.ToArray());
            byte[] nullTlv = new byte[] { 0x05, 0x00 };
            byte[] varBind = TlvEncode(0x30, Combine(oidTlv, nullTlv));
            byte[] varBindList = TlvEncode(0x30, varBind);

            byte[] requestId = TlvEncode(0x02, new byte[] { 0x01 });
            byte[] errorStatus = TlvEncode(0x02, new byte[] { 0x00 });
            byte[] errorIndex = TlvEncode(0x02, new byte[] { 0x00 });
            byte[] pdu = TlvEncode(0xA0, Combine(requestId, errorStatus, errorIndex, varBindList));

            byte[] version = TlvEncode(0x02, new byte[] { 0x01 });
            byte[] comm = TlvEncode(0x04, communityBytes);
            byte[] message = TlvEncode(0x30, Combine(version, comm, pdu));

            return message;
        }

        private static byte[] TlvEncode(byte tag, byte[] value)
        {
            var result = new List<byte> { tag };
            if (value.Length < 128) result.Add((byte)value.Length);
            else
            {
                int lenBytes = value.Length < 256 ? 1 : 2;
                result.Add((byte)(0x80 | lenBytes));
                if (lenBytes == 2) result.Add((byte)(value.Length >> 8));
                result.Add((byte)(value.Length & 0xFF));
            }
            result.AddRange(value);
            return result.ToArray();
        }

        private static byte[] Combine(params byte[][] arrays)
        {
            int len = arrays.Sum(a => a.Length);
            var result = new byte[len];
            int offset = 0;
            foreach (var a in arrays) { Buffer.BlockCopy(a, 0, result, offset, a.Length); offset += a.Length; }
            return result;
        }

        private bool TestTelnet(string ip, string username, string password, int port)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    client.ReceiveTimeout = 5000;
                    client.SendTimeout = 5000;
                    var connectTask = client.ConnectAsync(ip, port);
                    if (!connectTask.Wait(5000)) return false;

                    using (var stream = client.GetStream())
                    {
                        byte[] buffer = new byte[4096];
                        stream.ReadTimeout = 5000;
                        int bytes = stream.Read(buffer, 0, buffer.Length);
                        string prompt = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();

                        if (prompt.Contains("login") || prompt.Contains("username") || prompt.Contains("user"))
                        {
                            byte[] userBytes = System.Text.Encoding.ASCII.GetBytes(username + "\r\n");
                            stream.Write(userBytes, 0, userBytes.Length);
                            System.Threading.Thread.Sleep(1000);
                            bytes = stream.Read(buffer, 0, buffer.Length);
                            prompt = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();
                        }

                        if (prompt.Contains("pass"))
                        {
                            byte[] passBytes = System.Text.Encoding.ASCII.GetBytes(password + "\r\n");
                            stream.Write(passBytes, 0, passBytes.Length);
                            System.Threading.Thread.Sleep(1000);
                            bytes = stream.Read(buffer, 0, buffer.Length);
                            string response = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();

                            return !response.Contains("fail") && !response.Contains("denied") &&
                                   !response.Contains("incorrect") && !response.Contains("invalid") &&
                                   (response.Contains("#") || response.Contains(">") || response.Contains("$"));
                        }
                        return false;
                    }
                }
            }
            catch { return false; }
        }

        private bool TestHttp(string ip, string username, string password, int port, bool https)
        {
            try
            {
                string scheme = https ? "https" : "http";
                var request = (HttpWebRequest)WebRequest.Create($"{scheme}://{ip}:{port}/");
                request.Timeout = 5000;
                request.Credentials = new NetworkCredential(username, password);
                request.PreAuthenticate = true;

                string auth = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{username}:{password}"));
                request.Headers.Add("Authorization", "Basic " + auth);

                if (https)
                    System.Net.ServicePointManager.ServerCertificateValidationCallback =
                        (sender, cert, chain, errors) => true;

                using (var response = (HttpWebResponse)request.GetResponse())
                    return response.StatusCode == HttpStatusCode.OK ||
                           response.StatusCode == HttpStatusCode.Found ||
                           response.StatusCode == HttpStatusCode.Redirect;
            }
            catch (WebException ex)
            {
                if (ex.Response is HttpWebResponse resp)
                    return resp.StatusCode != HttpStatusCode.Unauthorized &&
                           resp.StatusCode != HttpStatusCode.Forbidden;
                return false;
            }
            catch { return false; }
        }
    }

}