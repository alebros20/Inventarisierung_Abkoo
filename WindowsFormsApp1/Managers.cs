using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace NmapInventory
{
    public class NmapScanner
    {
        private string nmapPath = "nmap";

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
                string hostname;
                if (!string.IsNullOrEmpty(vendor) && vendor != "Unknown")
                    // Hersteller bekannt → "Arcadyan (78:DD:12:D2:10:3E)"
                    hostname = !string.IsNullOrEmpty(mac) ? $"{vendor} ({mac})" : vendor;
                else if (!string.IsNullOrEmpty(dnsName))
                    // DNS-Name bekannt → "speedport.ip"
                    hostname = dnsName;
                else if (!string.IsNullOrEmpty(mac))
                    // Nur MAC bekannt → MAC als Hostname
                    hostname = mac;
                else
                    // Nichts bekannt → IP als letzter Fallback
                    hostname = ip;

                devices.Add(new DeviceInfo
                {
                    IP = ip,
                    Hostname = hostname,
                    MacAddress = mac,
                    Vendor = vendor,
                    Status = "up",
                    Ports = "-",
                    OpenPorts = new List<NmapPort>()
                });
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
        public string GetHardwareInfo()
        {
            var info = new System.Text.StringBuilder();
            try
            {
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem").Get())
                { info.AppendLine($"OS: {obj["Caption"]}"); info.AppendLine($"Version: {obj["Version"]}"); info.AppendLine($"Build: {obj["BuildNumber"]}"); }
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_Processor").Get())
                { info.AppendLine($"CPU: {obj["Name"]}"); info.AppendLine($"Cores: {obj["NumberOfCores"]}"); info.AppendLine($"Threads: {obj["NumberOfLogicalProcessors"]}"); }
                long totalRAM = 0;
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory").Get())
                    totalRAM += long.Parse(obj["Capacity"].ToString());
                info.AppendLine($"RAM: {totalRAM / (1024 * 1024 * 1024)} GB");
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3").Get())
                { string d = obj["Name"].ToString(); long s = long.Parse(obj["Size"].ToString()); long f = long.Parse(obj["FreeSpace"].ToString()); info.AppendLine($"Disk {d}: {s / (1024 * 1024 * 1024)} GB (Free: {f / (1024 * 1024 * 1024)} GB)"); }
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true").Get())
                { var ips = obj["IPAddress"] as string[]; if (ips?.Length > 0) info.AppendLine($"IP: {ips[0]}"); }
            }
            catch (Exception ex) { info.AppendLine($"Fehler: {ex.Message}"); }
            return info.ToString();
        }

        public string GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            var info = new System.Text.StringBuilder();
            try
            {
                var options = new ConnectionOptions { Authentication = AuthenticationLevel.PacketPrivacy };
                if (!string.IsNullOrEmpty(username)) options.Impersonation = ImpersonationLevel.Impersonate;
                var scope = new ManagementScope($"\\\\{computerIP}\\root\\cimv2", options);
                scope.Connect();
                foreach (ManagementObject obj in new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_OperatingSystem")).Get())
                    info.AppendLine($"OS: {obj["Caption"]}");
                foreach (ManagementObject obj in new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_Processor")).Get())
                    info.AppendLine($"CPU: {obj["Name"]}");
            }
            catch (Exception ex) { info.AppendLine($"Fehler: {ex.Message}"); }
            return info.ToString();
        }
    }

    public class SoftwareManager
    {
        public List<SoftwareInfo> GetInstalledSoftware()
        {
            var software = new List<SoftwareInfo>();
            try
            {
                foreach (ManagementObject obj in new ManagementObjectSearcher("SELECT * FROM Win32_Product").Get())
                    software.Add(new SoftwareInfo { Name = obj["Name"]?.ToString() ?? "", Version = obj["Version"]?.ToString() ?? "", Publisher = obj["Vendor"]?.ToString() ?? "", InstallDate = obj["InstallDate"]?.ToString() ?? "" });
            }
            catch (Exception ex) { software.Add(new SoftwareInfo { Name = $"Fehler: {ex.Message}" }); }
            return software;
        }

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            var software = new List<SoftwareInfo>();
            try
            {
                var options = new ConnectionOptions { Authentication = AuthenticationLevel.PacketPrivacy };
                if (!string.IsNullOrEmpty(username)) options.Impersonation = ImpersonationLevel.Impersonate;
                var scope = new ManagementScope($"\\\\{computerIP}\\root\\cimv2", options);
                scope.Connect();
                foreach (ManagementObject obj in new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_Product")).Get())
                    software.Add(new SoftwareInfo { Name = obj["Name"]?.ToString() ?? "", Version = obj["Version"]?.ToString() ?? "", Publisher = obj["Vendor"]?.ToString() ?? "", InstallDate = obj["InstallDate"]?.ToString() ?? "" });
            }
            catch (Exception ex) { software.Add(new SoftwareInfo { Name = $"Fehler: {ex.Message}" }); }
            return software;
        }

        public void UpdateSoftware(string softwareName, string computerIP, Label statusLabel)
        {
            statusLabel.Text = $"Aktualisiere {softwareName}...";
            MessageBox.Show($"Update-Funktion für '{softwareName}' ist nicht implementiert.", "Info");
            statusLabel.Text = "Bereit";
        }
    }
}