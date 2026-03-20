using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

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
                var options = BuildConnectionOptions(username, password);
                var scope = new ManagementScope($"\\\\{computerIP}\\root\\cimv2", options);
                scope.Connect();

                foreach (ManagementObject obj in new ManagementObjectSearcher(scope,
                    new ObjectQuery("SELECT * FROM Win32_OperatingSystem")).Get())
                {
                    info.AppendLine($"OS:      {obj["Caption"]}");
                    info.AppendLine($"Version: {obj["Version"]}");
                    info.AppendLine($"Build:   {obj["BuildNumber"]}");
                }
                foreach (ManagementObject obj in new ManagementObjectSearcher(scope,
                    new ObjectQuery("SELECT * FROM Win32_Processor")).Get())
                {
                    info.AppendLine($"CPU:     {obj["Name"]}");
                    info.AppendLine($"Kerne:   {obj["NumberOfCores"]}");
                    info.AppendLine($"Threads: {obj["NumberOfLogicalProcessors"]}");
                }
                long ram = 0;
                foreach (ManagementObject obj in new ManagementObjectSearcher(scope,
                    new ObjectQuery("SELECT * FROM Win32_PhysicalMemory")).Get())
                    ram += long.Parse(obj["Capacity"].ToString());
                if (ram > 0) info.AppendLine($"RAM:     {ram / (1024 * 1024 * 1024)} GB");

                foreach (ManagementObject obj in new ManagementObjectSearcher(scope,
                    new ObjectQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3")).Get())
                {
                    long size = long.Parse(obj["Size"].ToString());
                    long free = long.Parse(obj["FreeSpace"].ToString());
                    info.AppendLine($"Disk {obj["Name"]}: {size / (1024 * 1024 * 1024)} GB (Frei: {free / (1024 * 1024 * 1024)} GB)");
                }
            }
            catch (UnauthorizedAccessException)
            {
                info.AppendLine("FEHLER: Zugriff verweigert.");
                info.AppendLine("");
                info.AppendLine("Mögliche Ursachen:");
                info.AppendLine("1. Falsches Passwort / falscher Benutzername");
                info.AppendLine("   → Format: 'COMPUTERNAME\\Administrator' oder nur 'Administrator'");
                info.AppendLine("");
                info.AppendLine("2. UAC blockiert Remote-WMI (Windows 10/11)");
                info.AppendLine("   → Auf dem Zielgerät als Admin ausführen:");
                info.AppendLine("   reg add \"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System\"");
                info.AppendLine("       /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f");
                info.AppendLine("");
                info.AppendLine("3. Windows-Firewall blockiert WMI");
                info.AppendLine("   → Auf dem Zielgerät als Admin:");
                info.AppendLine("   netsh advfirewall firewall set rule");
                info.AppendLine("       group=\"Windows-Verwaltungsinstrumentation (WMI)\" new enable=yes");
            }
            catch (Exception ex)
            {
                info.AppendLine($"FEHLER: {ex.Message}");
            }
            return info.ToString();
        }

        // Erstellt ConnectionOptions mit Credentials — Username und Password werden korrekt gesetzt
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

}