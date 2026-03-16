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
        // 🔬 DETAIL SCAN — Ports, Services, OS, Banner
        // Läuft auf bekannten IPs aus dem Discovery-Scan
        // ─────────────────────────────────────────────────────────
        public ScanResult DetailScan(string targets)
        {
            // -sV       Service-Versionen
            // -O        OS-Erkennung (braucht Admin)
            // -sC       Standard-Skripte (Banner, SMB, HTTP-Titel...)
            // -T4       schneller Timing
            // --open    nur offene Ports
            return RunScan(targets, "-sV -O -sC -T4 --open", ParseDetailOutput);
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
                    hostname = !string.IsNullOrEmpty(mac) ? $"{vendor} ({mac})" : $"{vendor} ({ip})";
                else if (!string.IsNullOrEmpty(dnsName))
                    hostname = dnsName;
                else
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
        // Parser für Detail-Scan (Ports, Services, OS, Banner)
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

                foreach (var line in lines)
                {
                    if (line.StartsWith("Nmap scan report for"))
                    {
                        var m = Regex.Match(line, @"for\s+(\S+)\s+\((\d+\.\d+\.\d+\.\d+)\)");
                        if (m.Success) { device.Hostname = m.Groups[1].Value; device.IP = m.Groups[2].Value; }
                        else { var m2 = Regex.Match(line, @"(\d+\.\d+\.\d+\.\d+)"); if (m2.Success) device.IP = m2.Groups[1].Value; }
                    }
                    else if (line.Contains("MAC Address:"))
                    {
                        var m = Regex.Match(line,
                            @"MAC Address:\s+([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})\s*(?:\(([^)]*)\))?");
                        if (m.Success) { device.MacAddress = m.Groups[1].Value; device.Vendor = m.Groups[2].Value; }
                    }
                    else if (Regex.IsMatch(line, @"^\d+/(tcp|udp)\s+\w+"))
                    {
                        var m = Regex.Match(line, @"^(\d+)/(tcp|udp)\s+(\S+)\s+(\S+)\s*(.*)$");
                        if (m.Success)
                            device.OpenPorts.Add(new NmapPort
                            {
                                Port = int.Parse(m.Groups[1].Value),
                                Protocol = m.Groups[2].Value,
                                State = m.Groups[3].Value,
                                Service = m.Groups[4].Value,
                                Version = m.Groups[5].Value.Trim()
                            });
                    }
                    else if (line.TrimStart().StartsWith("| banner:") && device.OpenPorts.Count > 0)
                        device.OpenPorts[device.OpenPorts.Count - 1].Banner =
                            line.Substring(line.IndexOf("| banner:") + 9).Trim();
                    else if (line.StartsWith("OS details:"))
                        device.OSDetails = line.Substring(11).Trim();
                    else if (line.StartsWith("Running:"))
                        device.OS = line.Substring(8).Trim();
                    else if (line.StartsWith("Aggressive OS guesses:") && string.IsNullOrEmpty(device.OS))
                        device.OS = line.Substring(22).Trim();
                }

                if (!string.IsNullOrEmpty(device.IP))
                {
                    if (string.IsNullOrEmpty(device.Hostname)) device.Hostname = device.IP;
                    device.Ports = device.OpenPorts.Count > 0
                        ? string.Join(", ", device.OpenPorts.Select(p => $"{p.Port}/{p.Protocol} {p.Service}"))
                        : "-";
                    devices.Add(device);
                }
            }
            return devices;
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