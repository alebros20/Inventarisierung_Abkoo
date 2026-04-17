using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace NmapInventory
{
    public class NmapScanner
    {
        private readonly string nmapPath;

        public NmapScanner()
        {
            nmapPath = FindNmap();
        }

        private static string FindNmap()
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string local = Path.Combine(appDir, "nmap.exe");
            if (File.Exists(local)) return local;

            string subDir = Path.Combine(appDir, "nmap", "nmap.exe");
            if (File.Exists(subDir)) return subDir;

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

            return "nmap";
        }

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

        public ScanResult DiscoveryScan(string target)
        {
            return RunScan(target, "-sn -PR --send-eth -T4", ParseDiscoveryOutput);
        }

        public ScanResult DetailScan(string targets)
        {
            const string REMOTE_PORTS =
                "22,23,80,443,445,3389,5900,5985,5986,8080,8443";
            const string SCRIPTS =
                "rdp-info,ssh-hostkey,vnc-info,smb-os-discovery,http-title";
            string args = $"-sV --version-light -O -T4 --open " +
                          $"-p {REMOTE_PORTS} --script={SCRIPTS} " +
                          $"--min-parallelism 10 --max-retries 1 --host-timeout 60s";
            return RunScan(targets, args, ParseDetailOutput);
        }

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

        public ScanResult Scan(string target) => DiscoveryScan(target);

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
                    Hostname   = !string.IsNullOrEmpty(dnsName) ? dnsName : "",
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
            Flush();
            return devices;
        }

        private List<DeviceInfo> ParseDetailOutput(string output)
        {
            var devices = new List<DeviceInfo>();
            var blocks = Regex.Split(output, @"(?=Nmap scan report for )");

            foreach (var block in blocks)
            {
                if (!block.Contains("Nmap scan report for")) continue;
                var device = new DeviceInfo { OpenPorts = new List<NmapPort>(), Status = "up" };
                var lines = block.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                NmapPort lastPort = null;

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
                    else if (line.TrimStart().StartsWith("|") && lastPort != null)
                    {
                        string scriptLine = line.TrimStart('|', '_', ' ', '-');
                        if (scriptLine.StartsWith("rdp-info:") || scriptLine.Contains("Remote Desktop Protocol"))
                            AppendBanner(lastPort, scriptLine);
                        else if (scriptLine.StartsWith("ssh-hostkey:") || scriptLine.Contains("ssh-rsa") || scriptLine.Contains("ecdsa"))
                            AppendBanner(lastPort, scriptLine.Trim());
                        else if (scriptLine.Contains("VNC") || scriptLine.StartsWith("vnc-info:"))
                            AppendBanner(lastPort, scriptLine);
                        else if (scriptLine.Contains("OS:") || scriptLine.Contains("Computer name:") ||
                                 scriptLine.Contains("Domain name:") || scriptLine.Contains("Workgroup:"))
                        {
                            AppendBanner(lastPort, scriptLine.Trim());
                            var osMatch = Regex.Match(scriptLine, @"OS:\s*(.+)");
                            if (osMatch.Success && string.IsNullOrEmpty(device.OS))
                                device.OS = osMatch.Groups[1].Value.Trim();
                        }
                        else if (scriptLine.StartsWith("http-title:") || scriptLine.Contains("Title:"))
                            AppendBanner(lastPort, scriptLine.Trim());
                        else if (scriptLine.StartsWith("banner:"))
                            AppendBanner(lastPort, scriptLine.Substring(7).Trim());
                    }
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
                    device.DeviceType = DeviceTypeHelper.Detect(device);
                    devices.Add(device);
                }
            }
            return devices;
        }

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
}
