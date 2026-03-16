using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace NmapInventory
{
    public class NmapScanner
    {
        private string nmapPath = "nmap";

        public ScanResult Scan(string target)
        {
            var devices = new List<DeviceInfo>();
            string rawOutput = "";

            try
            {
                var process = new ProcessStartInfo
                {
                    FileName = nmapPath,
                    Arguments = $"-sn -PR {target}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(process))
                {
                    using (var reader = proc.StandardOutput)
                        rawOutput = reader.ReadToEnd();
                    proc.WaitForExit();
                }

                devices = ParseNmapOutput(rawOutput);
            }
            catch (Exception ex)
            {
                rawOutput = $"Fehler: {ex.Message}";
            }

            return new ScanResult { Devices = devices, RawOutput = rawOutput };
        }

        private List<DeviceInfo> ParseNmapOutput(string output)
        {
            var devices = new List<DeviceInfo>();
            var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            // Aktueller Block
            string currentIP = "";
            string currentDnsName = "";
            string currentMac = "";
            string currentVendor = "";
            bool currentIsUp = false;

            // Hilfsmethode: fertigen Block in Liste übernehmen
            void FlushDevice()
            {
                if (!currentIsUp || string.IsNullOrEmpty(currentIP)) return;

                // Priorität: 1. Hersteller  2. DNS-Name  3. IP
                string hostname;
                if (!string.IsNullOrEmpty(currentVendor) && currentVendor != "Unknown")
                    hostname = !string.IsNullOrEmpty(currentMac)
                        ? $"{currentVendor} ({currentMac})"
                        : $"{currentVendor} ({currentIP})";
                else if (!string.IsNullOrEmpty(currentDnsName))
                    hostname = currentDnsName;
                else
                    hostname = currentIP;

                devices.Add(new DeviceInfo
                {
                    IP = currentIP,
                    Hostname = hostname,
                    Status = "up",
                    Ports = "-",
                    MacAddress = currentMac
                });
            }

            foreach (var line in lines)
            {
                // Neuer Block beginnt → vorherigen abschließen
                if (line.Contains("Nmap scan report for"))
                {
                    FlushDevice();

                    // Reset
                    currentIP = ""; currentDnsName = ""; currentMac = "";
                    currentVendor = ""; currentIsUp = false;

                    // "for speedport.ip (192.168.2.1)" oder "for 192.168.2.206"
                    var withHost = Regex.Match(line, @"for\s+(\S+)\s+\((\d+\.\d+\.\d+\.\d+)\)");
                    if (withHost.Success)
                    {
                        currentDnsName = withHost.Groups[1].Value;
                        currentIP = withHost.Groups[2].Value;
                    }
                    else
                    {
                        var ipOnly = Regex.Match(line, @"(\d+\.\d+\.\d+\.\d+)");
                        if (ipOnly.Success) currentIP = ipOnly.Groups[1].Value;
                    }
                }
                // "Host is up" — merken, noch nicht sofort speichern
                else if (line.Contains("Host is up"))
                {
                    currentIsUp = true;
                }
                // "MAC Address: xx:xx:xx (Vendor)" — kommt NACH "Host is up"
                else if (line.Contains("MAC Address:"))
                {
                    var macMatch = Regex.Match(line,
                        @"MAC Address:\s+([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})\s+\(([^)]+)\)");
                    if (macMatch.Success)
                    {
                        currentMac = macMatch.Groups[1].Value;
                        currentVendor = macMatch.Groups[2].Value;
                    }
                    else
                    {
                        var macOnly = Regex.Match(line, @"([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})");
                        if (macOnly.Success) currentMac = macOnly.Groups[1].Value;
                    }
                }
            }

            // Letzten Block nicht vergessen
            FlushDevice();

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