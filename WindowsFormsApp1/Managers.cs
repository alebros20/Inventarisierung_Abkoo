using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NmapInventory
{
    public class NmapScanner
    {
        public NmapScanResult Scan(string network)
        {
            string nmapPath = FindNmapPath();
            if (nmapPath == null) throw new Exception("Nmap nicht gefunden");

            var psi = new ProcessStartInfo
            {
                FileName = nmapPath,
                Arguments = $"-p 22,80,139,443,445,3306,3389,53,135,515,6668,8009,8008,8080,8443,9000,9100,9220,9999,10025 -T4 --min-hostgroup 256 -oG - {network}",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            string output;
            using (var p = Process.Start(psi))
            {
                output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
            }

            return new NmapScanResult { RawOutput = output, Devices = ParseNmapOutput(output) };
        }

        private List<DeviceInfo> ParseNmapOutput(string output)
        {
            var devices = new Dictionary<string, DeviceInfo>();
            foreach (string line in output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
            {
                if (line.StartsWith("Host:"))
                {
                    var parts = line.Split('\t');
                    if (parts.Length < 2) continue;
                    var hostPart = parts[0].Replace("Host: ", "").Trim();
                    var ip = hostPart.Split(' ')[0];
                    if (!devices.ContainsKey(ip)) devices[ip] = new DeviceInfo { IP = ip };
                    if (hostPart.Contains("(")) devices[ip].Hostname = hostPart.Substring(hostPart.IndexOf('(') + 1, hostPart.IndexOf(')') - hostPart.IndexOf('(') - 1).Trim();
                    int statusIdx = line.IndexOf("Status:");
                    devices[ip].Status = statusIdx >= 0 ? line.Substring(statusIdx + 7).Trim().Split(' ')[0] : "Up";
                    var portList = new List<string>();
                    foreach (var part in parts)
                    {
                        if (part.Contains("Ports:"))
                        {
                            foreach (var portEntry in part.Replace("Ports:", "").Trim().Split(','))
                            {
                                var pp = portEntry.Trim().Split('/');
                                if (pp.Length >= 3 && pp[1].ToLower() == "open")
                                {
                                    string formatted = $"{pp[0]}/{pp[2]}";
                                    if (pp.Length > 4 && !string.IsNullOrEmpty(pp[4])) formatted += $" ({pp[4]})";
                                    portList.Add(formatted);
                                }
                            }
                            break;
                        }
                    }
                    devices[ip].Ports = portList.Count > 0 ? string.Join(", ", portList) : "-";
                }
            }
            return devices.Values.ToList();
        }

        private string FindNmapPath() => new[] { @"C:\Program Files\Nmap\nmap.exe", @"C:\Program Files (x86)\Nmap\nmap.exe", "nmap" }.FirstOrDefault(File.Exists);
    }

    public class HardwareManager
    {
        public string GetHardwareInfo()
        {
            var sb = new StringBuilder();
            AppendSystemInfo(sb, Environment.MachineName, null);
            return sb.ToString();
        }

        public string GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== REMOTE HARDWARE: {computerIP} ===\n");
            try
            {
                var scope = CreateManagementScope(computerIP, username, password);
                scope.Connect();
                AppendSystemInfo(sb, computerIP, scope);
            }
            catch (Exception ex) 
            { 
                sb.AppendLine($"\nFehler: {ex.Message}");
                if (ex.InnerException != null)
                    sb.AppendLine($"Details: {ex.InnerException.Message}");
            }
            return sb.ToString();
        }

        private ManagementScope CreateManagementScope(string computerIP, string username, string password)
        {
            var options = new ConnectionOptions 
            { 
                Authentication = AuthenticationLevel.PacketPrivacy,
                Timeout = TimeSpan.FromSeconds(30)
            };
            
            if (!string.IsNullOrEmpty(username))
            {
                options.Username = username;
                options.Password = password;
            }
            
            return new ManagementScope($@"\\{computerIP}\root\cimv2", options);
        }

        private void AppendSystemInfo(StringBuilder sb, string computerIP, ManagementScope scope)
        {
            sb.AppendLine("=== BETRIEBSSYSTEM ===");
            sb.AppendLine("Computer: " + computerIP);
            QueryWMI(sb, scope, "SELECT Caption, Version, OSArchitecture FROM Win32_OperatingSystem", "OS");
            
            sb.AppendLine("\n=== SYSTEM ===");
            QueryWMI(sb, scope, "SELECT Manufacturer, Model FROM Win32_ComputerSystem", "System");
            
            sb.AppendLine("\n=== PROZESSOR ===");
            QueryWMI(sb, scope, "SELECT Name, Manufacturer, MaxClockSpeed, NumberOfCores FROM Win32_Processor", "CPU");
            
            sb.AppendLine("\n=== ARBEITSSPEICHER ===");
            QueryWMI(sb, scope, "SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem", "Memory");
            
            sb.AppendLine("\n=== RAM-MODULE ===");
            QueryWMI(sb, scope, "SELECT Manufacturer, Capacity, Speed FROM Win32_PhysicalMemory", "RAM");
            
            sb.AppendLine("\n=== FESTPLATTEN ===");
            QueryWMI(sb, scope, "SELECT DeviceID, Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType=3", "Disk");
        }

        private void QueryWMI(StringBuilder sb, ManagementScope scope, string query, string type)
        {
            try
            {
                var searcher = scope == null 
                    ? new ManagementObjectSearcher(query) 
                    : new ManagementObjectSearcher(scope, new ObjectQuery(query));
                    
                foreach (ManagementObject obj in searcher.Get())
                {
                    switch (type)
                    {
                        case "OS":
                            sb.AppendLine("OS: " + obj["Caption"]);
                            sb.AppendLine("Version: " + obj["Version"]);
                            break;
                        case "System":
                            sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                            sb.AppendLine("Modell: " + obj["Model"]);
                            break;
                        case "CPU":
                            sb.AppendLine("Name: " + obj["Name"]);
                            sb.AppendLine("Taktfrequenz: " + obj["MaxClockSpeed"] + " MHz");
                            sb.AppendLine("Kerne: " + obj["NumberOfCores"]);
                            break;
                        case "Memory":
                            long totalMB = Convert.ToInt64(obj["TotalVisibleMemorySize"]) / 1024;
                            long freeMB = Convert.ToInt64(obj["FreePhysicalMemory"]) / 1024;
                            sb.AppendLine($"Gesamt: {totalMB} MB");
                            sb.AppendLine($"Belegt: {totalMB - freeMB} MB");
                            break;
                        case "RAM":
                            long capacityGB = Convert.ToInt64(obj["Capacity"]) / (1024 * 1024 * 1024);
                            sb.AppendLine($"Kapazität: {capacityGB} GB, Geschwindigkeit: {obj["Speed"]} MHz");
                            break;
                        case "Disk":
                            long sizeGB = Convert.ToInt64(obj["Size"]) / (1024 * 1024 * 1024);
                            sb.AppendLine($"{obj["DeviceID"]}: {sizeGB} GB");
                            break;
                    }
                }
            }
            catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
        }
    }

    public class SoftwareManager
    {
        public List<SoftwareInfo> GetInstalledSoftware()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-Command \"winget list --accept-source-agreements | Out-String\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    return ParseWingetOutput(output);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei winget list: {ex.Message}\n\nFallback auf Registry...", "Warnung");
                var fallbackSoftware = new Dictionary<string, SoftwareInfo>();
                ReadRegistry(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", fallbackSoftware);
                ReadRegistry(Registry.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", fallbackSoftware);
                return fallbackSoftware.Values.ToList();
            }
        }

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("PowerShell Remoting erfordert Benutzername und Passwort.\n\nVerwende WMI Registry-Abfrage...", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return GetRemoteSoftwareViaWMI(computerIP, username, password);
            }
            
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $@"-Command ""$pass = ConvertTo-SecureString '{password}' -AsPlainText -Force; $cred = New-Object System.Management.Automation.PSCredential ('{username}', $pass); Invoke-Command -ComputerName {computerIP} -Credential $cred -ScriptBlock {{ winget list --accept-source-agreements }} -ErrorAction Stop | Out-String""",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    return process.ExitCode == 0 && !string.IsNullOrEmpty(output) ? ParseWingetOutput(output) : GetRemoteSoftwareViaWMI(computerIP, username, password);
                }
            }
            catch
            {
                return GetRemoteSoftwareViaWMI(computerIP, username, password);
            }
        }

        private List<SoftwareInfo> ParseWingetOutput(string output)
        {
            var software = new List<SoftwareInfo>();
            var lines = output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            bool dataStarted = false;
            
            foreach (var line in lines)
            {
                if (line.Contains("---")) { dataStarted = true; continue; }
                if (!dataStarted || line.Trim().Length < 10) continue;
                
                var parts = Regex.Split(line, @"\s{2,}");
                if (parts.Length >= 3 && !parts[0].StartsWith("Name"))
                {
                    software.Add(new SoftwareInfo
                    {
                        Name = parts[0].Trim(),
                        Version = parts.Length > 2 ? parts[2].Trim() : "",
                        Publisher = parts.Length > 1 ? parts[1].Trim() : "",
                        Source = "winget",
                        LastUpdate = "Aktuell"
                    });
                }
            }
            return software;
        }

        private List<SoftwareInfo> GetRemoteSoftwareViaWMI(string computerIP, string username, string password)
        {
            var software = new Dictionary<string, SoftwareInfo>();
            try
            {
                var options = new ConnectionOptions 
                { 
                    Authentication = AuthenticationLevel.PacketPrivacy,
                    Timeout = TimeSpan.FromSeconds(60)
                };
                
                if (!string.IsNullOrEmpty(username))
                {
                    options.Username = username;
                    options.Password = password;
                }
                
                var scope = new ManagementScope($@"\\{computerIP}\root\default", options);
                scope.Connect();
                var registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);
                
                foreach (string regPath in new[] { @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" })
                {
                    var inParams = registry.GetMethodParameters("EnumKey");
                    inParams["hDefKey"] = 0x80000002;
                    inParams["sSubKeyName"] = regPath;
                    var outParams = registry.InvokeMethod("EnumKey", inParams, null);
                    
                    if (outParams["sNames"] != null)
                    {
                        foreach (string subkey in (string[])outParams["sNames"])
                        {
                            string fullPath = regPath + @"\" + subkey;
                            string displayName = GetRegistryValue(registry, fullPath, "DisplayName");
                            if (!string.IsNullOrEmpty(displayName) && !displayName.Contains("Update"))
                            {
                                var uniqueKey = displayName + "|" + GetRegistryValue(registry, fullPath, "DisplayVersion");
                                if (!software.ContainsKey(uniqueKey))
                                    software[uniqueKey] = new SoftwareInfo 
                                    { 
                                        Name = displayName, 
                                        Version = GetRegistryValue(registry, fullPath, "DisplayVersion") ?? "N/A", 
                                        Publisher = GetRegistryValue(registry, fullPath, "Publisher") ?? "",
                                        Source = "Remote Registry",
                                        LastUpdate = "-"
                                    };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Abrufen der Remote-Software:\n{ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return software.Values.ToList();
        }

        private string GetRegistryValue(ManagementClass registry, string keyPath, string valueName)
        {
            try
            {
                var inParams = registry.GetMethodParameters("GetStringValue");
                inParams["hDefKey"] = 0x80000002;
                inParams["sSubKeyName"] = keyPath;
                inParams["sValueName"] = valueName;
                var outParams = registry.InvokeMethod("GetStringValue", inParams, null);
                return outParams["sValue"]?.ToString();
            }
            catch { return null; }
        }

        public void UpdateSoftware(string softwareName, string remotePC, Label statusLabel)
        {
            Task.Run(() =>
            {
                try
                {
                    var psi = new ProcessStartInfo { FileName = "powershell.exe", Arguments = $"-Command \"winget upgrade --id '{softwareName}' --accept-source-agreements\"", RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                    using (var p = Process.Start(psi))
                    {
                        p.WaitForExit();
                        statusLabel.Invoke(new Action(() => statusLabel.Text = p.ExitCode == 0 ? "Update erfolgreich" : "Update fehlgeschlagen"));
                    }
                }
                catch (Exception ex) { statusLabel.Invoke(new Action(() => MessageBox.Show($"Fehler: {ex.Message}"))); }
            });
        }

        private void ReadRegistry(RegistryKey rootKey, string path, Dictionary<string, SoftwareInfo> software)
        {
            try
            {
                using (var key = rootKey.OpenSubKey(path))
                {
                    if (key == null) return;
                    foreach (string subkey in key.GetSubKeyNames())
                    {
                        using (var sub = key.OpenSubKey(subkey))
                        {
                            var name = sub?.GetValue("DisplayName")?.ToString();
                            if (string.IsNullOrEmpty(name) || name.Contains("Update")) continue;
                            var version = sub?.GetValue("DisplayVersion")?.ToString() ?? "N/A";
                            var uniqueKey = name + "|" + version;
                            if (!software.ContainsKey(uniqueKey))
                                software[uniqueKey] = new SoftwareInfo 
                                { 
                                    Name = name, 
                                    Version = version, 
                                    Publisher = sub?.GetValue("Publisher")?.ToString() ?? "", 
                                    Source = "Registry",
                                    LastUpdate = "-"
                                };
                        }
                    }
                }
            }
            catch { }
        }
    }
}
