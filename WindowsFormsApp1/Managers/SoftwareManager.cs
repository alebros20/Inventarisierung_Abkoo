using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Windows.Forms;

namespace NmapInventory
{
    public class SoftwareManager
    {
        private static readonly string[] REG_PATHS = {
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        };

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
                            catch { }
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

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            var software = new List<SoftwareInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var options = HardwareManager.BuildConnectionOptions(username, password);
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
                            if (!seen.Add(name)) continue;

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
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                software.Add(new SoftwareInfo { Name = $"Fehler: {ex.Message}" });
            }

            return software.OrderBy(s => s.Name).ToList();
        }

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
