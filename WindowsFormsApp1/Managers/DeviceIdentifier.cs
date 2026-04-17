using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;
using Renci.SshNet;

namespace NmapInventory
{
    public class DeviceIdentifier
    {
        public class IdentifyResult
        {
            public string UniqueID { get; set; }
            public string Source { get; set; }
            public List<(string Mac, string IP, string Type)> Interfaces { get; set; } = new List<(string, string, string)>();
        }

        private static readonly string[] INVALID_UUIDS = {
            "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF",
            "00000000-0000-0000-0000-000000000000",
            "03000200-0400-0500-0006-000700080009"
        };

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
                        inParams["hDefKey"] = 0x80000002;
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

        public IdentifyResult IdentifyViaSsh(string ip, string username, string password, int port = 22)
        {
            var result = new IdentifyResult();
            try
            {
                using (var client = new SshClient(
                    new ConnectionInfo(ip, port, username,
                        new PasswordAuthenticationMethod(username, password))
                    { Timeout = TimeSpan.FromSeconds(15) }))
                {
                    client.Connect();
                    if (!client.IsConnected) return result;

                    string uname = RunSshCommand(client, "uname -s")?.Trim().ToLower() ?? "";

                    if (uname == "darwin")
                    {
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
                        string machineId = RunSshCommand(client, "cat /etc/machine-id 2>/dev/null")?.Trim();
                        if (!string.IsNullOrEmpty(machineId) && machineId.Length >= 32)
                        {
                            result.UniqueID = machineId.ToUpperInvariant();
                            result.Source = "machine-id";
                        }

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

                    string ifOutput = RunSshCommand(client, "ip -o link show 2>/dev/null || ifconfig -a 2>/dev/null");
                    if (!string.IsNullOrEmpty(ifOutput))
                    {
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
}
