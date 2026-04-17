using System;
using System.Diagnostics;
using System.Management;

namespace NmapInventory
{
    public class HardwareManager
    {
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

        public string GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            try
            {
                string tempFile = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), $"msinfo_{Guid.NewGuid():N}.txt");

                var psi = new ProcessStartInfo
                {
                    FileName = "msinfo32.exe",
                    Arguments = $"/computer {computerIP} /report \"{tempFile}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var proc = Process.Start(psi))
                {
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

                string content = System.IO.File.ReadAllText(tempFile, System.Text.Encoding.Unicode);
                System.IO.File.Delete(tempFile);

                if (string.IsNullOrWhiteSpace(content))
                    return FallbackRemoteWmi(computerIP, username, password);

                return content;
            }
            catch (Exception)
            {
                return FallbackRemoteWmi(computerIP, username, password);
            }
        }

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
                    long.TryParse(Val(o, "TotalPhysicalMemory"), out long ram2);
                    sb.AppendLine($"Physischer RAM   : {ram2 / (1024 * 1024 * 1024)} GB");
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
}
