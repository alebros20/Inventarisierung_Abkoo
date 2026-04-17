using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace NmapInventory
{
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

        public string PairDevice(string ipPairPort, string pairCode)
        {
            string result = RunAdb($"pair {ipPairPort} {pairCode}", out string err);
            if (result.Contains("Successfully") || result.Contains("successfully"))
                return $"Kopplung erfolgreich: {result.Trim()}";
            string combined = (result + "\n" + err).Trim();
            throw new Exception($"Kopplung fehlgeschlagen:\n{combined}");
        }

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
                string connectResult = RunAdb($"connect {ipPort}", out string connectErr);
                hw.AppendLine($"  ADB Connect: {connectResult.Trim()}");

                if (connectResult.Contains("failed") || connectResult.Contains("refused") || connectResult.Contains("unable"))
                    throw new Exception($"Verbindung fehlgeschlagen: {connectResult.Trim()}");

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
}
