using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace NmapInventory
{
    public class IosManager
    {
        private readonly string _ideviceinfoPath;
        public bool IsAvailable => _ideviceinfoPath != "ideviceinfo" ? File.Exists(_ideviceinfoPath) : IsInPath("ideviceinfo");

        public IosManager() { _ideviceinfoPath = FindTool(); }

        private static string FindTool()
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates = {
                Path.Combine(appDir, "ideviceinfo.exe"),
                Path.Combine(appDir, "libimobiledevice", "ideviceinfo.exe"),
                @"C:\Program Files\libimobiledevice\ideviceinfo.exe",
            };
            foreach (var c in candidates) if (File.Exists(c)) return c;
            return "ideviceinfo";
        }

        public string GetDeviceInfo(string udid = "")
        {
            var hw = new System.Text.StringBuilder();
            hw.AppendLine($"=== iOS-Abfrage via USB ({DateTime.Now:dd.MM.yyyy HH:mm}) ===");
            try
            {
                string args = string.IsNullOrEmpty(udid) ? "" : $"-u {udid}";
                string raw  = RunTool("ideviceinfo.exe", args);

                if (string.IsNullOrWhiteSpace(raw))
                    throw new Exception("Kein iOS-Gerät gefunden oder Gerät nicht vertraut.\n→ iPhone: 'Vertrauen' antippen und iTunes / Apple Devices installieren.");

                hw.AppendLine();
                hw.AppendLine("=== Gerät ===");
                hw.AppendLine(ExtractKey(raw, "DeviceName"));
                hw.AppendLine(ExtractKey(raw, "ProductType"));
                hw.AppendLine(ExtractKey(raw, "ProductVersion", "iOS-Version"));
                hw.AppendLine(ExtractKey(raw, "BuildVersion"));
                hw.AppendLine(ExtractKey(raw, "SerialNumber"));
                hw.AppendLine(ExtractKey(raw, "UniqueDeviceID", "UDID"));

                hw.AppendLine();
                hw.AppendLine("=== Speicher ===");
                hw.AppendLine(ExtractKey(raw, "TotalDiskCapacity", "Gesamt"));
                hw.AppendLine(ExtractKey(raw, "TotalDataCapacity", "Daten"));

                hw.AppendLine();
                hw.AppendLine("ℹ️  Softwareliste nicht verfügbar (kein Jailbreak).");
            }
            catch (Exception ex)
            {
                hw.AppendLine($"Fehler: {ex.Message}");
            }
            return hw.ToString();
        }

        public List<string> ListDeviceUdids()
        {
            var list = new List<string>();
            try
            {
                string raw = RunTool("idevice_id.exe", "-l");
                foreach (var line in raw.Split('\n'))
                    if (!string.IsNullOrWhiteSpace(line)) list.Add(line.Trim());
            }
            catch { }
            return list;
        }

        private string RunTool(string exe, string args)
        {
            string fullPath = Path.Combine(Path.GetDirectoryName(_ideviceinfoPath) ?? "", exe);
            if (!File.Exists(fullPath)) fullPath = exe;

            var psi = new ProcessStartInfo
            {
                FileName               = fullPath,
                Arguments              = args,
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true
            };
            using (var p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                return output;
            }
        }

        private static string ExtractKey(string raw, string key, string label = "")
        {
            foreach (var line in raw.Split('\n'))
                if (line.TrimStart().StartsWith(key + ":"))
                    return $"  {(string.IsNullOrEmpty(label) ? key : label),-20}: {line.Split(':')[1].Trim()}";
            return $"  {(string.IsNullOrEmpty(label) ? key : label),-20}: -";
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
