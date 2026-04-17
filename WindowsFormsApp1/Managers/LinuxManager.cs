using System;
using System.Collections.Generic;
using System.IO;
using Renci.SshNet;

namespace NmapInventory
{
    public class LinuxManager
    {
        private readonly int _defaultTimeout = 15;

        public string GetHardwareInfo(string host, string username, string password, string keyFile = "", int port = 22)
        {
            var hw = new System.Text.StringBuilder();
            hw.AppendLine($"=== Linux SSH-Abfrage: {host} ({DateTime.Now:dd.MM.yyyy HH:mm}) ===");

            try
            {
                using (var client = BuildClient(host, port, username, password, keyFile))
                {
                    client.Connect();

                    hw.AppendLine();
                    hw.AppendLine("=== Betriebssystem ===");
                    hw.AppendLine(Run(client, "cat /etc/os-release | grep -E '^(NAME|VERSION|ID)=' | tr -d '\"'"));
                    hw.AppendLine(Run(client, "uname -r"));

                    hw.AppendLine();
                    hw.AppendLine("=== Prozessor ===");
                    hw.AppendLine(Run(client, "lscpu | grep -E 'Model name|CPU\\(s\\)|Thread|Socket|Architecture'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Arbeitsspeicher ===");
                    hw.AppendLine(Run(client, "free -h | grep -E 'Mem|Swap'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Festplatten ===");
                    hw.AppendLine(Run(client, "df -h -x tmpfs -x devtmpfs -x overlay | grep -v 'udev'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Netzwerk ===");
                    hw.AppendLine(Run(client, "ip -br addr show | grep -v '^lo'"));

                    hw.AppendLine();
                    hw.AppendLine("=== Uptime ===");
                    hw.AppendLine(Run(client, "uptime -p 2>/dev/null || uptime"));

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                hw.AppendLine($"Fehler: {ex.Message}");
            }

            return hw.ToString();
        }

        public List<SoftwareInfo> GetInstalledSoftware(string host, string username, string password, string keyFile = "", int port = 22)
        {
            var list = new List<SoftwareInfo>();
            try
            {
                using (var client = BuildClient(host, port, username, password, keyFile))
                {
                    client.Connect();

                    string pm = Run(client, "which dpkg apt rpm pacman 2>/dev/null | head -1");
                    string raw;
                    string source;

                    if (pm.Contains("dpkg") || pm.Contains("apt"))
                    {
                        raw = Run(client, "dpkg-query -W -f='${Package}\\t${Version}\\t${Maintainer}\\n' 2>/dev/null");
                        source = "Linux/dpkg";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: 2);
                    }
                    else if (pm.Contains("rpm"))
                    {
                        raw = Run(client, "rpm -qa --queryformat '%{NAME}\\t%{VERSION}\\t%{VENDOR}\\n' 2>/dev/null");
                        source = "Linux/rpm";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: 2);
                    }
                    else if (pm.Contains("pacman"))
                    {
                        raw = Run(client, "pacman -Q 2>/dev/null");
                        source = "Linux/pacman";
                        list = ParseTsv(raw, host, source, nameCol: 0, versionCol: 1, publisherCol: -1);
                    }
                    else
                    {
                        list.Add(new SoftwareInfo { Name = "Kein unterstützter Paketmanager gefunden", Source = "Linux", PCName = host });
                    }

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                list.Add(new SoftwareInfo { Name = $"SSH-Fehler: {ex.Message}", Source = "Linux", PCName = host });
            }

            return list;
        }

        private SshClient BuildClient(string host, int port, string username, string password, string keyFile)
        {
            if (!string.IsNullOrWhiteSpace(keyFile) && File.Exists(keyFile))
            {
                var keyFiles = new[] { new PrivateKeyFile(keyFile) };
                return new SshClient(new ConnectionInfo(host, port, username,
                    new PrivateKeyAuthenticationMethod(username, keyFiles))
                { Timeout = TimeSpan.FromSeconds(_defaultTimeout) });
            }
            return new SshClient(new ConnectionInfo(host, port, username,
                new PasswordAuthenticationMethod(username, password))
            { Timeout = TimeSpan.FromSeconds(_defaultTimeout) });
        }

        private static string Run(SshClient client, string command)
        {
            using (var cmd = client.CreateCommand(command))
            {
                cmd.CommandTimeout = TimeSpan.FromSeconds(30);
                return (cmd.Execute() ?? "").Trim();
            }
        }

        private static List<SoftwareInfo> ParseTsv(string raw, string host, string source, int nameCol, int versionCol, int publisherCol)
        {
            var list = new List<SoftwareInfo>();
            foreach (var line in raw.Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var cols = line.Split('\t');
                list.Add(new SoftwareInfo
                {
                    Name      = cols.Length > nameCol    && nameCol    >= 0 ? cols[nameCol].Trim()    : line.Trim(),
                    Version   = cols.Length > versionCol && versionCol >= 0 ? cols[versionCol].Trim() : "",
                    Publisher = cols.Length > publisherCol && publisherCol >= 0 ? cols[publisherCol].Trim() : "",
                    Source    = source,
                    PCName    = host,
                    Timestamp = DateTime.Now
                });
            }
            return list;
        }
    }
}
