using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;
using Newtonsoft.Json;
using JsonFormatting = Newtonsoft.Json.Formatting;

namespace NmapInventory
{
    public class MainForm : Form
    {
        private TabControl tabControl;
        private TextBox networkTextBox;
        private Button scanButton, hardwareButton, exportButton, remoteHardwareButton;
        private Label statusLabel;
        private DataGridView deviceTable, softwareGridView, dbDeviceTable, dbSoftwareTable;
        private TextBox rawOutputTextBox, hardwareInfoTextBox;

        private NmapScanner nmapScanner;
        private DatabaseManager dbManager;
        private HardwareManager hardwareManager;
        private SoftwareManager softwareManager;

        private List<DeviceInfo> currentDevices = new List<DeviceInfo>();
        private List<SoftwareInfo> currentSoftware = new List<SoftwareInfo>();
        private string currentRemotePC = "";

        public MainForm()
        {
            nmapScanner = new NmapScanner();
            dbManager = new DatabaseManager();
            hardwareManager = new HardwareManager();
            softwareManager = new SoftwareManager();

            InitializeComponent();
            dbManager.InitializeDatabase();
        }

        private void InitializeComponent()
        {
            Text = "Nmap Inventarisierung";
            Size = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;

            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 60 };
            networkTextBox = new TextBox { Text = "192.168.2.0/24", Location = new Point(80, 17), Width = 180 };

            scanButton = new Button { Text = "Nmap Scan", Location = new Point(280, 15) };
            scanButton.Click += (s, e) => StartScan();

            hardwareButton = new Button { Text = "Hardware / Software", Location = new Point(380, 15), Width = 160 };
            hardwareButton.Click += (s, e) => StartHardwareQuery();

            exportButton = new Button { Text = "Exportieren", Location = new Point(560, 15), Width = 100 };
            exportButton.Click += (s, e) => ExportData();

            remoteHardwareButton = new Button { Text = "Remote Hardware", Location = new Point(670, 15), Width = 140 };
            remoteHardwareButton.Click += (s, e) => StartRemoteHardwareQuery();

            topPanel.Controls.AddRange(new Control[] {
                new Label { Text = "Netzwerk:", Location = new Point(10, 20), AutoSize = true },
                networkTextBox, scanButton, hardwareButton, exportButton, remoteHardwareButton
            });

            tabControl = new TabControl { Dock = DockStyle.Fill };

            // Geräte Tab
            deviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 150 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 80 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 800 });
            tabControl.TabPages.Add(new TabPage("Geräte") { Controls = { deviceTable } });

            // Nmap Tab
            rawOutputTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            tabControl.TabPages.Add(new TabPage("Nmap Ausgabe") { Controls = { rawOutputTextBox } });

            // Hardware Tab
            Panel hardwareInfoPanel = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = Color.LightSteelBlue };
            var hardwarePCLabel = new Label
            {
                Name = "hardwarePCLabel",
                Text = "Lokaler PC: " + Environment.MachineName,
                Location = new Point(10, 8),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            var hardwareUpdateLabel = new Label
            {
                Name = "hardwareUpdateLabel",
                Text = "Letzte Aktualisierung: -",
                Location = new Point(400, 8),
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };
            hardwareInfoPanel.Controls.AddRange(new Control[] { hardwarePCLabel, hardwareUpdateLabel });

            hardwareInfoTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 9), ReadOnly = true };

            var hardwareContainer = new Panel { Dock = DockStyle.Fill };
            hardwareContainer.Controls.Add(hardwareInfoTextBox);
            hardwareContainer.Controls.Add(hardwareInfoPanel);

            var hardwareTabPage = new TabPage("Hardware") { Controls = { hardwareContainer } };
            hardwareTabPage.Name = "Hardware";
            tabControl.TabPages.Add(hardwareTabPage);

            // Software Tab
            Panel softwareInfoPanel = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = Color.LightSteelBlue };
            var softwarePCLabel = new Label
            {
                Name = "softwarePCLabel",
                Text = "Lokaler PC: " + Environment.MachineName,
                Location = new Point(10, 8),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            var softwareUpdateLabel = new Label
            {
                Name = "softwareUpdateLabel",
                Text = "Letzte Aktualisierung: -",
                Location = new Point(400, 8),
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };
            softwareInfoPanel.Controls.AddRange(new Control[] { softwarePCLabel, softwareUpdateLabel });

            softwareGridView = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Software", Width = 250, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert", Width = 120, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Letztes Update", Width = 140, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewButtonColumn { HeaderText = "Aktion", Text = "Update", Width = 80, UseColumnTextForButtonValue = true });
            softwareGridView.CellClick += (s, e) => OnSoftwareUpdateClick(e);

            var softwareContainer = new Panel { Dock = DockStyle.Fill };
            softwareContainer.Controls.Add(softwareGridView);
            softwareContainer.Controls.Add(softwareInfoPanel);

            var softwareTabPage = new TabPage("Software") { Controls = { softwareContainer } };
            softwareTabPage.Name = "Software";
            tabControl.TabPages.Add(softwareTabPage);

            // DB Geräte
            Panel dbDevicePanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var dbDeviceFilter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            dbDeviceFilter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            dbDeviceFilter.SelectedIndex = 0;
            dbDeviceFilter.SelectedIndexChanged += (s, e) => LoadDatabaseDevices(dbDeviceFilter.SelectedItem.ToString());

            dbDevicePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true },
                dbDeviceFilter
            });

            dbDeviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 80 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 400 });

            var dbDeviceContainer = new Panel { Dock = DockStyle.Fill };
            dbDeviceContainer.Controls.Add(dbDeviceTable);
            dbDeviceContainer.Controls.Add(dbDevicePanel);

            tabControl.TabPages.Add(new TabPage("DB - Geräte") { Controls = { dbDeviceContainer } });

            // DB Software
            Panel dbSoftwarePanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var dbSoftwareFilter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            dbSoftwareFilter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            dbSoftwareFilter.SelectedIndex = 0;
            dbSoftwareFilter.SelectedIndexChanged += (s, e) => LoadDatabaseSoftware(dbSoftwareFilter.SelectedItem.ToString());

            dbSoftwarePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true },
                dbSoftwareFilter
            });

            dbSoftwareTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "PC Name/IP", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 200 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert am", Width = 120 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Letztes Update", Width = 140 });

            var dbSoftwareContainer = new Panel { Dock = DockStyle.Fill };
            dbSoftwareContainer.Controls.Add(dbSoftwareTable);
            dbSoftwareContainer.Controls.Add(dbSoftwarePanel);

            tabControl.TabPages.Add(new TabPage("DB - Software") { Controls = { dbSoftwareContainer } });

            statusLabel = new Label { Dock = DockStyle.Bottom, Height = 25, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle };

            Controls.AddRange(new Control[] { tabControl, topPanel, statusLabel });
        }

        private void StartScan()
        {
            scanButton.Enabled = false;
            statusLabel.Text = "Scan läuft...";
            Task.Run(() =>
            {
                try
                {
                    var result = nmapScanner.Scan(networkTextBox.Text);
                    Invoke(new MethodInvoker(() =>
                    {
                        rawOutputTextBox.Text = result.RawOutput;
                        currentDevices = result.Devices;
                        DisplayDevices(currentDevices);
                        dbManager.SaveDevices(currentDevices);
                        LoadDatabaseDevices();
                        statusLabel.Text = "Scan abgeschlossen";
                        scanButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new MethodInvoker(() =>
                    {
                        MessageBox.Show($"Fehler: {ex.Message}");
                        statusLabel.Text = "Fehler beim Scan";
                        scanButton.Enabled = true;
                    }));
                }
            });
        }

        private void StartHardwareQuery()
        {
            hardwareButton.Enabled = false;
            statusLabel.Text = "Daten werden gesammelt...";
            Task.Run(() =>
            {
                string hw = hardwareManager.GetHardwareInfo();
                List<SoftwareInfo> sw = softwareManager.GetInstalledSoftware();
                // Setze PCName für lokale Software
                string pcName = Environment.MachineName;
                foreach (var software in sw)
                {
                    software.PCName = pcName;
                    software.Timestamp = DateTime.Now;
                }
                // Prüfe auf Updates
                dbManager.CheckForUpdates(sw, pcName);

                Invoke(new MethodInvoker(() =>
                {
                    hardwareInfoTextBox.Text = hw;
                    currentSoftware = sw;
                    DisplaySoftwareGrid(sw);
                    dbManager.SaveSoftware(sw);
                    UpdateHardwareLabels(pcName, false);
                    statusLabel.Text = "Fertig";
                    hardwareButton.Enabled = true;
                }));
            });
        }

        private void StartRemoteHardwareQuery()
        {
            using (var form = new RemoteConnectionForm())
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    remoteHardwareButton.Enabled = false;
                    statusLabel.Text = $"Verbinde zu {form.ComputerIP}...";
                    Task.Run(() =>
                    {
                        try
                        {
                            string hw = hardwareManager.GetRemoteHardwareInfo(form.ComputerIP, form.Username, form.Password);
                            List<SoftwareInfo> sw = softwareManager.GetRemoteSoftware(form.ComputerIP, form.Username, form.Password);
                            // Setze PCName für Remote-Software
                            string pcName = form.ComputerIP;
                            foreach (var software in sw)
                            {
                                software.PCName = pcName;
                                software.Timestamp = DateTime.Now;
                            }
                            // Prüfe auf Updates
                            dbManager.CheckForUpdates(sw, pcName);

                            Invoke(new MethodInvoker(() =>
                            {
                                hardwareInfoTextBox.Text = hw;
                                DisplaySoftwareGrid(sw, pcName);
                                dbManager.SaveSoftware(sw);
                                UpdateHardwareLabels(pcName, true);
                                statusLabel.Text = "Remote Abfrage abgeschlossen";
                                remoteHardwareButton.Enabled = true;
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new MethodInvoker(() =>
                            {
                                MessageBox.Show($"Fehler: {ex.Message}");
                                statusLabel.Text = "Fehler bei Remote Abfrage";
                                remoteHardwareButton.Enabled = true;
                            }));
                        }
                    });
                }
            }
        }

        private void DisplayDevices(List<DeviceInfo> devices)
        {
            deviceTable.Rows.Clear();
            foreach (var dev in devices)
                deviceTable.Rows.Add(dev.IP, dev.Hostname, dev.Status, dev.Ports);
        }

        private void DisplaySoftwareGrid(List<SoftwareInfo> software, string remotePC = "")
        {
            currentRemotePC = remotePC;
            softwareGridView.Rows.Clear();

            // Finde die Labels im Software-Tab und aktualisiere sie
            var softwareTab = tabControl.TabPages["Software"];
            if (softwareTab != null)
            {
                var container = softwareTab.Controls[0] as Panel;
                if (container != null)
                {
                    var infoPanel = container.Controls.OfType<Panel>().FirstOrDefault();
                    if (infoPanel != null)
                    {
                        var pcLabel = infoPanel.Controls["softwarePCLabel"] as Label;
                        var updateLabel = infoPanel.Controls["softwareUpdateLabel"] as Label;

                        if (pcLabel != null)
                        {
                            if (!string.IsNullOrEmpty(remotePC))
                            {
                                pcLabel.Text = $"Remote-PC: {remotePC}";
                                pcLabel.ForeColor = Color.DarkRed;
                            }
                            else
                            {
                                pcLabel.Text = $"Lokaler PC: {Environment.MachineName}";
                                pcLabel.ForeColor = Color.DarkBlue;
                            }
                        }

                        if (updateLabel != null)
                        {
                            updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
                        }
                    }
                }
            }

            foreach (var sw in software.OrderBy(s => s.Name))
                softwareGridView.Rows.Add(sw.Name, sw.Version ?? "N/A", sw.Publisher ?? "", sw.InstallDate ?? "", sw.LastUpdate ?? "-", "Update");
        }

        private void LoadDatabaseDevices(string filter = "Alle")
        {
            dbDeviceTable.Rows.Clear();
            var devices = dbManager.LoadDevices(filter);
            foreach (var dev in devices)
                dbDeviceTable.Rows.Add(dev.Zeitstempel, dev.IP, dev.Hostname, dev.Status, dev.Ports);
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var software = dbManager.LoadSoftware(filter);
            foreach (var sw in software)
                dbSoftwareTable.Rows.Add(sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
        }

        private void ExportData()
        {
            using (var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json|XML (*.xml)|*.xml" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var exporter = new DataExporter();
                        exporter.Export(sfd.FileName, currentDevices, currentSoftware, hardwareInfoTextBox.Text);
                        MessageBox.Show("Daten exportiert!", "Erfolg");
                    }
                    catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
                }
            }
        }

        private void UpdateHardwareLabels(string pcName, bool isRemote)
        {
            var hardwareTab = tabControl.TabPages["Hardware"];
            if (hardwareTab != null)
            {
                var container = hardwareTab.Controls[0] as Panel;
                if (container != null)
                {
                    var infoPanel = container.Controls.OfType<Panel>().FirstOrDefault();
                    if (infoPanel != null)
                    {
                        var pcLabel = infoPanel.Controls["hardwarePCLabel"] as Label;
                        var updateLabel = infoPanel.Controls["hardwareUpdateLabel"] as Label;

                        if (pcLabel != null)
                        {
                            if (isRemote)
                            {
                                pcLabel.Text = $"Remote-PC: {pcName}";
                                pcLabel.ForeColor = Color.DarkRed;
                            }
                            else
                            {
                                pcLabel.Text = $"Lokaler PC: {pcName}";
                                pcLabel.ForeColor = Color.DarkBlue;
                            }
                        }

                        if (updateLabel != null)
                        {
                            updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
                        }
                    }
                }
            }
        }

        private void OnSoftwareUpdateClick(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == softwareGridView.Columns.Count - 1)
            {
                string softwareName = softwareGridView.Rows[e.RowIndex].Cells[0].Value?.ToString();
                if (!string.IsNullOrEmpty(softwareName))
                    softwareManager.UpdateSoftware(softwareName, currentRemotePC, statusLabel);
            }
        }

    }

    // ===== DATA CLASSES =====
    public class DeviceInfo
    {
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
    }

    public class SoftwareInfo
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallLocation { get; set; }
        public string Source { get; set; }
        public string InstallDate { get; set; }
        public string PCName { get; set; }
        public DateTime Timestamp { get; set; }
        public string LastUpdate { get; set; }
    }

    public class NmapScanResult
    {
        public string RawOutput { get; set; }
        public List<DeviceInfo> Devices { get; set; }
    }

    public class DatabaseDevice
    {
        public string Zeitstempel { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
    }

    public class DatabaseSoftware
    {
        public string Zeitstempel { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string PCName { get; set; }
        public string LastUpdate { get; set; }
    }

    // ===== MANAGER CLASSES =====
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
            sb.AppendLine("=== BETRIEBSSYSTEM ===");
            sb.AppendLine("Computer: " + Environment.MachineName);
            sb.AppendLine("OS: " + Environment.OSVersion.VersionString);
            sb.AppendLine();
            sb.AppendLine("=== PROZESSOR ===");
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT Name, MaxClockSpeed FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Name: " + obj["Name"]);
                    sb.AppendLine("Taktfrequenz: " + obj["MaxClockSpeed"] + " MHz");
                }
            }
            catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
            return sb.ToString();
        }

        public string GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== REMOTE HARDWARE: {computerIP} ===\n");
            try
            {
                var options = new ConnectionOptions
                {
                    Authentication = AuthenticationLevel.PacketPrivacy,
                    Timeout = TimeSpan.FromSeconds(30)
                };

                // Nur Username/Password setzen wenn sie angegeben wurden
                if (!string.IsNullOrEmpty(username))
                {
                    options.Username = username;
                    options.Password = password;
                }

                var scope = new ManagementScope($@"\\{computerIP}\root\cimv2", options);
                scope.Connect();

                sb.AppendLine("=== BETRIEBSSYSTEM ===");
                var osSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Caption, Version, OSArchitecture FROM Win32_OperatingSystem"));
                foreach (ManagementObject obj in osSearcher.Get())
                {
                    sb.AppendLine("OS: " + obj["Caption"]);
                    sb.AppendLine("Version: " + obj["Version"]);
                    sb.AppendLine("Architektur: " + obj["OSArchitecture"]);
                }

                sb.AppendLine("\n=== PROZESSOR ===");
                var cpuSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor"));
                foreach (ManagementObject obj in cpuSearcher.Get())
                {
                    sb.AppendLine("Name: " + obj["Name"]);
                    sb.AppendLine("Taktfrequenz: " + obj["MaxClockSpeed"] + " MHz");
                    sb.AppendLine("Kerne: " + obj["NumberOfCores"]);
                    sb.AppendLine("Logische Prozessoren: " + obj["NumberOfLogicalProcessors"]);
                }

                sb.AppendLine("\n=== ARBEITSSPEICHER ===");
                var memSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"));
                foreach (ManagementObject obj in memSearcher.Get())
                {
                    long totalMB = Convert.ToInt64(obj["TotalVisibleMemorySize"]) / 1024;
                    long freeMB = Convert.ToInt64(obj["FreePhysicalMemory"]) / 1024;
                    sb.AppendLine($"Gesamt: {totalMB} MB");
                    sb.AppendLine($"Frei: {freeMB} MB");
                    sb.AppendLine($"Belegt: {totalMB - freeMB} MB");
                }

                sb.AppendLine("\n=== FESTPLATTEN ===");
                var diskSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT DeviceID, Size, FreeSpace, VolumeName FROM Win32_LogicalDisk WHERE DriveType=3"));
                foreach (ManagementObject obj in diskSearcher.Get())
                {
                    string deviceID = obj["DeviceID"]?.ToString();
                    string volumeName = obj["VolumeName"]?.ToString();
                    long sizeGB = Convert.ToInt64(obj["Size"]) / (1024 * 1024 * 1024);
                    long freeGB = Convert.ToInt64(obj["FreeSpace"]) / (1024 * 1024 * 1024);
                    sb.AppendLine($"\nLaufwerk: {deviceID} ({volumeName})");
                    sb.AppendLine($"  Größe: {sizeGB} GB");
                    sb.AppendLine($"  Frei: {freeGB} GB");
                    sb.AppendLine($"  Belegt: {sizeGB - freeGB} GB");
                }

                sb.AppendLine("\n=== NETZWERKADAPTER ===");
                var netSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Description, MACAddress, IPEnabled FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"));
                foreach (ManagementObject obj in netSearcher.Get())
                {
                    sb.AppendLine("Adapter: " + obj["Description"]);
                    sb.AppendLine("MAC: " + obj["MACAddress"]);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("\nFehler: " + ex.Message);
                if (ex.InnerException != null)
                    sb.AppendLine("Details: " + ex.InnerException.Message);
            }
            return sb.ToString();
        }
    }

    public class SoftwareManager
    {
        public List<SoftwareInfo> GetInstalledSoftware()
        {
            var software = new Dictionary<string, SoftwareInfo>();
            ReadRegistry(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", software);
            return software.Values.ToList();
        }

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            var software = new Dictionary<string, SoftwareInfo>();
            try
            {
                // Verwende Registry-Abfrage statt Win32_Product (schneller und zuverlässiger)
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

                // Registry über WMI abfragen
                var registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);

                // Beide Registry-Pfade durchsuchen
                string[] regPaths = new string[]
                {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };

                foreach (string regPath in regPaths)
                {
                    // Subkeys auflisten
                    ManagementBaseObject inParams = registry.GetMethodParameters("EnumKey");
                    inParams["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
                    inParams["sSubKeyName"] = regPath;

                    ManagementBaseObject outParams = registry.InvokeMethod("EnumKey", inParams, null);
                    if (outParams["sNames"] != null)
                    {
                        string[] subkeys = (string[])outParams["sNames"];

                        foreach (string subkey in subkeys)
                        {
                            string fullPath = regPath + @"\" + subkey;

                            // DisplayName auslesen
                            string displayName = GetRegistryValue(registry, fullPath, "DisplayName");
                            if (string.IsNullOrEmpty(displayName) || displayName.Contains("Update")) continue;

                            string version = GetRegistryValue(registry, fullPath, "DisplayVersion") ?? "N/A";
                            string publisher = GetRegistryValue(registry, fullPath, "Publisher") ?? "";
                            string installDate = GetRegistryValue(registry, fullPath, "InstallDate") ?? "";

                            // Datum formatieren (von YYYYMMDD zu DD.MM.YYYY)
                            if (!string.IsNullOrEmpty(installDate) && installDate.Length == 8)
                            {
                                installDate = $"{installDate.Substring(6, 2)}.{installDate.Substring(4, 2)}.{installDate.Substring(0, 4)}";
                            }

                            var uniqueKey = displayName + "|" + version;
                            if (!software.ContainsKey(uniqueKey))
                            {
                                software[uniqueKey] = new SoftwareInfo
                                {
                                    Name = displayName,
                                    Version = version,
                                    Publisher = publisher,
                                    InstallDate = installDate,
                                    Source = "Remote Registry"
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Abrufen der Remote-Software:\n{ex.Message}\n\nTipp: Stelle sicher, dass:\n1. WMI auf dem Remote-PC aktiviert ist\n2. Die Firewall WMI-Zugriff erlaubt\n3. Du Administrator-Rechte hast", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return software.Values.ToList();
        }

        private string GetRegistryValue(ManagementClass registry, string keyPath, string valueName)
        {
            try
            {
                ManagementBaseObject inParams = registry.GetMethodParameters("GetStringValue");
                inParams["hDefKey"] = 0x80000002; // HKEY_LOCAL_MACHINE
                inParams["sSubKeyName"] = keyPath;
                inParams["sValueName"] = valueName;

                ManagementBaseObject outParams = registry.InvokeMethod("GetStringValue", inParams, null);
                if (outParams["sValue"] != null)
                {
                    return outParams["sValue"].ToString();
                }
            }
            catch { }
            return null;
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
                                software[uniqueKey] = new SoftwareInfo { Name = name, Version = version, Publisher = sub?.GetValue("Publisher")?.ToString() ?? "", Source = "Registry" };
                        }
                    }
                }
            }
            catch { }
        }
    }

    public class DatabaseManager
    {
        private string dbPath = "nmap_inventory.db";

        public void InitializeDatabase()
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                string createTableQuery = @"CREATE TABLE IF NOT EXISTS Geraete (ID INTEGER PRIMARY KEY, Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP, IP TEXT, Hostname TEXT, Status TEXT, Ports TEXT);
                CREATE TABLE IF NOT EXISTS Software (ID INTEGER PRIMARY KEY, Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP, Name TEXT, Version TEXT, Publisher TEXT, InstallLocation TEXT, Source TEXT, InstallDate TEXT, PCName TEXT, LastUpdate TEXT);";
                using (var cmd = new SQLiteCommand(createTableQuery, conn)) cmd.ExecuteNonQuery();

                // Prüfe ob PCName-Spalte existiert, falls nicht, füge sie hinzu (für bestehende Datenbanken)
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE Software ADD COLUMN PCName TEXT", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { /* Spalte existiert bereits */ }

                // Prüfe ob LastUpdate-Spalte existiert, falls nicht, füge sie hinzu
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE Software ADD COLUMN LastUpdate TEXT", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { /* Spalte existiert bereits */ }
            }
        }

        public void SaveDevices(List<DeviceInfo> devices)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                foreach (var dev in devices)
                    using (var cmd = new SQLiteCommand("INSERT INTO Geraete (IP, Hostname, Status, Ports) VALUES (@IP, @Hostname, @Status, @Ports)", conn))
                    {
                        cmd.Parameters.AddWithValue("@IP", dev.IP);
                        cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                        cmd.Parameters.AddWithValue("@Status", dev.Status);
                        cmd.Parameters.AddWithValue("@Ports", dev.Ports);
                        cmd.ExecuteNonQuery();
                    }
            }
        }

        public void SaveSoftware(List<SoftwareInfo> software)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                foreach (var sw in software)
                    using (var cmd = new SQLiteCommand("INSERT INTO Software (Name, Version, Publisher, InstallLocation, Source, InstallDate, PCName, LastUpdate) VALUES (@Name, @Version, @Publisher, @InstallLocation, @Source, @InstallDate, @PCName, @LastUpdate)", conn))
                    {
                        cmd.Parameters.AddWithValue("@Name", sw.Name);
                        cmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        cmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        cmd.Parameters.AddWithValue("@InstallLocation", sw.InstallLocation ?? "");
                        cmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                        cmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        cmd.Parameters.AddWithValue("@PCName", sw.PCName ?? "");
                        cmd.Parameters.AddWithValue("@LastUpdate", sw.LastUpdate ?? "");
                        cmd.ExecuteNonQuery();
                    }
            }
        }

        public void CheckForUpdates(List<SoftwareInfo> softwareList, string pcName)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                foreach (var sw in softwareList)
                {
                    // Hole die letzte Version dieser Software für diesen PC
                    string query = "SELECT Version, Zeitstempel FROM Software WHERE Name = @Name AND PCName = @PCName ORDER BY Zeitstempel DESC LIMIT 1";
                    using (var cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Name", sw.Name);
                        cmd.Parameters.AddWithValue("@PCName", pcName);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string oldVersion = reader["Version"]?.ToString();
                                string lastTimestamp = reader["Zeitstempel"]?.ToString();

                                // Vergleiche Versionen
                                if (!string.IsNullOrEmpty(oldVersion) && oldVersion != sw.Version)
                                {
                                    // Version hat sich geändert - Update erkannt!
                                    sw.LastUpdate = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                                }
                                else if (!string.IsNullOrEmpty(lastTimestamp))
                                {
                                    // Gleiche Version - nutze letzten Zeitstempel
                                    sw.LastUpdate = lastTimestamp;
                                }
                                else
                                {
                                    sw.LastUpdate = "-";
                                }
                            }
                            else
                            {
                                // Erste Erfassung dieser Software
                                sw.LastUpdate = "Neu erfasst";
                            }
                        }
                    }
                }
            }
        }

        public List<DatabaseDevice> LoadDevices(string filter = "Alle")
        {
            var devices = new List<DatabaseDevice>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                string whereClause = GetDateFilter(filter);
                string query = $"SELECT Zeitstempel, IP, Hostname, Status, Ports FROM Geraete {whereClause} ORDER BY Zeitstempel DESC LIMIT 100";
                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        devices.Add(new DatabaseDevice
                        {
                            Zeitstempel = reader["Zeitstempel"].ToString(),
                            IP = reader["IP"].ToString(),
                            Hostname = reader["Hostname"].ToString(),
                            Status = reader["Status"].ToString(),
                            Ports = reader["Ports"].ToString()
                        });
            }
            return devices;
        }

        public List<DatabaseSoftware> LoadSoftware(string filter = "Alle")
        {
            var software = new List<DatabaseSoftware>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                string whereClause = GetDateFilter(filter);
                string query = $"SELECT Zeitstempel, Name, Version, Publisher, InstallDate, PCName, LastUpdate FROM Software {whereClause} ORDER BY Zeitstempel DESC LIMIT 100";
                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        software.Add(new DatabaseSoftware
                        {
                            Zeitstempel = reader["Zeitstempel"].ToString(),
                            Name = reader["Name"].ToString(),
                            Version = reader["Version"].ToString(),
                            Publisher = reader["Publisher"].ToString(),
                            InstallDate = reader["InstallDate"].ToString(),
                            PCName = reader["PCName"].ToString(),
                            LastUpdate = reader["LastUpdate"].ToString()
                        });
            }
            return software;
        }

        private string GetDateFilter(string filter)
        {
            switch (filter)
            {
                case "Heute":
                    return "WHERE DATE(Zeitstempel) = DATE('now')";
                case "Diese Woche":
                    return "WHERE DATE(Zeitstempel) >= DATE('now', '-7 days')";
                case "Dieser Monat":
                    return "WHERE DATE(Zeitstempel) >= DATE('now', 'start of month')";
                case "Dieses Jahr":
                    return "WHERE DATE(Zeitstempel) >= DATE('now', 'start of year')";
                default:
                    return "";
            }
        }
    }

    public class DataExporter
    {
        public void Export(string filePath, List<DeviceInfo> devices, List<SoftwareInfo> software, string hardware)
        {
            if (filePath.EndsWith(".json"))
            {
                var data = new { Zeitstempel = DateTime.Now, Geräte = devices, Software = software, Hardware = hardware };
                File.WriteAllText(filePath, JsonConvert.SerializeObject(data, JsonFormatting.Indented));
            }
            else
            {
                var doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement("Inventar", new XElement("Zeitstempel", DateTime.Now), new XElement("Geräte", devices.Select(d => new XElement("Gerät", new XElement("IP", d.IP), new XElement("Hostname", d.Hostname ?? "")))), new XElement("Hardware", new XCData(hardware))));
                doc.Save(filePath);
            }
        }
    }

    public class RemoteConnectionForm : Form
    {
        public string ComputerIP { get; private set; }
        public string Username { get; private set; }
        public string Password { get; private set; }

        public RemoteConnectionForm()
        {
            Text = "Remote Verbindung";
            Width = 500;
            Height = 280;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var ipTb = new TextBox { Location = new Point(150, 30), Width = 300 };
            var userTb = new TextBox { Location = new Point(150, 70), Width = 300 };
            var passTb = new TextBox { Location = new Point(150, 110), Width = 300, UseSystemPasswordChar = true };

            var okBtn = new Button { Text = "Verbinden", Location = new Point(150, 160), Width = 100, DialogResult = DialogResult.OK };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 160), Width = 100, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] {
                new Label { Text = "Computer-IP:", Location = new Point(20, 33), AutoSize = true },
                ipTb,
                new Label { Text = "Benutzername:", Location = new Point(20, 73), AutoSize = true },
                userTb,
                new Label { Text = "Passwort:", Location = new Point(20, 113), AutoSize = true },
                passTb,
                okBtn, cancelBtn
            });

            AcceptButton = okBtn;
            CancelButton = cancelBtn;

            okBtn.Click += (s, e) => { ComputerIP = ipTb.Text; Username = userTb.Text; Password = passTb.Text; };
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}