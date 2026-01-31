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
            hardwareInfoTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 9), ReadOnly = true };
            tabControl.TabPages.Add(new TabPage("Hardware") { Controls = { hardwareInfoTextBox } });

            // Software Tab
            softwareGridView = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Software", Width = 250, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert", Width = 120, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewButtonColumn { HeaderText = "Aktion", Text = "Update", Width = 80, UseColumnTextForButtonValue = true });
            softwareGridView.CellClick += (s, e) => OnSoftwareUpdateClick(e);
            tabControl.TabPages.Add(new TabPage("Software") { Controls = { softwareGridView } });

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
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 200 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert am", Width = 120 });

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
                Invoke(new MethodInvoker(() =>
                {
                    hardwareInfoTextBox.Text = hw;
                    currentSoftware = sw;
                    DisplaySoftwareGrid(sw);
                    dbManager.SaveSoftware(sw);
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
                            string hw = hardwareManager.GetRemoteHardwareInfo(form.ComputerIP);
                            List<SoftwareInfo> sw = softwareManager.GetRemoteSoftware(form.ComputerIP);
                            Invoke(new MethodInvoker(() =>
                            {
                                hardwareInfoTextBox.Text = hw;
                                DisplaySoftwareGrid(sw, form.ComputerIP);
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
            foreach (var sw in software.OrderBy(s => s.Name))
                softwareGridView.Rows.Add(sw.Name, sw.Version ?? "N/A", sw.Publisher ?? "", sw.InstallDate ?? "", "Update");
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
                dbSoftwareTable.Rows.Add(sw.Zeitstempel, sw.Name, sw.Version, sw.Publisher, sw.InstallDate);
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

        public string GetRemoteHardwareInfo(string computerIP)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== REMOTE HARDWARE: {computerIP} ===\n");
            try
            {
                var scope = new ManagementScope($@"\\{computerIP}\root\cimv2", new ConnectionOptions { Authentication = AuthenticationLevel.PacketPrivacy, Timeout = TimeSpan.FromSeconds(30) });
                scope.Connect();
                sb.AppendLine("=== BETRIEBSSYSTEM ===");
                var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Caption, Version FROM Win32_OperatingSystem"));
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("OS: " + obj["Caption"]);
                    sb.AppendLine("Version: " + obj["Version"]);
                }
            }
            catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
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

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP) => new List<SoftwareInfo>();

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
                CREATE TABLE IF NOT EXISTS Software (ID INTEGER PRIMARY KEY, Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP, Name TEXT, Version TEXT, Publisher TEXT, InstallLocation TEXT, Source TEXT, InstallDate TEXT);";
                using (var cmd = new SQLiteCommand(createTableQuery, conn)) cmd.ExecuteNonQuery();
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
                    using (var cmd = new SQLiteCommand("INSERT INTO Software (Name, Version, Publisher, InstallLocation, Source, InstallDate) VALUES (@Name, @Version, @Publisher, @InstallLocation, @Source, @InstallDate)", conn))
                    {
                        cmd.Parameters.AddWithValue("@Name", sw.Name);
                        cmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        cmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        cmd.Parameters.AddWithValue("@InstallLocation", sw.InstallLocation ?? "");
                        cmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                        cmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        cmd.ExecuteNonQuery();
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
                string query = $"SELECT Zeitstempel, Name, Version, Publisher, InstallDate FROM Software {whereClause} ORDER BY Zeitstempel DESC LIMIT 100";
                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        software.Add(new DatabaseSoftware
                        {
                            Zeitstempel = reader["Zeitstempel"].ToString(),
                            Name = reader["Name"].ToString(),
                            Version = reader["Version"].ToString(),
                            Publisher = reader["Publisher"].ToString(),
                            InstallDate = reader["InstallDate"].ToString()
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