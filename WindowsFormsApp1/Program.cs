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
        private Button scanButton;
        private Button hardwareButton;
        private Button exportButton;
        private Button remoteHardwareButton;
        private Label statusLabel;

        private DataGridView deviceTable;
        private TextBox rawOutputTextBox;
        private TextBox hardwareInfoTextBox;
        private TextBox softwareInfoTextBox;

        private DataGridView dbDeviceTable;
        private DataGridView dbSoftwareTable;
        private ComboBox timeRangeComboBox;
        private Button refreshDbButton;

        private List<DeviceInfo> currentDevices = new List<DeviceInfo>();
        private string currentHardwareInfo = "";
        private List<SoftwareInfo> currentSoftware = new List<SoftwareInfo>();

        private string dbPath = "nmap_inventory.db";

        public MainForm()
        {
            InitializeComponent();
            InitializeDatabase();
        }

        private void InitializeComponent()
        {
            Text = "Nmap Inventarisierung";
            Size = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;

            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60
            };

            Label lbl = new Label
            {
                Text = "Netzwerk:",
                Location = new Point(10, 20),
                AutoSize = true
            };

            networkTextBox = new TextBox
            {
                Text = "192.168.2.0/24",
                Location = new Point(80, 17),
                Width = 180
            };

            scanButton = new Button
            {
                Text = "Nmap Scan",
                Location = new Point(280, 15)
            };
            scanButton.Click += new EventHandler(StartScan);

            hardwareButton = new Button
            {
                Text = "Hardware / Software",
                Location = new Point(380, 15),
                Width = 160
            };
            hardwareButton.Click += new EventHandler(StartHardwareQuery);

            exportButton = new Button
            {
                Text = "Exportieren",
                Location = new Point(560, 15),
                Width = 100
            };
            exportButton.Click += new EventHandler(ExportData);

            remoteHardwareButton = new Button
            {
                Text = "Remote Hardware",
                Location = new Point(670, 15),
                Width = 140
            };
            remoteHardwareButton.Click += new EventHandler(StartRemoteHardwareQuery);

            topPanel.Controls.Add(lbl);
            topPanel.Controls.Add(networkTextBox);
            topPanel.Controls.Add(scanButton);
            topPanel.Controls.Add(hardwareButton);
            topPanel.Controls.Add(exportButton);
            topPanel.Controls.Add(remoteHardwareButton);

            tabControl = new TabControl { Dock = DockStyle.Fill };

            // Geräte
            deviceTable = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToAddRows = false,
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                ScrollBars = ScrollBars.Both
            };

            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", DataPropertyName = "IP", Width = 120 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", DataPropertyName = "Hostname", Width = 150 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", DataPropertyName = "Status", Width = 80 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", DataPropertyName = "Ports", Width = 800, AutoSizeMode = DataGridViewAutoSizeColumnMode.None });

            TabPage devicesTab = new TabPage("Geräte");
            devicesTab.Controls.Add(deviceTable);

            // Nmap
            rawOutputTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 10),
                ReadOnly = true
            };
            TabPage nmapTab = new TabPage("Nmap Ausgabe");
            nmapTab.Controls.Add(rawOutputTextBox);

            // Hardware
            hardwareInfoTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 9),
                ReadOnly = true
            };
            TabPage hwTab = new TabPage("Hardware");
            hwTab.Controls.Add(hardwareInfoTextBox);

            // Software
            softwareInfoTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Font = new Font("Consolas", 9),
                ReadOnly = true
            };
            TabPage swTab = new TabPage("Software");
            swTab.Controls.Add(softwareInfoTextBox);

            tabControl.TabPages.Add(devicesTab);
            tabControl.TabPages.Add(nmapTab);
            tabControl.TabPages.Add(hwTab);
            tabControl.TabPages.Add(swTab);

            // ===== Datenbank-Tabs =====

            Panel dbDevicePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50
            };

            Label timeLabel = new Label
            {
                Text = "Zeitraum:",
                Location = new Point(10, 15),
                AutoSize = true
            };

            timeRangeComboBox = new ComboBox
            {
                Location = new Point(80, 12),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            timeRangeComboBox.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat" });
            timeRangeComboBox.SelectedIndex = 0;

            refreshDbButton = new Button
            {
                Text = "Aktualisieren",
                Location = new Point(250, 12),
                Width = 100
            };
            refreshDbButton.Click += new EventHandler(RefreshDatabaseView);

            dbDevicePanel.Controls.Add(timeLabel);
            dbDevicePanel.Controls.Add(timeRangeComboBox);
            dbDevicePanel.Controls.Add(refreshDbButton);

            Panel dbDeviceContainerPanel = new Panel { Dock = DockStyle.Fill };

            dbDeviceTable = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                ScrollBars = ScrollBars.Both
            };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", DataPropertyName = "Zeitstempel", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", DataPropertyName = "IP", Width = 120 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", DataPropertyName = "Hostname", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", DataPropertyName = "Status", Width = 80 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", DataPropertyName = "Ports", Width = 400 });

            dbDeviceContainerPanel.Controls.Add(dbDeviceTable);
            dbDeviceContainerPanel.Controls.Add(dbDevicePanel);

            TabPage dbDeviceTab = new TabPage("DB - Geräte");
            dbDeviceTab.Controls.Add(dbDeviceContainerPanel);

            // Software-Datenbank
            Panel dbSoftwarePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50
            };

            Label softwareTimeLabel = new Label
            {
                Text = "Zeitraum:",
                Location = new Point(10, 15),
                AutoSize = true
            };

            ComboBox softwareTimeRangeComboBox = new ComboBox
            {
                Location = new Point(80, 12),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Tag = dbSoftwareTable
            };
            softwareTimeRangeComboBox.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat" });
            softwareTimeRangeComboBox.SelectedIndex = 0;
            softwareTimeRangeComboBox.SelectedIndexChanged += (s, e) => RefreshSoftwareDatabaseView(softwareTimeRangeComboBox);

            Button refreshSoftwareDbButton = new Button
            {
                Text = "Aktualisieren",
                Location = new Point(250, 12),
                Width = 100
            };
            refreshSoftwareDbButton.Click += (s, e) => RefreshSoftwareDatabaseView(softwareTimeRangeComboBox);

            dbSoftwarePanel.Controls.Add(softwareTimeLabel);
            dbSoftwarePanel.Controls.Add(softwareTimeRangeComboBox);
            dbSoftwarePanel.Controls.Add(refreshSoftwareDbButton);

            Panel dbSoftwareContainerPanel = new Panel { Dock = DockStyle.Fill };

            dbSoftwareTable = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                ScrollBars = ScrollBars.Both
            };
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", DataPropertyName = "Zeitstempel", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", DataPropertyName = "Name", Width = 200 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", DataPropertyName = "Version", Width = 100 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", DataPropertyName = "Publisher", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installationsort", DataPropertyName = "InstallLocation", Width = 250 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert am", DataPropertyName = "InstallDate", Width = 120 });

            dbSoftwareContainerPanel.Controls.Add(dbSoftwareTable);
            dbSoftwareContainerPanel.Controls.Add(dbSoftwarePanel);

            TabPage dbSoftwareTabNew = new TabPage("DB - Software");
            dbSoftwareTabNew.Controls.Add(dbSoftwareContainerPanel);

            tabControl.TabPages.Add(dbDeviceTab);
            tabControl.TabPages.Add(dbSoftwareTabNew);

            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 25,
                Text = "Bereit",
                BorderStyle = BorderStyle.FixedSingle
            };

            Controls.Add(tabControl);
            Controls.Add(topPanel);
            Controls.Add(statusLabel);
        }

        // ================= DATENBANK =================

        private void InitializeDatabase()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    string createTableQuery = @"
                        CREATE TABLE IF NOT EXISTS Geraete (
                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                            Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP,
                            IP TEXT NOT NULL,
                            Hostname TEXT,
                            Status TEXT,
                            Ports TEXT
                        );

                        CREATE TABLE IF NOT EXISTS Software (
                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                            Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP,
                            Name TEXT NOT NULL,
                            Version TEXT,
                            Publisher TEXT,
                            InstallLocation TEXT,
                            Source TEXT,
                            InstallDate TEXT
                        );

                        CREATE TABLE IF NOT EXISTS Hardware (
                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                            Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP,
                            ComputerName TEXT,
                            BenutzerName TEXT,
                            Betriebssystem TEXT,
                            Info TEXT
                        );

                        CREATE TABLE IF NOT EXISTS Treiber (
                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                            Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP,
                            ComputerName TEXT,
                            Treiber TEXT,
                            Version TEXT,
                            LetztesUpdate TEXT,
                            Quelle TEXT
                        );
                    ";

                    using (SQLiteCommand cmd = new SQLiteCommand(createTableQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    statusLabel.Text = "Datenbank initialisiert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler bei Datenbank-Initialisierung: " + ex.Message);
            }
        }

        private void SaveDevicesToDatabase()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    foreach (var device in currentDevices)
                    {
                        string query = @"INSERT INTO Geraete (IP, Hostname, Status, Ports) 
                                       VALUES (@IP, @Hostname, @Status, @Ports)";

                        using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@IP", device.IP);
                            cmd.Parameters.AddWithValue("@Hostname", device.Hostname ?? "");
                            cmd.Parameters.AddWithValue("@Status", device.Status);
                            cmd.Parameters.AddWithValue("@Ports", device.Ports);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    statusLabel.Text = $"{currentDevices.Count} Geräte in Datenbank gespeichert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Speichern: " + ex.Message);
            }
        }

        private void SaveSoftwareToDatabase()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    foreach (var software in currentSoftware)
                    {
                        string query = @"INSERT INTO Software (Name, Version, Publisher, InstallLocation, Source, InstallDate) 
                                       VALUES (@Name, @Version, @Publisher, @InstallLocation, @Source, @InstallDate)";

                        using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@Name", software.Name);
                            cmd.Parameters.AddWithValue("@Version", software.Version ?? "");
                            cmd.Parameters.AddWithValue("@Publisher", software.Publisher ?? "");
                            cmd.Parameters.AddWithValue("@InstallLocation", software.InstallLocation ?? "");
                            cmd.Parameters.AddWithValue("@Source", software.Source ?? "");
                            cmd.Parameters.AddWithValue("@InstallDate", software.InstallDate ?? "");
                            cmd.ExecuteNonQuery();
                        }
                    }

                    statusLabel.Text = $"{currentSoftware.Count} Software-Einträge in Datenbank gespeichert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Speichern: " + ex.Message);
            }
        }

        public List<DeviceInfo> LoadDevicesFromDatabase()
        {
            List<DeviceInfo> devices = new List<DeviceInfo>();

            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    string query = "SELECT IP, Hostname, Status, Ports FROM Geraete ORDER BY Zeitstempel DESC LIMIT 100";

                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                devices.Add(new DeviceInfo
                                {
                                    IP = reader["IP"].ToString(),
                                    Hostname = reader["Hostname"].ToString(),
                                    Status = reader["Status"].ToString(),
                                    Ports = reader["Ports"].ToString()
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Laden: " + ex.Message);
            }

            return devices;
        }

        private void RefreshDatabaseView(object sender, EventArgs e)
        {
            LoadDevicesFromDatabaseView();
        }

        private void LoadDevicesFromDatabaseView()
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    string whereClause = GetTimeRangeWhereClause();
                    string query = $@"SELECT Zeitstempel, IP, Hostname, Status, Ports 
                                     FROM Geraete 
                                     WHERE {whereClause}
                                     ORDER BY Zeitstempel DESC";

                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            dbDeviceTable.Rows.Clear();

                            while (reader.Read())
                            {
                                dbDeviceTable.Rows.Add(
                                    reader["Zeitstempel"].ToString(),
                                    reader["IP"].ToString(),
                                    reader["Hostname"].ToString(),
                                    reader["Status"].ToString(),
                                    reader["Ports"].ToString()
                                );
                            }
                        }
                    }
                }

                statusLabel.Text = $"Geräte geladen: {dbDeviceTable.Rows.Count} Einträge";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Laden der Geräte: " + ex.Message);
            }
        }

        private string GetTimeRangeWhereClause()
        {
            string timeRange = timeRangeComboBox.SelectedItem?.ToString() ?? "Alle";

            if (timeRange == "Heute")
                return "DATE(Zeitstempel) = DATE('now')";
            else if (timeRange == "Diese Woche")
                return "DATE(Zeitstempel) >= DATE('now', '-7 days')";
            else if (timeRange == "Dieser Monat")
                return "DATE(Zeitstempel) >= DATE('now', '-30 days')";
            else
                return "1=1";
        }

        private void RefreshSoftwareDatabaseView(ComboBox softwareTimeRangeComboBox)
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    string timeRange = softwareTimeRangeComboBox.SelectedItem?.ToString() ?? "Alle";
                    string whereClause = "";

                    if (timeRange == "Heute")
                        whereClause = "DATE(Zeitstempel) = DATE('now')";
                    else if (timeRange == "Diese Woche")
                        whereClause = "DATE(Zeitstempel) >= DATE('now', '-7 days')";
                    else if (timeRange == "Dieser Monat")
                        whereClause = "DATE(Zeitstempel) >= DATE('now', '-30 days')";
                    else
                        whereClause = "1=1";

                    string query = $@"SELECT Zeitstempel, Name, Version, Publisher, InstallLocation, InstallDate 
                                     FROM Software 
                                     WHERE {whereClause}
                                     ORDER BY Zeitstempel DESC";

                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            dbSoftwareTable.Rows.Clear();

                            while (reader.Read())
                            {
                                dbSoftwareTable.Rows.Add(
                                    reader["Zeitstempel"].ToString(),
                                    reader["Name"].ToString(),
                                    reader["Version"].ToString(),
                                    reader["Publisher"].ToString(),
                                    reader["InstallLocation"].ToString(),
                                    reader["InstallDate"].ToString()
                                );
                            }
                        }
                    }
                }

                statusLabel.Text = $"Software geladen: {dbSoftwareTable.Rows.Count} Einträge";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Laden der Software: " + ex.Message);
            }
        }

        // ================= EXPORT =================

        private void ExportData(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "JSON Datei (*.json)|*.json|XML Datei (*.xml)|*.xml",
                Title = "Daten exportieren"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string extension = Path.GetExtension(sfd.FileName).ToLower();

                    if (extension == ".json")
                        ExportToJson(sfd.FileName);
                    else if (extension == ".xml")
                        ExportToXml(sfd.FileName);

                    MessageBox.Show("Daten erfolgreich exportiert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = "Daten exportiert: " + Path.GetFileName(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Exportieren: " + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ExportToJson(string filePath)
        {
            var exportData = new
            {
                Zeitstempel = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                Geräte = currentDevices.Select(d => new
                {
                    d.IP,
                    d.Hostname,
                    d.Status,
                    d.Ports
                }).ToList(),
                Hardware = currentHardwareInfo,
                Software = currentSoftware.Select(s => new
                {
                    s.Name,
                    s.Version,
                    s.Publisher,
                    s.InstallLocation,
                    s.Source
                }).OrderBy(s => s.Name).ToList()
            };

            string json = JsonConvert.SerializeObject(exportData, JsonFormatting.Indented);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        private void ExportToXml(string filePath)
        {
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("Inventar",
                    new XElement("Zeitstempel", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                    new XElement("Geräte",
                        currentDevices.Select(d =>
                            new XElement("Gerät",
                                new XElement("IP", d.IP),
                                new XElement("Hostname", d.Hostname ?? ""),
                                new XElement("Status", d.Status),
                                new XElement("Ports", d.Ports)
                            )
                        )
                    ),
                    new XElement("Hardware", new XCData(currentHardwareInfo)),
                    new XElement("Software",
                        currentSoftware.OrderBy(s => s.Name).Select(s =>
                            new XElement("Programm",
                                new XElement("Name", s.Name),
                                new XElement("Version", s.Version ?? "N/A"),
                                new XElement("Hersteller", s.Publisher ?? ""),
                                new XElement("Installationsort", s.InstallLocation ?? ""),
                                new XElement("Quelle", s.Source ?? "")
                            )
                        )
                    )
                )
            );

            doc.Save(filePath);
        }

        // ================= NMAP =================

        private void StartScan(object sender, EventArgs e)
        {
            scanButton.Enabled = false;
            statusLabel.Text = "Scan läuft...";

            Task.Run(() => ExecuteNmapScan(networkTextBox.Text));
        }

        private void ExecuteNmapScan(string network)
        {
            try
            {
                string nmapPath = FindNmapExecutable();
                if (nmapPath == null)
                    throw new Exception("Nmap nicht gefunden");

                string output;

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = nmapPath,
                    Arguments = "-p 22,80,139,443,445,3306,3389,53,135,515,6668,8009,8008,8080,8443,9000,9100,9220,9999,10025 -T4 --min-hostgroup 256 -oG - " + network,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process p = Process.Start(psi))
                {
                    output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                }

                Invoke(new MethodInvoker(() =>
                {
                    rawOutputTextBox.Text = output;
                    ParseNmapOutput(output);
                    SaveDevicesToDatabase();
                    LoadDevicesFromDatabaseView();
                    statusLabel.Text = "Scan abgeschlossen und in DB gespeichert";
                    scanButton.Enabled = true;
                }));
            }
            catch (Exception ex)
            {
                Invoke(new MethodInvoker(() =>
                {
                    MessageBox.Show(ex.Message);
                    statusLabel.Text = "Fehler beim Scan";
                    scanButton.Enabled = true;
                }));
            }
        }

        private void ParseNmapOutput(string output)
        {
            deviceTable.Rows.Clear();
            currentDevices.Clear();
            Dictionary<string, DeviceInfo> devices = new Dictionary<string, DeviceInfo>();

            foreach (string line in output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
            {
                if (line.StartsWith("Host:"))
                {
                    string[] parts = line.Split('\t');
                    if (parts.Length >= 2)
                    {
                        string hostPart = parts[0].Replace("Host: ", "").Trim();
                        string ip = hostPart.Split(' ')[0];

                        if (!devices.ContainsKey(ip))
                        {
                            devices[ip] = new DeviceInfo { IP = ip };
                        }

                        string hostname = "";
                        if (hostPart.Contains("(") && hostPart.Contains(")"))
                        {
                            int startIdx = hostPart.IndexOf('(') + 1;
                            int endIdx = hostPart.IndexOf(')');
                            hostname = hostPart.Substring(startIdx, endIdx - startIdx).Trim();
                        }

                        if (!string.IsNullOrEmpty(hostname))
                            devices[ip].Hostname = hostname;

                        string status = "Up";
                        int statusIndex = line.IndexOf("Status:");
                        if (statusIndex >= 0)
                        {
                            string statusText = line.Substring(statusIndex + 7).Trim();
                            status = statusText.Split(' ')[0];
                        }
                        devices[ip].Status = status;

                        string portsRaw = "";
                        foreach (string part in parts)
                        {
                            if (part.Contains("Ports:"))
                            {
                                portsRaw = part.Replace("Ports:", "").Trim();
                                break;
                            }
                        }

                        List<string> portEntries = new List<string>();

                        if (!string.IsNullOrEmpty(portsRaw))
                        {
                            string[] portList = portsRaw.Split(',');

                            foreach (string portEntry in portList)
                            {
                                string port = portEntry.Trim();
                                if (!string.IsNullOrEmpty(port))
                                {
                                    string[] portDetails = port.Split('/');
                                    if (portDetails.Length >= 3)
                                    {
                                        string portNumber = portDetails[0];
                                        string portState = portDetails[1];
                                        string protocol = portDetails[2];
                                        string service = "";

                                        if (portDetails.Length > 4 && !string.IsNullOrEmpty(portDetails[4]))
                                            service = portDetails[4];

                                        if (portState.ToLower() == "open")
                                        {
                                            string formatted = $"{portNumber}/{protocol}";
                                            if (!string.IsNullOrEmpty(service))
                                                formatted += $" ({service})";
                                            portEntries.Add(formatted);
                                        }
                                    }
                                }
                            }
                        }

                        string ports = string.Join(", ", portEntries);
                        if (string.IsNullOrEmpty(ports))
                            ports = "-";

                        devices[ip].Ports = ports;
                    }
                }
            }

            foreach (var device in devices.Values)
            {
                currentDevices.Add(device);
                deviceTable.Rows.Add(device.IP, device.Hostname, device.Status, device.Ports);
            }
        }

        // ================= HARDWARE / SOFTWARE =================

        private void StartHardwareQuery(object sender, EventArgs e)
        {
            hardwareButton.Enabled = false;
            statusLabel.Text = "Daten werden gesammelt...";

            Task.Run(() =>
            {
                string hw = GetCompleteHardwareInfo();
                List<SoftwareInfo> sw = GetCompleteInstalledSoftware();

                Invoke(new MethodInvoker(() =>
                {
                    currentHardwareInfo = hw;
                    currentSoftware = sw;
                    hardwareInfoTextBox.Text = hw;

                    StringBuilder swText = new StringBuilder();
                    swText.AppendLine("=== INSTALLIERTE SOFTWARE ===\n");
                    int counter = 1;
                    foreach (var software in sw.OrderBy(s => s.Name))
                    {
                        swText.AppendLine($"{counter}. {software.Name}");
                        if (!string.IsNullOrEmpty(software.Version) && software.Version != "N/A")
                            swText.AppendLine($"   Version: {software.Version}");
                        if (!string.IsNullOrEmpty(software.Publisher))
                            swText.AppendLine($"   Hersteller: {software.Publisher}");
                        if (!string.IsNullOrEmpty(software.InstallLocation))
                            swText.AppendLine($"   Installationsort: {software.InstallLocation}");
                        if (!string.IsNullOrEmpty(software.InstallDate))
                            swText.AppendLine($"   Installiert: {software.InstallDate}");
                        swText.AppendLine();
                        counter++;
                    }
                    swText.AppendLine($"\n========================================");
                    swText.AppendLine($"Insgesamt {sw.Count} Programme gefunden");
                    swText.AppendLine($"========================================");

                    softwareInfoTextBox.Text = swText.ToString();
                    SaveSoftwareToDatabase();
                    statusLabel.Text = "Fertig";
                    hardwareButton.Enabled = true;
                }));
            });
        }

        private void StartRemoteHardwareQuery(object sender, EventArgs e)
        {
            Form remoteForm = new Form
            {
                Text = "Remote Hardware Abfrage",
                Width = 450,
                Height = 200,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true
            };

            Label ipLabel = new Label { Text = "Computer-IP:", Location = new Point(20, 30), AutoSize = true };
            TextBox ipTextBox = new TextBox { Location = new Point(130, 27), Width = 250, Text = "" };

            Label credLabel = new Label { Text = "Benutzername:", Location = new Point(20, 70), AutoSize = true };
            TextBox credTextBox = new TextBox { Location = new Point(130, 67), Width = 250, Text = "" };

            Label passLabel = new Label { Text = "Passwort:", Location = new Point(20, 110), AutoSize = true };
            TextBox passTextBox = new TextBox { Location = new Point(130, 107), Width = 250, UseSystemPasswordChar = true, Text = "" };

            Button queryButton = new Button { Text = "Abfragen", Location = new Point(130, 150), Width = 100, Height = 30 };
            Button cancelButton = new Button { Text = "Abbrechen", Location = new Point(240, 150), Width = 100, Height = 30 };

            cancelButton.Click += (s, e1) => remoteForm.Close();

            queryButton.Click += (s, e1) =>
            {
                if (string.IsNullOrEmpty(ipTextBox.Text))
                {
                    MessageBox.Show("Bitte IP-Adresse eingeben!");
                    return;
                }

                string ip = ipTextBox.Text;
                string username = credTextBox.Text;
                string password = passTextBox.Text;

                remoteForm.Close();
                remoteHardwareButton.Enabled = false;
                statusLabel.Text = "Remote Hardware wird abgefragt...";

                Task.Run(() => GetRemoteHardwareInfo(ip, username, password));
            };

            remoteForm.Controls.Add(ipLabel);
            remoteForm.Controls.Add(ipTextBox);
            remoteForm.Controls.Add(credLabel);
            remoteForm.Controls.Add(credTextBox);
            remoteForm.Controls.Add(passLabel);
            remoteForm.Controls.Add(passTextBox);
            remoteForm.Controls.Add(queryButton);
            remoteForm.Controls.Add(cancelButton);

            remoteForm.ShowDialog(this);
        }

        private void GetRemoteHardwareInfo(string computerIP, string username, string password)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"=== REMOTE HARDWARE INFO: {computerIP} ===\n");

                ConnectionOptions connOptions = new ConnectionOptions();
                if (!string.IsNullOrEmpty(username))
                {
                    connOptions.Authentication = AuthenticationLevel.PacketPrivacy;
                    connOptions.Impersonation = ImpersonationLevel.Impersonate;
                    connOptions.EnablePrivileges = true;
                }

                ManagementScope scope = new ManagementScope($@"\\{computerIP}\root\cimv2", connOptions);

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    scope.Options.Authentication = AuthenticationLevel.PacketPrivacy;
                    ConnectionOptions opts = new ConnectionOptions
                    {
                        Impersonation = ImpersonationLevel.Impersonate,
                        Authentication = AuthenticationLevel.Default,
                        EnablePrivileges = true,
                        Timeout = TimeSpan.FromSeconds(30)
                    };
                    scope.Options = opts;
                }

                scope.Connect();

                // Betriebssystem
                sb.AppendLine("=== BETRIEBSSYSTEM ===");
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Caption, Version FROM Win32_OperatingSystem"));
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        sb.AppendLine("OS: " + obj["Caption"]);
                        sb.AppendLine("Version: " + obj["Version"]);
                    }
                }
                catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
                sb.AppendLine();

                // Prozessor
                sb.AppendLine("=== PROZESSOR ===");
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, Manufacturer, MaxClockSpeed, NumberOfCores FROM Win32_Processor"));
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        sb.AppendLine("Name: " + obj["Name"]);
                        sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                        sb.AppendLine("Taktfrequenz: " + obj["MaxClockSpeed"] + " MHz");
                        sb.AppendLine("Kerne: " + obj["NumberOfCores"]);
                    }
                }
                catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
                sb.AppendLine();

                // RAM
                sb.AppendLine("=== SPEICHER (RAM) ===");
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Manufacturer, Capacity, Speed FROM Win32_PhysicalMemory"));
                    int count = 0;
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        count++;
                        sb.AppendLine("RAM " + count + ":");
                        sb.AppendLine("  Hersteller: " + obj["Manufacturer"]);
                        long bytes = Convert.ToInt64(obj["Capacity"]);
                        sb.AppendLine("  Kapazität: " + (bytes / 1024 / 1024 / 1024) + " GB");
                        sb.AppendLine("  Geschwindigkeit: " + obj["Speed"] + " MHz");
                    }
                }
                catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
                sb.AppendLine();

                // Festplatten
                sb.AppendLine("=== FESTPLATTEN ===");
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, Model, Size FROM Win32_DiskDrive"));
                    int count = 0;
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        count++;
                        sb.AppendLine("Disk " + count + ":");
                        sb.AppendLine("  Name: " + obj["Name"]);
                        sb.AppendLine("  Modell: " + obj["Model"]);
                        long bytes = Convert.ToInt64(obj["Size"]);
                        sb.AppendLine("  Größe: " + (bytes / 1024 / 1024 / 1024) + " GB");
                    }
                }
                catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
                sb.AppendLine();

                // Grafikkarte
                sb.AppendLine("=== GRAFIKKARTE ===");
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, DriverVersion FROM Win32_VideoController"));
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        sb.AppendLine("Name: " + obj["Name"]);
                        sb.AppendLine("Treiber: " + obj["DriverVersion"]);
                    }
                }
                catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }

                Invoke(new MethodInvoker(() =>
                {
                    hardwareInfoTextBox.Text = sb.ToString();
                    statusLabel.Text = "Remote Hardware Abfrage abgeschlossen";
                    remoteHardwareButton.Enabled = true;
                }));
            }
            catch (Exception ex)
            {
                Invoke(new MethodInvoker(() =>
                {
                    MessageBox.Show("Fehler bei Remote Abfrage: " + ex.Message);
                    statusLabel.Text = "Fehler bei Remote Hardware Abfrage";
                    remoteHardwareButton.Enabled = true;
                }));
            }
        }

        private string GetCompleteHardwareInfo()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("=== BETRIEBSSYSTEM ===");
            sb.AppendLine("Computer: " + Environment.MachineName);
            sb.AppendLine("Benutzer: " + Environment.UserName);
            sb.AppendLine("OS: " + Environment.OSVersion.VersionString);
            sb.AppendLine("64 Bit: " + (Environment.Is64BitOperatingSystem ? "Ja" : "Nein"));
            sb.AppendLine();

            sb.AppendLine("=== PROZESSOR ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Name: " + obj["Name"]);
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                    sb.AppendLine("Taktfrequenz: " + obj["MaxClockSpeed"] + " MHz");
                    sb.AppendLine("Kerne: " + obj["NumberOfCores"]);
                    sb.AppendLine("Logische Prozessoren: " + obj["NumberOfLogicalProcessors"]);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== SPEICHER (RAM) ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Manufacturer, Capacity, Speed, PartNumber FROM Win32_PhysicalMemory");
                int count = 0;
                foreach (ManagementObject obj in searcher.Get())
                {
                    count++;
                    sb.AppendLine("RAM " + count + ":");
                    sb.AppendLine("  Hersteller: " + obj["Manufacturer"]);
                    long bytes = Convert.ToInt64(obj["Capacity"]);
                    sb.AppendLine("  Kapazität: " + (bytes / 1024 / 1024 / 1024) + " GB");
                    sb.AppendLine("  Geschwindigkeit: " + obj["Speed"] + " MHz");
                    sb.AppendLine("  Seriennummer: " + obj["PartNumber"]);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== FESTPLATTEN ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, Model, Size, Partitions FROM Win32_DiskDrive");
                int count = 0;
                foreach (ManagementObject obj in searcher.Get())
                {
                    count++;
                    sb.AppendLine("Disk " + count + ":");
                    sb.AppendLine("  Name: " + obj["Name"]);
                    sb.AppendLine("  Modell: " + obj["Model"]);
                    long bytes = Convert.ToInt64(obj["Size"]);
                    sb.AppendLine("  Größe: " + (bytes / 1024 / 1024 / 1024) + " GB");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== LAUFWERKE ===");
            try
            {
                foreach (DriveInfo drive in DriveInfo.GetDrives())
                {
                    if (drive.IsReady)
                    {
                        sb.AppendLine(drive.Name);
                        sb.AppendLine("  Dateisystem: " + drive.DriveFormat);
                        sb.AppendLine("  Gesamt: " + (drive.TotalSize / 1024 / 1024 / 1024) + " GB");
                        sb.AppendLine("  Frei: " + (drive.AvailableFreeSpace / 1024 / 1024 / 1024) + " GB");
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== GRAFIKKARTE ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, DriverVersion, AdapterRAM FROM Win32_VideoController");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Name: " + obj["Name"]);
                    sb.AppendLine("Treiber: " + obj["DriverVersion"]);
                    long ram = Convert.ToInt64(obj["AdapterRAM"] ?? 0);
                    if (ram > 0)
                        sb.AppendLine("VRAM: " + (ram / 1024 / 1024) + " MB");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== NETZWERK-ADAPTER ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, Description, MACAddress, Speed, DriverVersion FROM Win32_NetworkAdapter");
                foreach (ManagementObject obj in searcher.Get())
                {
                    if (obj["Description"] != null)
                    {
                        sb.AppendLine("Name: " + obj["Description"]);
                        sb.AppendLine("MAC: " + obj["MACAddress"]);
                        sb.AppendLine("Geschwindigkeit: " + obj["Speed"]);
                        sb.AppendLine("Treiber: " + obj["DriverVersion"]);
                        sb.AppendLine();
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }
            sb.AppendLine();

            sb.AppendLine("=== TREIBER (Letztes Update) ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Description, DriverVersion, DriverDate FROM Win32_NetworkAdapter");
                foreach (ManagementObject obj in searcher.Get())
                {
                    if (obj["Description"] != null)
                    {
                        sb.AppendLine("Treiber: " + obj["Description"]);
                        sb.AppendLine("Version: " + obj["DriverVersion"]);
                        object driverDate = obj["DriverDate"];
                        if (driverDate != null && !string.IsNullOrEmpty(driverDate.ToString()))
                        {
                            string formattedDate = FormatWmiDate(driverDate.ToString());
                            sb.AppendLine("Letztes Update: " + formattedDate);
                        }
                        sb.AppendLine();
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }

            sb.AppendLine("=== BIOS ===");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Manufacturer, Version, ReleaseDate, SerialNumber FROM Win32_BIOS");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                    sb.AppendLine("Version: " + obj["Version"]);
                    sb.AppendLine("Release Date: " + obj["ReleaseDate"]);
                    sb.AppendLine("Seriennummer: " + obj["SerialNumber"]);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Fehler: " + ex.Message);
            }

            return sb.ToString();
        }

        private List<SoftwareInfo> GetCompleteInstalledSoftware()
        {
            List<SoftwareInfo> allSoftware = new List<SoftwareInfo>();
            Dictionary<string, SoftwareInfo> softwareDict = new Dictionary<string, SoftwareInfo>();

            ReadRegistrySoftware(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", softwareDict, "HKLM 64-Bit");
            ReadRegistrySoftware(Registry.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", softwareDict, "HKLM 32-Bit");
            ReadRegistrySoftware(Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", softwareDict, "HKCU");
            ReadStoreApps(softwareDict);

            return softwareDict.Values.ToList();
        }

        private void ReadRegistrySoftware(RegistryKey rootKey, string registryPath, Dictionary<string, SoftwareInfo> allSoftware, string source)
        {
            try
            {
                using (RegistryKey key = rootKey.OpenSubKey(registryPath))
                {
                    if (key == null) return;

                    foreach (string subKeyName in key.GetSubKeyNames())
                    {
                        try
                        {
                            using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                            {
                                if (subKey == null) continue;

                                string name = subKey.GetValue("DisplayName")?.ToString();
                                if (string.IsNullOrEmpty(name)) continue;

                                if (name.Contains("Update") || name.Contains("Hotfix") || name.StartsWith("KB"))
                                    continue;

                                string version = subKey.GetValue("DisplayVersion")?.ToString() ?? "N/A";
                                string publisher = subKey.GetValue("Publisher")?.ToString() ?? "";
                                string installLocation = subKey.GetValue("InstallLocation")?.ToString() ?? "";

                                string uniqueKey = name + "|" + version;

                                if (!allSoftware.ContainsKey(uniqueKey))
                                {
                                    string installDate = "";
                                    object installDateObj = subKey.GetValue("InstallDate");
                                    if (installDateObj != null && !string.IsNullOrEmpty(installDateObj.ToString()))
                                    {
                                        installDate = FormatInstallDate(installDateObj.ToString());
                                    }

                                    allSoftware[uniqueKey] = new SoftwareInfo
                                    {
                                        Name = name,
                                        Version = version,
                                        Publisher = publisher,
                                        InstallLocation = installLocation,
                                        Source = source,
                                        InstallDate = installDate
                                    };
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void ReadStoreApps(Dictionary<string, SoftwareInfo> allSoftware)
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, Version, Description FROM Win32_AppxPackage WHERE IsFramework=False");
                searcher.Options.Timeout = TimeSpan.FromSeconds(10);

                foreach (ManagementObject obj in searcher.Get())
                {
                    try
                    {
                        string name = obj["Name"]?.ToString();
                        string version = obj["Version"]?.ToString() ?? "N/A";
                        string description = obj["Description"]?.ToString() ?? "";

                        if (!string.IsNullOrEmpty(name))
                        {
                            string uniqueKey = name + "|" + version;
                            if (!allSoftware.ContainsKey(uniqueKey))
                            {
                                allSoftware[uniqueKey] = new SoftwareInfo
                                {
                                    Name = name,
                                    Version = version,
                                    Publisher = description,
                                    InstallLocation = "Microsoft Store",
                                    Source = "Store App"
                                };
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private string FormatWmiDate(string wmiDate)
        {
            try
            {
                if (wmiDate.Length >= 8)
                {
                    string year = wmiDate.Substring(0, 4);
                    string month = wmiDate.Substring(4, 2);
                    string day = wmiDate.Substring(6, 2);
                    return $"{day}.{month}.{year}";
                }
            }
            catch { }
            return "N/A";
        }

        private string FormatInstallDate(string installDate)
        {
            try
            {
                // Format: YYYYMMDD
                if (installDate.Length >= 8 && installDate.All(char.IsDigit))
                {
                    string year = installDate.Substring(0, 4);
                    string month = installDate.Substring(4, 2);
                    string day = installDate.Substring(6, 2);
                    return $"{day}.{month}.{year}";
                }
            }
            catch { }
            return "N/A";
        }

        private string FindNmapExecutable()
        {
            string[] paths =
            {
                @"C:\Program Files\Nmap\nmap.exe",
                @"C:\Program Files (x86)\Nmap\nmap.exe",
                "nmap"
            };

            return paths.FirstOrDefault(File.Exists);
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

        public class DeviceInfo
        {
            public string IP { get; set; }
            public string Hostname { get; set; }
            public string Status { get; set; }
            public string Ports { get; set; }

            public DeviceInfo()
            {
                Hostname = "";
                Status = "Unknown";
                Ports = "-";
            }
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