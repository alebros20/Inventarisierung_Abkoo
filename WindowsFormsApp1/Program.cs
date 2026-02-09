using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
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
            LoadCustomerTree();
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

            var dbDeviceSaveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100 };
            dbDeviceSaveBtn.Click += (s, e) => SaveDatabaseDevices();

            var dbDeviceDeleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = Color.IndianRed };
            dbDeviceDeleteBtn.Click += (s, e) => DeleteDatabaseDeviceRow();

            dbDevicePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true },
                dbDeviceFilter,
                dbDeviceSaveBtn,
                dbDeviceDeleteBtn
            });

            dbDeviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
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

            var dbSoftwareSaveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100 };
            dbSoftwareSaveBtn.Click += (s, e) => SaveDatabaseSoftware();

            var dbSoftwareDeleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = Color.IndianRed };
            dbSoftwareDeleteBtn.Click += (s, e) => DeleteDatabaseSoftwareRow();

            dbSoftwarePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true },
                dbSoftwareFilter,
                dbSoftwareSaveBtn,
                dbSoftwareDeleteBtn
            });

            dbSoftwareTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both };
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
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

            // Kunden/Standorte Verwaltung Tab
            var customerLocationTab = new TabPage("Kunden / Standorte");

            var splitContainer = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 300 };

            // Linke Seite: TreeView
            Panel leftPanel = new Panel { Dock = DockStyle.Fill };
            var treeView = new TreeView { Name = "customerTreeView", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            treeView.AfterSelect += (s, e) => OnTreeNodeSelected(e.Node);

            Panel treeButtonPanel = new Panel { Dock = DockStyle.Top, Height = 40 };
            var addCustomerBtn = new Button { Text = "+ Kunde", Location = new Point(5, 8), Width = 90, BackColor = Color.LightGreen };
            addCustomerBtn.Click += (s, e) => AddCustomer();
            var addLocationBtn = new Button { Text = "+ Standort", Location = new Point(100, 8), Width = 90, BackColor = Color.LightBlue };
            addLocationBtn.Click += (s, e) => AddLocation();
            var deleteNodeBtn = new Button { Text = "Löschen", Location = new Point(195, 8), Width = 90, BackColor = Color.IndianRed };
            deleteNodeBtn.Click += (s, e) => DeleteNode();
            treeButtonPanel.Controls.AddRange(new Control[] { addCustomerBtn, addLocationBtn, deleteNodeBtn });

            leftPanel.Controls.Add(treeView);
            leftPanel.Controls.Add(treeButtonPanel);
            splitContainer.Panel1.Controls.Add(leftPanel);

            // Rechte Seite: Details und IP-Zuordnung
            Panel rightPanel = new Panel { Dock = DockStyle.Fill };

            var detailsLabel = new Label { Name = "detailsLabel", Text = "Details", Location = new Point(10, 10), Font = new Font("Segoe UI", 12, FontStyle.Bold), AutoSize = true };

            var nameLabel = new Label { Text = "Name:", Location = new Point(10, 50), AutoSize = true };
            var nameTextBox = new TextBox { Name = "nameTextBox", Location = new Point(100, 47), Width = 300 };

            var addressLabel = new Label { Text = "Adresse:", Location = new Point(10, 80), AutoSize = true };
            var addressTextBox = new TextBox { Name = "addressTextBox", Location = new Point(100, 77), Width = 300, Multiline = true, Height = 60 };

            var saveDetailsBtn = new Button { Text = "Details speichern", Location = new Point(100, 145), Width = 150, BackColor = Color.LightGreen };
            saveDetailsBtn.Click += (s, e) => SaveNodeDetails();

            var ipLabel = new Label { Text = "Zugeordnete IP-Adressen:", Location = new Point(10, 190), Font = new Font("Segoe UI", 10, FontStyle.Bold), AutoSize = true };

            var ipDataGridView = new DataGridView
            {
                Name = "ipDataGridView",
                Location = new Point(10, 220),
                Width = 600,
                Height = 200,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true
            };
            ipDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Arbeitsplatz", Width = 200 });
            ipDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 150 });
            ipDataGridView.CellDoubleClick += (s, e) => EditIPEntry(e.RowIndex);

            var ipInputLabel = new Label { Text = "Arbeitsplatz:", Location = new Point(10, 430), AutoSize = true };
            var workstationInputTextBox = new TextBox { Name = "workstationInputTextBox", Location = new Point(100, 427), Width = 200 };

            var ipInputLabel2 = new Label { Text = "IP:", Location = new Point(310, 430), AutoSize = true };
            var ipInputTextBox = new TextBox { Name = "ipInputTextBox", Location = new Point(340, 427), Width = 150 };

            var addIpBtn = new Button { Text = "IP hinzufügen", Location = new Point(500, 425), Width = 110, BackColor = Color.LightBlue };
            addIpBtn.Click += (s, e) => AddIPToNode();

            var removeIpBtn = new Button { Text = "IP(s) entfernen", Location = new Point(10, 460), Width = 120, BackColor = Color.IndianRed };
            removeIpBtn.Click += (s, e) => RemoveIPFromNode();

            var importFromDbBtn = new Button { Text = "Aus DB importieren", Location = new Point(140, 460), Width = 150, BackColor = Color.LightGoldenrodYellow };
            importFromDbBtn.Click += (s, e) => ImportIPsFromDatabase();

            rightPanel.Controls.AddRange(new Control[] {
                detailsLabel, nameLabel, nameTextBox, addressLabel, addressTextBox, saveDetailsBtn,
                ipLabel, ipDataGridView, ipInputLabel, workstationInputTextBox, ipInputLabel2, ipInputTextBox,
                addIpBtn, removeIpBtn, importFromDbBtn
            });

            splitContainer.Panel2.Controls.Add(rightPanel);
            customerLocationTab.Controls.Add(splitContainer);
            tabControl.TabPages.Add(customerLocationTab);

            statusLabel = new Label { Dock = DockStyle.Bottom, Height = 25, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle };

            Controls.AddRange(new Control[] { tabControl, topPanel, statusLabel });
        }

        private void StartScan()
        {
            scanButton.Enabled = false;
            statusLabel.Text = "Scan läuft...";
            System.Threading.Tasks.Task.Run(() =>
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
            System.Threading.Tasks.Task.Run(() =>
            {
                string hw = hardwareManager.GetHardwareInfo();
                List<SoftwareInfo> sw = softwareManager.GetInstalledSoftware();
                string pcName = Environment.MachineName;
                foreach (var software in sw)
                {
                    software.PCName = pcName;
                    software.Timestamp = DateTime.Now;
                }
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
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            string hw = hardwareManager.GetRemoteHardwareInfo(form.ComputerIP, form.Username, form.Password);
                            List<SoftwareInfo> sw = softwareManager.GetRemoteSoftware(form.ComputerIP, form.Username, form.Password);
                            string pcName = form.ComputerIP;
                            foreach (var software in sw)
                            {
                                software.PCName = pcName;
                                software.Timestamp = DateTime.Now;
                            }
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
                            updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
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
                dbDeviceTable.Rows.Add(dev.ID, dev.Zeitstempel, dev.IP, dev.Hostname, dev.Status, dev.Ports);
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var software = dbManager.LoadSoftware(filter);
            foreach (var sw in software)
                dbSoftwareTable.Rows.Add(sw.ID, sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
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

        private void SaveDatabaseDevices()
        {
            if (!VerifyPassword("Änderungen speichern"))
                return;

            try
            {
                dbManager.UpdateDevices(GetDevicesFromGrid());
                MessageBox.Show("Änderungen wurden gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseDevices();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Speichern: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteDatabaseDeviceRow()
        {
            if (dbDeviceTable.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bitte wähle eine Zeile zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!VerifyPassword("Zeile löschen"))
                return;

            try
            {
                var selectedRow = dbDeviceTable.SelectedRows[0];
                int id = Convert.ToInt32(selectedRow.Cells[0].Value);

                dbManager.DeleteDevice(id);
                MessageBox.Show("Zeile wurde gelöscht!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseDevices();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Löschen: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveDatabaseSoftware()
        {
            if (!VerifyPassword("Änderungen speichern"))
                return;

            try
            {
                dbManager.UpdateSoftwareEntries(GetSoftwareFromGrid());
                MessageBox.Show("Änderungen wurden gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseSoftware();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Speichern: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteDatabaseSoftwareRow()
        {
            if (dbSoftwareTable.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bitte wähle eine Zeile zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!VerifyPassword("Zeile löschen"))
                return;

            try
            {
                var selectedRow = dbSoftwareTable.SelectedRows[0];
                int id = Convert.ToInt32(selectedRow.Cells[0].Value);

                dbManager.DeleteSoftwareEntry(id);
                MessageBox.Show("Zeile wurde gelöscht!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseSoftware();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Löschen: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool VerifyPassword(string action)
        {
            using (var passwordForm = new PasswordVerificationForm(action))
            {
                return passwordForm.ShowDialog(this) == DialogResult.OK;
            }
        }

        private List<DatabaseDevice> GetDevicesFromGrid()
        {
            var devices = new List<DatabaseDevice>();
            foreach (DataGridViewRow row in dbDeviceTable.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    devices.Add(new DatabaseDevice
                    {
                        ID = Convert.ToInt32(row.Cells[0].Value),
                        Zeitstempel = row.Cells[1].Value?.ToString(),
                        IP = row.Cells[2].Value?.ToString(),
                        Hostname = row.Cells[3].Value?.ToString(),
                        Status = row.Cells[4].Value?.ToString(),
                        Ports = row.Cells[5].Value?.ToString()
                    });
                }
            }
            return devices;
        }

        private List<DatabaseSoftware> GetSoftwareFromGrid()
        {
            var software = new List<DatabaseSoftware>();
            foreach (DataGridViewRow row in dbSoftwareTable.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    software.Add(new DatabaseSoftware
                    {
                        ID = Convert.ToInt32(row.Cells[0].Value),
                        Zeitstempel = row.Cells[1].Value?.ToString(),
                        PCName = row.Cells[2].Value?.ToString(),
                        Name = row.Cells[3].Value?.ToString(),
                        Version = row.Cells[4].Value?.ToString(),
                        Publisher = row.Cells[5].Value?.ToString(),
                        InstallDate = row.Cells[6].Value?.ToString(),
                        LastUpdate = row.Cells[7].Value?.ToString()
                    });
                }
            }
            return software;
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
                            updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
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

        // === KUNDEN/STANDORT VERWALTUNG ===

        private void LoadCustomerTree()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView == null) return;

            treeView.Nodes.Clear();
            var customers = dbManager.GetCustomers();

            foreach (var customer in customers)
            {
                var customerNode = new TreeNode(customer.Name) { Tag = new NodeData { Type = "Customer", ID = customer.ID, Data = customer } };

                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                foreach (var location in locations)
                {
                    var locationNode = new TreeNode(location.Name) { Tag = new NodeData { Type = "Location", ID = location.ID, Data = location } };

                    var ips = dbManager.GetIPsWithWorkstationByLocation(location.ID);
                    foreach (var ip in ips)
                    {
                        string displayText = string.IsNullOrEmpty(ip.WorkstationName)
                            ? $"📍 {ip.IPAddress}"
                            : $"💻 {ip.WorkstationName} ({ip.IPAddress})";
                        var ipNode = new TreeNode(displayText) { Tag = new NodeData { Type = "IP", Data = ip } };
                        locationNode.Nodes.Add(ipNode);
                    }

                    customerNode.Nodes.Add(locationNode);
                }

                treeView.Nodes.Add(customerNode);
            }

            treeView.ExpandAll();
        }

        private void OnTreeNodeSelected(TreeNode node)
        {
            if (node?.Tag == null) return;

            var nodeData = (NodeData)node.Tag;
            var nameTextBox = FindControl<TextBox>("nameTextBox");
            var addressTextBox = FindControl<TextBox>("addressTextBox");
            var ipDataGridView = FindControl<DataGridView>("ipDataGridView");
            var detailsLabel = FindControl<Label>("detailsLabel");

            if (nameTextBox == null || addressTextBox == null || ipDataGridView == null || detailsLabel == null) return;

            ipDataGridView.Rows.Clear();

            if (nodeData.Type == "Customer")
            {
                var customer = (Customer)nodeData.Data;
                detailsLabel.Text = "Kunde bearbeiten";
                nameTextBox.Text = customer.Name;
                addressTextBox.Text = customer.Address;
                nameTextBox.Enabled = true;
                addressTextBox.Enabled = true;
            }
            else if (nodeData.Type == "Location")
            {
                var location = (Location)nodeData.Data;
                detailsLabel.Text = "Standort bearbeiten";
                nameTextBox.Text = location.Name;
                addressTextBox.Text = location.Address;
                nameTextBox.Enabled = true;
                addressTextBox.Enabled = true;

                var ips = dbManager.GetIPsWithWorkstationByLocation(location.ID);
                foreach (var ip in ips)
                    ipDataGridView.Rows.Add(ip.WorkstationName ?? "", ip.IPAddress);
            }
            else if (nodeData.Type == "IP")
            {
                var ip = (LocationIP)nodeData.Data;
                detailsLabel.Text = "IP-Adresse";
                nameTextBox.Text = ip.WorkstationName ?? "";
                addressTextBox.Text = ip.IPAddress;
                nameTextBox.Enabled = false;
                addressTextBox.Enabled = false;
            }
        }

        private void AddCustomer()
        {
            using (var form = new InputDialog("Neuer Kunde", "Kundenname:", "Adresse:"))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("Kunde hinzufügen"))
                        return;

                    dbManager.AddCustomer(form.Value1, form.Value2);
                    LoadCustomerTree();
                    statusLabel.Text = "Kunde hinzugefügt";
                }
            }
        }

        private void AddLocation()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) return;

            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || (nodeData.Type != "Customer" && nodeData.Type != "Location"))
            {
                MessageBox.Show("Bitte wähle zuerst einen Kunden aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int customerId = nodeData.Type == "Customer" ? nodeData.ID : ((Location)nodeData.Data).CustomerID;

            using (var form = new InputDialog("Neuer Standort", "Standortname:", "Adresse:"))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("Standort hinzufügen"))
                        return;

                    dbManager.AddLocation(customerId, form.Value1, form.Value2);
                    LoadCustomerTree();
                    statusLabel.Text = "Standort hinzugefügt";
                }
            }
        }

        private void DeleteNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null)
            {
                MessageBox.Show("Bitte wähle einen Knoten zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null) return;

            if (!VerifyPassword($"{nodeData.Type} löschen"))
                return;

            if (nodeData.Type == "Customer")
            {
                dbManager.DeleteCustomer(nodeData.ID);
                statusLabel.Text = "Kunde gelöscht";
            }
            else if (nodeData.Type == "Location")
            {
                dbManager.DeleteLocation(nodeData.ID);
                statusLabel.Text = "Standort gelöscht";
            }

            LoadCustomerTree();
        }

        private void SaveNodeDetails()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) return;

            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null) return;

            var nameTextBox = FindControl<TextBox>("nameTextBox");
            var addressTextBox = FindControl<TextBox>("addressTextBox");
            if (nameTextBox == null || addressTextBox == null) return;

            if (!VerifyPassword("Details speichern"))
                return;

            if (nodeData.Type == "Customer")
            {
                dbManager.UpdateCustomer(nodeData.ID, nameTextBox.Text, addressTextBox.Text);
                statusLabel.Text = "Kunde aktualisiert";
            }
            else if (nodeData.Type == "Location")
            {
                dbManager.UpdateLocation(nodeData.ID, nameTextBox.Text, addressTextBox.Text);
                statusLabel.Text = "Standort aktualisiert";
            }

            LoadCustomerTree();
        }

        private void AddIPToNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var ipInputTextBox = FindControl<TextBox>("ipInputTextBox");
            var workstationInputTextBox = FindControl<TextBox>("workstationInputTextBox");

            if (treeView?.SelectedNode == null || ipInputTextBox == null || workstationInputTextBox == null) return;

            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location")
            {
                MessageBox.Show("Bitte wähle zuerst einen Standort aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string ip = ipInputTextBox.Text.Trim();
            string workstation = workstationInputTextBox.Text.Trim();

            if (string.IsNullOrEmpty(ip))
            {
                MessageBox.Show("Bitte gib eine IP-Adresse ein!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!VerifyPassword("IP-Adresse hinzufügen"))
                return;

            dbManager.AddIPToLocation(nodeData.ID, ip, workstation);
            ipInputTextBox.Clear();
            workstationInputTextBox.Clear();
            LoadCustomerTree();
            OnTreeNodeSelected(treeView.SelectedNode);
            statusLabel.Text = "IP-Adresse hinzugefügt";
        }

        private void RemoveIPFromNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var ipDataGridView = FindControl<DataGridView>("ipDataGridView");

            if (ipDataGridView?.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bitte wähle mindestens eine IP-Adresse aus der Liste!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location") return;

            if (!VerifyPassword($"{ipDataGridView.SelectedRows.Count} IP-Adresse(n) entfernen"))
                return;

            foreach (DataGridViewRow row in ipDataGridView.SelectedRows)
            {
                string ip = row.Cells[1].Value?.ToString();
                if (!string.IsNullOrEmpty(ip))
                    dbManager.RemoveIPFromLocation(nodeData.ID, ip);
            }

            LoadCustomerTree();
            OnTreeNodeSelected(treeView.SelectedNode);
            statusLabel.Text = $"{ipDataGridView.SelectedRows.Count} IP-Adresse(n) entfernt";
        }

        private void EditIPEntry(int rowIndex)
        {
            if (rowIndex < 0) return;

            var treeView = FindControl<TreeView>("customerTreeView");
            var ipDataGridView = FindControl<DataGridView>("ipDataGridView");

            if (ipDataGridView == null) return;

            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location") return;

            string currentWorkstation = ipDataGridView.Rows[rowIndex].Cells[0].Value?.ToString();
            string currentIP = ipDataGridView.Rows[rowIndex].Cells[1].Value?.ToString();

            using (var form = new InputDialog("IP-Eintrag bearbeiten", "Arbeitsplatz:", "IP-Adresse:"))
            {
                form.Value1 = currentWorkstation;
                form.Value2 = currentIP;

                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("IP-Eintrag bearbeiten"))
                        return;

                    dbManager.RemoveIPFromLocation(nodeData.ID, currentIP);
                    dbManager.AddIPToLocation(nodeData.ID, form.Value2, form.Value1);

                    LoadCustomerTree();
                    OnTreeNodeSelected(treeView.SelectedNode);
                    statusLabel.Text = "IP-Eintrag aktualisiert";
                }
            }
        }

        private void ImportIPsFromDatabase()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) return;

            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location")
            {
                MessageBox.Show("Bitte wähle zuerst einen Standort aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var form = new IPImportDialog(dbManager))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (form.SelectedIPs.Count == 0)
                    {
                        MessageBox.Show("Keine IPs ausgewählt!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (!VerifyPassword($"{form.SelectedIPs.Count} IP(s) importieren"))
                        return;

                    foreach (var ipEntry in form.SelectedIPs)
                    {
                        dbManager.AddIPToLocation(nodeData.ID, ipEntry.Item1, ipEntry.Item2);
                    }

                    LoadCustomerTree();
                    OnTreeNodeSelected(treeView.SelectedNode);
                    statusLabel.Text = $"{form.SelectedIPs.Count} IP(s) importiert";
                }
            }
        }

        private T FindControl<T>(string name) where T : Control
        {
            return Controls.Find(name, true).FirstOrDefault() as T;
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