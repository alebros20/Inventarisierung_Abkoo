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

            // Split Container: Links TreeView, Rechts Details
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

            // IP-Adressen Zuordnung
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

                    // Alte IP entfernen und neue hinzufügen
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
        public int ID { get; set; }
        public string Zeitstempel { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
    }

    public class DatabaseSoftware
    {
        public int ID { get; set; }
        public string Zeitstempel { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string PCName { get; set; }
        public string LastUpdate { get; set; }
    }

    public class Customer
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }

    public class Location
    {
        public int ID { get; set; }
        public int CustomerID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }

    public class LocationIP
    {
        public int ID { get; set; }
        public int LocationID { get; set; }
        public string IPAddress { get; set; }
        public string WorkstationName { get; set; }
    }

    public class NodeData
    {
        public string Type { get; set; } // "Customer", "Location", "IP"
        public int ID { get; set; }
        public object Data { get; set; }
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

            sb.AppendLine("=== SYSTEM ===");
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                    sb.AppendLine("Modell: " + obj["Model"]);
                }
            }
            catch (Exception ex) { sb.AppendLine("Fehler: " + ex.Message); }
            sb.AppendLine();

            sb.AppendLine("=== PROZESSOR ===");
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT Name, Manufacturer, MaxClockSpeed FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
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

                sb.AppendLine("=== SYSTEM ===");
                var sysSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Manufacturer, Model FROM Win32_ComputerSystem"));
                foreach (ManagementObject obj in sysSearcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                    sb.AppendLine("Modell: " + obj["Model"]);
                }

                sb.AppendLine("\n=== BETRIEBSSYSTEM ===");
                var osSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Caption, Version, OSArchitecture, Manufacturer FROM Win32_OperatingSystem"));
                foreach (ManagementObject obj in osSearcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
                    sb.AppendLine("OS: " + obj["Caption"]);
                    sb.AppendLine("Version: " + obj["Version"]);
                    sb.AppendLine("Architektur: " + obj["OSArchitecture"]);
                }

                sb.AppendLine("\n=== PROZESSOR ===");
                var cpuSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, Manufacturer, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor"));
                foreach (ManagementObject obj in cpuSearcher.Get())
                {
                    sb.AppendLine("Hersteller: " + obj["Manufacturer"]);
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

                // RAM-Module mit Hersteller
                sb.AppendLine("\n=== RAM-MODULE ===");
                var ramSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Manufacturer, Capacity, Speed, PartNumber FROM Win32_PhysicalMemory"));
                int moduleCount = 0;
                foreach (ManagementObject obj in ramSearcher.Get())
                {
                    moduleCount++;
                    long capacityGB = Convert.ToInt64(obj["Capacity"]) / (1024 * 1024 * 1024);
                    sb.AppendLine($"\nModul {moduleCount}:");
                    sb.AppendLine("  Hersteller: " + (obj["Manufacturer"]?.ToString()?.Trim() ?? "Unbekannt"));
                    sb.AppendLine($"  Kapazität: {capacityGB} GB");
                    sb.AppendLine("  Geschwindigkeit: " + obj["Speed"] + " MHz");
                    sb.AppendLine("  Teilenummer: " + (obj["PartNumber"]?.ToString()?.Trim() ?? "N/A"));
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

                // Physische Festplatten mit Hersteller
                sb.AppendLine("\n=== PHYSISCHE FESTPLATTEN ===");
                var physDiskSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Model, Manufacturer, Size, MediaType, InterfaceType FROM Win32_DiskDrive"));
                int diskCount = 0;
                foreach (ManagementObject obj in physDiskSearcher.Get())
                {
                    diskCount++;
                    long sizeGB = Convert.ToInt64(obj["Size"]) / (1024 * 1024 * 1024);
                    sb.AppendLine($"\nFestplatte {diskCount}:");
                    sb.AppendLine("  Hersteller: " + (obj["Manufacturer"]?.ToString() ?? "Unbekannt"));
                    sb.AppendLine("  Modell: " + obj["Model"]);
                    sb.AppendLine($"  Größe: {sizeGB} GB");
                    sb.AppendLine("  Typ: " + (obj["MediaType"]?.ToString() ?? "N/A"));
                    sb.AppendLine("  Schnittstelle: " + (obj["InterfaceType"]?.ToString() ?? "N/A"));
                }

                sb.AppendLine("\n=== NETZWERKADAPTER ===");
                var netSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Description, Manufacturer, MACAddress, IPEnabled FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True"));
                foreach (ManagementObject obj in netSearcher.Get())
                {
                    sb.AppendLine("\nAdapter: " + obj["Description"]);
                    sb.AppendLine("  Hersteller: " + (obj["Manufacturer"]?.ToString() ?? "Unbekannt"));
                    sb.AppendLine("  MAC: " + (obj["MACAddress"]?.ToString() ?? "N/A"));
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
            var software = new List<SoftwareInfo>();
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-Command \"winget list --accept-source-agreements | Out-String\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    software = ParseWingetOutput(output);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei winget list: {ex.Message}\n\nFallback auf Registry...", "Warnung");
                // Fallback auf Registry-Methode
                var fallbackSoftware = new Dictionary<string, SoftwareInfo>();
                ReadRegistry(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", fallbackSoftware);
                ReadRegistry(Registry.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", fallbackSoftware);
                return fallbackSoftware.Values.ToList();
            }

            return software;
        }

        private List<SoftwareInfo> ParseWingetOutput(string output)
        {
            var software = new List<SoftwareInfo>();
            var lines = output.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            bool dataStarted = false;
            foreach (var line in lines)
            {
                // Überspringe Header-Zeilen bis zur Trennlinie
                if (line.Contains("---"))
                {
                    dataStarted = true;
                    continue;
                }

                if (!dataStarted || line.Trim().Length < 10) continue;

                // Winget-Format: Name   ID   Version   Available   Source
                // Verwende Regex oder feste Positionen
                var parts = System.Text.RegularExpressions.Regex.Split(line, @"\s{2,}");

                if (parts.Length >= 3)
                {
                    string name = parts[0].Trim();
                    string id = parts.Length > 1 ? parts[1].Trim() : "";
                    string version = parts.Length > 2 ? parts[2].Trim() : "";
                    string available = parts.Length > 3 ? parts[3].Trim() : "";

                    if (string.IsNullOrEmpty(name) || name.StartsWith("Name")) continue;

                    var sw = new SoftwareInfo
                    {
                        Name = name,
                        Version = version,
                        Publisher = id, // ID als Publisher verwenden
                        Source = "winget",
                        InstallDate = ""
                    };

                    // Wenn verfügbare Version vorhanden und unterschiedlich
                    if (!string.IsNullOrEmpty(available) && available != version && !available.Contains("<"))
                    {
                        sw.LastUpdate = $"Update verfügbar: {available}";
                    }
                    else
                    {
                        sw.LastUpdate = "Aktuell";
                    }

                    software.Add(sw);
                }
            }

            return software;
        }

        public List<SoftwareInfo> GetRemoteSoftware(string computerIP, string username, string password)
        {
            var software = new List<SoftwareInfo>();

            // Wenn Username/Passwort leer sind, gehe direkt zu WMI (PowerShell Remoting funktioniert nicht ohne Credentials)
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("PowerShell Remoting erfordert Benutzername und Passwort.\n\nVerwende WMI Registry-Abfrage...", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return GetRemoteSoftwareViaWMI(computerIP, username, password);
            }

            try
            {
                // Versuche winget über PowerShell Remoting
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $@"-Command ""
                        $pass = ConvertTo-SecureString '{password}' -AsPlainText -Force
                        $cred = New-Object System.Management.Automation.PSCredential ('{username}', $pass)
                        Invoke-Command -ComputerName {computerIP} -Credential $cred -ScriptBlock {{ winget list --accept-source-agreements }} -ErrorAction Stop | Out-String
                    """,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string errors = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0 && !string.IsNullOrEmpty(output) && output.Contains("Name"))
                    {
                        software = ParseWingetOutput(output);
                        if (software.Count > 0)
                            return software;
                    }

                    // Fehler oder keine Daten
                    throw new Exception($"PowerShell Remoting fehlgeschlagen: {errors}");
                }
            }
            catch (Exception ex)
            {
                // Fallback auf Registry über WMI (ohne Fehlermeldung, da dies erwartet wird)
                return GetRemoteSoftwareViaWMI(computerIP, username, password);
            }
        }

        private List<SoftwareInfo> GetRemoteSoftwareViaWMI(string computerIP, string username, string password)
        {
            var software = new Dictionary<string, SoftwareInfo>();
            try
            {
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

                var registry = new ManagementClass(scope, new ManagementPath("StdRegProv"), null);

                string[] regPaths = new string[]
                {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };

                foreach (string regPath in regPaths)
                {
                    ManagementBaseObject inParams = registry.GetMethodParameters("EnumKey");
                    inParams["hDefKey"] = 0x80000002;
                    inParams["sSubKeyName"] = regPath;

                    ManagementBaseObject outParams = registry.InvokeMethod("EnumKey", inParams, null);
                    if (outParams["sNames"] != null)
                    {
                        string[] subkeys = (string[])outParams["sNames"];

                        foreach (string subkey in subkeys)
                        {
                            string fullPath = regPath + @"\" + subkey;

                            string displayName = GetRegistryValue(registry, fullPath, "DisplayName");
                            if (string.IsNullOrEmpty(displayName) || displayName.Contains("Update")) continue;

                            string version = GetRegistryValue(registry, fullPath, "DisplayVersion") ?? "N/A";
                            string publisher = GetRegistryValue(registry, fullPath, "Publisher") ?? "";
                            string installDate = GetRegistryValue(registry, fullPath, "InstallDate") ?? "";

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
                                    Source = "Remote Registry",
                                    LastUpdate = "-"
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
                                software[uniqueKey] = new SoftwareInfo
                                {
                                    Name = name,
                                    Version = version,
                                    Publisher = sub?.GetValue("Publisher")?.ToString() ?? "",
                                    Source = "Registry",
                                    LastUpdate = "-"
                                };
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
                CREATE TABLE IF NOT EXISTS Software (ID INTEGER PRIMARY KEY, Zeitstempel DATETIME DEFAULT CURRENT_TIMESTAMP, Name TEXT, Version TEXT, Publisher TEXT, InstallLocation TEXT, Source TEXT, InstallDate TEXT, PCName TEXT, LastUpdate TEXT);
                CREATE TABLE IF NOT EXISTS Customers (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT NOT NULL, Address TEXT);
                CREATE TABLE IF NOT EXISTS Locations (ID INTEGER PRIMARY KEY AUTOINCREMENT, CustomerID INTEGER, Name TEXT NOT NULL, Address TEXT, FOREIGN KEY(CustomerID) REFERENCES Customers(ID) ON DELETE CASCADE);
                CREATE TABLE IF NOT EXISTS LocationIPs (ID INTEGER PRIMARY KEY AUTOINCREMENT, LocationID INTEGER, IPAddress TEXT NOT NULL, WorkstationName TEXT, FOREIGN KEY(LocationID) REFERENCES Locations(ID) ON DELETE CASCADE);";
                using (var cmd = new SQLiteCommand(createTableQuery, conn)) cmd.ExecuteNonQuery();

                // Prüfe ob PCName-Spalte existiert
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE Software ADD COLUMN PCName TEXT", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }

                // Prüfe ob LastUpdate-Spalte existiert
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE Software ADD COLUMN LastUpdate TEXT", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }

                // Prüfe ob WorkstationName-Spalte existiert
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE LocationIPs ADD COLUMN WorkstationName TEXT", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }
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
                string query = $"SELECT ID, Zeitstempel, IP, Hostname, Status, Ports FROM Geraete {whereClause} ORDER BY Zeitstempel DESC LIMIT 100";
                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        devices.Add(new DatabaseDevice
                        {
                            ID = Convert.ToInt32(reader["ID"]),
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
                string query = $"SELECT ID, Zeitstempel, Name, Version, Publisher, InstallDate, PCName, LastUpdate FROM Software {whereClause} ORDER BY Zeitstempel DESC LIMIT 100";
                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        software.Add(new DatabaseSoftware
                        {
                            ID = Convert.ToInt32(reader["ID"]),
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

        public void UpdateDevices(List<DatabaseDevice> devices)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                foreach (var dev in devices)
                {
                    using (var cmd = new SQLiteCommand("UPDATE Geraete SET IP=@IP, Hostname=@Hostname, Status=@Status, Ports=@Ports WHERE ID=@ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", dev.ID);
                        cmd.Parameters.AddWithValue("@IP", dev.IP ?? "");
                        cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                        cmd.Parameters.AddWithValue("@Status", dev.Status ?? "");
                        cmd.Parameters.AddWithValue("@Ports", dev.Ports ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public void DeleteDevice(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM Geraete WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void UpdateSoftwareEntries(List<DatabaseSoftware> software)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                foreach (var sw in software)
                {
                    using (var cmd = new SQLiteCommand("UPDATE Software SET PCName=@PCName, Name=@Name, Version=@Version, Publisher=@Publisher, InstallDate=@InstallDate, LastUpdate=@LastUpdate WHERE ID=@ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", sw.ID);
                        cmd.Parameters.AddWithValue("@PCName", sw.PCName ?? "");
                        cmd.Parameters.AddWithValue("@Name", sw.Name ?? "");
                        cmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        cmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        cmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        cmd.Parameters.AddWithValue("@LastUpdate", sw.LastUpdate ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public void DeleteSoftwareEntry(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM Software WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // === KUNDEN/STANDORT VERWALTUNG ===

        public List<Customer> GetCustomers()
        {
            var customers = new List<Customer>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, Name, Address FROM Customers ORDER BY Name", conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        customers.Add(new Customer
                        {
                            ID = Convert.ToInt32(reader["ID"]),
                            Name = reader["Name"].ToString(),
                            Address = reader["Address"]?.ToString() ?? ""
                        });
            }
            return customers;
        }

        public List<Location> GetLocationsByCustomer(int customerId)
        {
            var locations = new List<Location>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, CustomerID, Name, Address FROM Locations WHERE CustomerID=@CustomerID ORDER BY Name", conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            locations.Add(new Location
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                CustomerID = Convert.ToInt32(reader["CustomerID"]),
                                Name = reader["Name"].ToString(),
                                Address = reader["Address"]?.ToString() ?? ""
                            });
                }
            }
            return locations;
        }

        public List<string> GetIPsByLocation(int locationId)
        {
            var ips = new List<string>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT IPAddress FROM LocationIPs WHERE LocationID=@LocationID ORDER BY IPAddress", conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            ips.Add(reader["IPAddress"].ToString());
                }
            }
            return ips;
        }

        public List<LocationIP> GetIPsWithWorkstationByLocation(int locationId)
        {
            var ips = new List<LocationIP>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, LocationID, IPAddress, WorkstationName FROM LocationIPs WHERE LocationID=@LocationID ORDER BY WorkstationName, IPAddress", conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            ips.Add(new LocationIP
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                LocationID = Convert.ToInt32(reader["LocationID"]),
                                IPAddress = reader["IPAddress"].ToString(),
                                WorkstationName = reader["WorkstationName"]?.ToString()
                            });
                }
            }
            return ips;
        }

        public List<(string IP, string Hostname)> GetAllIPsFromDevices()
        {
            var ips = new List<(string, string)>();
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT DISTINCT IP, Hostname FROM Geraete WHERE IP IS NOT NULL AND IP != '' ORDER BY IP", conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                    {
                        string ip = reader["IP"].ToString();
                        string hostname = reader["Hostname"]?.ToString() ?? "";
                        ips.Add((ip, hostname));
                    }
            }
            return ips;
        }

        public void AddCustomer(string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("INSERT INTO Customers (Name, Address) VALUES (@Name, @Address)", conn))
                {
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void AddLocation(int customerId, string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("INSERT INTO Locations (CustomerID, Name, Address) VALUES (@CustomerID, @Name, @Address)", conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerId);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void AddIPToLocation(int locationId, string ip, string workstationName = "")
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("INSERT INTO LocationIPs (LocationID, IPAddress, WorkstationName) VALUES (@LocationID, @IPAddress, @WorkstationName)", conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    cmd.Parameters.AddWithValue("@IPAddress", ip);
                    cmd.Parameters.AddWithValue("@WorkstationName", workstationName ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void UpdateCustomer(int id, string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("UPDATE Customers SET Name=@Name, Address=@Address WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void UpdateLocation(int id, string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("UPDATE Locations SET Name=@Name, Address=@Address WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteCustomer(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM Customers WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteLocation(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM Locations WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void RemoveIPFromLocation(int locationId, string ip)
        {
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM LocationIPs WHERE LocationID=@LocationID AND IPAddress=@IPAddress", conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    cmd.Parameters.AddWithValue("@IPAddress", ip);
                    cmd.ExecuteNonQuery();
                }
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
            Height = 320;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var ipTb = new TextBox { Location = new Point(150, 30), Width = 300 };
            var userTb = new TextBox { Location = new Point(150, 70), Width = 300, Text = Environment.UserName };
            var passTb = new TextBox { Location = new Point(150, 110), Width = 300, UseSystemPasswordChar = true };

            var infoLabel = new Label
            {
                Text = "Hinweis: Für Remote-Zugriff wird ein Passwort benötigt.\nBei leerem Passwort wird nur WMI Registry verwendet.",
                Location = new Point(20, 150),
                Width = 450,
                Height = 50,
                ForeColor = Color.DarkBlue
            };

            var okBtn = new Button { Text = "Verbinden", Location = new Point(150, 210), Width = 100, DialogResult = DialogResult.OK };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 210), Width = 100, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] {
                new Label { Text = "Computer-IP:", Location = new Point(20, 33), AutoSize = true },
                ipTb,
                new Label { Text = "Benutzername:", Location = new Point(20, 73), AutoSize = true },
                userTb,
                new Label { Text = "Passwort:", Location = new Point(20, 113), AutoSize = true },
                passTb,
                infoLabel,
                okBtn, cancelBtn
            });

            AcceptButton = okBtn;
            CancelButton = cancelBtn;

            okBtn.Click += (s, e) =>
            {
                ComputerIP = ipTb.Text.Trim();
                Username = userTb.Text.Trim();
                Password = passTb.Text;

                if (string.IsNullOrEmpty(ComputerIP))
                {
                    MessageBox.Show("Bitte gib eine Computer-IP ein!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };
        }
    }

    public class PasswordVerificationForm : Form
    {
        private const string REQUIRED_PASSWORD = "Administrator";
        private TextBox passwordTextBox;
        private string actionDescription;

        public PasswordVerificationForm(string action)
        {
            actionDescription = action;

            Text = "Sicherheitsabfrage";
            Width = 450;
            Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var infoLabel = new Label
            {
                Text = $"WARNUNG: Du bist dabei '{action}' auszuführen!\n\nGib das Sicherheitspasswort ein:",
                Location = new Point(20, 20),
                Width = 400,
                Height = 60,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkRed
            };

            passwordTextBox = new TextBox
            {
                Location = new Point(20, 90),
                Width = 400,
                Font = new Font("Segoe UI", 11),
                UseSystemPasswordChar = true
            };

            var hintLabel = new Label
            {
                Text = $"Erforderliches Passwort: \"{REQUIRED_PASSWORD}\"",
                Location = new Point(20, 120),
                Width = 400,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Italic)
            };

            var okBtn = new Button
            {
                Text = "Bestätigen",
                Location = new Point(150, 150),
                Width = 100,
                BackColor = Color.LightGreen
            };
            okBtn.Click += OkButton_Click;

            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(260, 150),
                Width = 100,
                DialogResult = DialogResult.Cancel,
                BackColor = Color.LightCoral
            };

            Controls.AddRange(new Control[] { infoLabel, passwordTextBox, hintLabel, okBtn, cancelBtn });

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (passwordTextBox.Text == REQUIRED_PASSWORD)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(
                    $"Falsches Passwort!\n\nDas korrekte Passwort lautet: \"{REQUIRED_PASSWORD}\"",
                    "Zugriff verweigert",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                passwordTextBox.Clear();
                passwordTextBox.Focus();
            }
        }
    }

    public class InputDialog : Form
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }

        private TextBox txt1;
        private TextBox txt2;

        public InputDialog(string title, string label1, string label2)
        {
            Text = title;
            Width = 450;
            Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var lbl1 = new Label { Text = label1, Location = new Point(20, 20), AutoSize = true };
            txt1 = new TextBox { Location = new Point(20, 45), Width = 400 };

            var lbl2 = new Label { Text = label2, Location = new Point(20, 80), AutoSize = true };
            txt2 = new TextBox { Location = new Point(20, 105), Width = 400, Multiline = true, Height = 50 };

            var okBtn = new Button { Text = "OK", Location = new Point(220, 165), Width = 90, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(320, 165), Width = 100, DialogResult = DialogResult.Cancel };

            okBtn.Click += (s, e) => { Value1 = txt1.Text; Value2 = txt2.Text; };

            this.Load += (s, e) =>
            {
                txt1.Text = Value1 ?? "";
                txt2.Text = Value2 ?? "";
            };

            Controls.AddRange(new Control[] { lbl1, txt1, lbl2, txt2, okBtn, cancelBtn });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }
    }

    public class IPImportDialog : Form
    {
        public List<(string IP, string Workstation)> SelectedIPs { get; private set; } = new List<(string, string)>();
        private DataGridView ipGrid;

        public IPImportDialog(DatabaseManager dbManager)
        {
            Text = "IPs aus Datenbank importieren";
            Width = 600;
            Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var infoLabel = new Label
            {
                Text = "Wähle eine oder mehrere IP-Adressen aus der Geräte-Datenbank:",
                Location = new Point(20, 20),
                Width = 550,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            ipGrid = new DataGridView
            {
                Location = new Point(20, 50),
                Width = 550,
                Height = 350,
                AllowUserToAddRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true
            };

            ipGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "✓", Width = 30 });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 150, ReadOnly = true });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname (DB)", Width = 150, ReadOnly = true });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Als Arbeitsplatz", Width = 200 });

            // Lade IPs aus der Datenbank
            var ips = dbManager.GetAllIPsFromDevices();
            foreach (var ip in ips)
            {
                ipGrid.Rows.Add(false, ip.IP, ip.Hostname, ip.Hostname);
            }

            var selectAllBtn = new Button { Text = "Alle auswählen", Location = new Point(20, 410), Width = 120 };
            selectAllBtn.Click += (s, e) =>
            {
                foreach (DataGridViewRow row in ipGrid.Rows)
                    row.Cells[0].Value = true;
            };

            var deselectAllBtn = new Button { Text = "Alle abwählen", Location = new Point(150, 410), Width = 120 };
            deselectAllBtn.Click += (s, e) =>
            {
                foreach (DataGridViewRow row in ipGrid.Rows)
                    row.Cells[0].Value = false;
            };

            var okBtn = new Button { Text = "Importieren", Location = new Point(350, 410), Width = 100, BackColor = Color.LightGreen };
            okBtn.Click += OkButton_Click;

            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(460, 410), Width = 100, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { infoLabel, ipGrid, selectAllBtn, deselectAllBtn, okBtn, cancelBtn });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SelectedIPs.Clear();

            foreach (DataGridViewRow row in ipGrid.Rows)
            {
                bool isChecked = row.Cells[0].Value != null && (bool)row.Cells[0].Value;
                if (isChecked)
                {
                    string ip = row.Cells[1].Value?.ToString();
                    string workstation = row.Cells[3].Value?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(ip))
                        SelectedIPs.Add((ip, workstation));
                }
            }

            DialogResult = DialogResult.OK;
            Close();
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