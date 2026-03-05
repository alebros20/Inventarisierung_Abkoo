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
        // === UI COMPONENTS ===
        private TabControl tabControl;
        private TextBox networkTextBox;
        private Button scanButton, hardwareButton, exportButton, remoteHardwareButton;
        private Label statusLabel;
        private DataGridView deviceTable, softwareGridView, dbDeviceTable, dbSoftwareTable;
        private TextBox rawOutputTextBox, hardwareInfoTextBox;
        private ComboBox locationComboBox;
        private Button newLocationButton;

        // === MANAGERS ===
        private NmapScanner nmapScanner;
        private DatabaseManager dbManager;
        private HardwareManager hardwareManager;
        private SoftwareManager softwareManager;

        // === DATA ===
        private List<DeviceInfo> currentDevices = new List<DeviceInfo>();
        private List<DeviceInfo> lastScanDevices = new List<DeviceInfo>();
        private List<SoftwareInfo> currentSoftware = new List<SoftwareInfo>();
        private string currentRemotePC = "";
        private int selectedLocationID = -1;

        public MainForm()
        {
            nmapScanner = new NmapScanner();
            dbManager = new DatabaseManager();
            hardwareManager = new HardwareManager();
            softwareManager = new SoftwareManager();

            InitializeComponent();
            dbManager.InitializeDatabase();
            LoadCustomerTree();
            ApplyAutoSizing();
        }

        // === INITIALIZATION ===
        private void ApplyAutoSizing()
        {
            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int screenHeight = Screen.PrimaryScreen.WorkingArea.Height;
            int width = Math.Max((int)(screenWidth * 0.9), 1200);
            int height = Math.Max((int)(screenHeight * 0.9), 800);
            Size = new Size(width, height);
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void InitializeComponent()
        {
            Text = "Nmap Inventarisierung";
            Size = new Size(1400, 900);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Font;

            // === PANELS ===
            Panel topPanel = CreateTopPanel();
            tabControl = CreateTabControl();
            statusLabel = CreateStatusLabel();

            Controls.AddRange(new Control[] { tabControl, topPanel, statusLabel });
        }

        // === TOP PANEL ===
        private Panel CreateTopPanel()
        {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = Color.LightGray };

            var networkLabel = new Label { Text = "Netzwerk:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            networkTextBox = new TextBox { Text = "192.168.2.0/24", Location = new Point(100, 8), Width = 200, Height = 28, Font = new Font("Segoe UI", 10) };

            var locationLabel = new Label { Text = "Standort:", Location = new Point(10, 45), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            locationComboBox = new ComboBox { Location = new Point(100, 43), Width = 200, Height = 28, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            locationComboBox.SelectedIndexChanged += (s, e) => OnLocationSelected();

            newLocationButton = new Button { Text = "+ Neuer Standort", Location = new Point(310, 43), Width = 140, Height = 28, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
            newLocationButton.Click += (s, e) => CreateNewLocation();

            scanButton = new Button { Text = "Scan starten", Location = new Point(460, 8), Width = 120, Height = 28, BackColor = Color.LightBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            scanButton.Click += (s, e) => StartScan();

            hardwareButton = new Button { Text = "Hardware/Software", Location = new Point(590, 8), Width = 160, Height = 28, BackColor = Color.LightYellow, Font = new Font("Segoe UI", 10) };
            hardwareButton.Click += (s, e) => StartHardwareQuery();

            exportButton = new Button { Text = "Exportieren", Location = new Point(760, 8), Width = 120, Height = 28, BackColor = Color.LightCyan, Font = new Font("Segoe UI", 10) };
            exportButton.Click += (s, e) => ExportData();

            remoteHardwareButton = new Button { Text = "Remote Hardware", Location = new Point(890, 8), Width = 160, Height = 28, BackColor = Color.LightSalmon, Font = new Font("Segoe UI", 10) };
            remoteHardwareButton.Click += (s, e) => StartRemoteHardwareQuery();

            topPanel.Controls.AddRange(new Control[] { networkLabel, networkTextBox, locationLabel, locationComboBox, newLocationButton, scanButton, hardwareButton, exportButton, remoteHardwareButton });
            return topPanel;
        }

        // === TAB CONTROL ===
        private TabControl CreateTabControl()
        {
            var control = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            control.TabPages.Add(CreateDeviceTab());
            control.TabPages.Add(CreateNmapTab());
            control.TabPages.Add(CreateHardwareTab());
            control.TabPages.Add(CreateSoftwareTab());
            control.TabPages.Add(CreateDBDeviceTab());
            control.TabPages.Add(CreateDBSoftwareTab());
            control.TabPages.Add(CreateCustomerLocationTab());
            return control;
        }

        // === DEVICE TAB ===
        private TabPage CreateDeviceTab()
        {
            Panel devicePanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.WhiteSmoke, Padding = new Padding(10) };
            var legendLabel = new Label { Text = "Legende:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            var greenCircle = new Label { Text = "●", Location = new Point(90, 10), Font = new Font("Arial", 16), ForeColor = Color.Green, AutoSize = true };
            var greenText = new Label { Text = "Aktiv (online)", Location = new Point(110, 12), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var yellowCircle = new Label { Text = "●", Location = new Point(230, 10), Font = new Font("Arial", 16), ForeColor = Color.Gold, AutoSize = true };
            var yellowText = new Label { Text = "Neu seit letztem Scan", Location = new Point(250, 12), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var redCircle = new Label { Text = "●", Location = new Point(430, 10), Font = new Font("Arial", 16), ForeColor = Color.Red, AutoSize = true };
            var redText = new Label { Text = "Offline (fehlend)", Location = new Point(450, 12), AutoSize = true, Font = new Font("Segoe UI", 10) };
            devicePanel.Controls.AddRange(new Control[] { legendLabel, greenCircle, greenText, yellowCircle, yellowText, redCircle, redText });

            deviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10), RowTemplate = { Height = 35 } };
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 50 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 200 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Online Status", Width = 100 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 600 });
            deviceTable.CellPainting += (s, e) =>
            {
                if (e.ColumnIndex == 0 && e.RowIndex >= 0)
                {
                    e.PaintBackground(e.ClipBounds, false);
                    string status = e.Value?.ToString() ?? "";
                    Color circleColor = Color.Green;
                    if (status == "NEU") circleColor = Color.Gold;
                    else if (status == "OFFLINE") circleColor = Color.Red;
                    using (var brush = new SolidBrush(circleColor))
                        e.Graphics.FillEllipse(brush, e.CellBounds.Left + 15, e.CellBounds.Top + 10, 20, 20);
                    e.Handled = true;
                }
            };

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(deviceTable);
            container.Controls.Add(devicePanel);
            return new TabPage("Geräte") { Controls = { container } };
        }

        // === NMAP TAB ===
        private TabPage CreateNmapTab()
        {
            rawOutputTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            return new TabPage("Nmap Ausgabe") { Controls = { rawOutputTextBox } };
        }

        // === HARDWARE TAB ===
        private TabPage CreateHardwareTab()
        {
            Panel hardwareInfoPanel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.LightSteelBlue, Padding = new Padding(10) };
            var pcLabel = new Label { Name = "hardwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue };
            var updateLabel = new Label { Name = "hardwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = Color.DarkSlateGray };
            hardwareInfoPanel.Controls.AddRange(new Control[] { pcLabel, updateLabel });

            hardwareInfoTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(hardwareInfoTextBox);
            container.Controls.Add(hardwareInfoPanel);

            var tabPage = new TabPage("Hardware") { Controls = { container } };
            tabPage.Name = "Hardware";
            return tabPage;
        }

        // === SOFTWARE TAB ===
        private TabPage CreateSoftwareTab()
        {
            Panel softwareInfoPanel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.LightSteelBlue, Padding = new Padding(10) };
            var pcLabel = new Label { Name = "softwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue };
            var updateLabel = new Label { Name = "softwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = Color.DarkSlateGray };
            softwareInfoPanel.Controls.AddRange(new Control[] { pcLabel, updateLabel });

            softwareGridView = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Software", Width = 250, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 120, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert", Width = 120, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Letztes Update", Width = 140, ReadOnly = true });
            softwareGridView.Columns.Add(new DataGridViewButtonColumn { HeaderText = "Aktion", Text = "Update", Width = 80, UseColumnTextForButtonValue = true });
            softwareGridView.CellClick += (s, e) => OnSoftwareUpdateClick(e);

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(softwareGridView);
            container.Controls.Add(softwareInfoPanel);

            var tabPage = new TabPage("Software") { Controls = { container } };
            tabPage.Name = "Software";
            return tabPage;
        }

        // === DB DEVICE TAB ===
        private TabPage CreateDBDeviceTab()
        {
            Panel dbDevicePanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var dbDeviceFilter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            dbDeviceFilter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            dbDeviceFilter.SelectedIndex = 0;
            dbDeviceFilter.SelectedIndexChanged += (s, e) => LoadDatabaseDevices(dbDeviceFilter.SelectedItem.ToString());

            var dbDeviceSaveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            dbDeviceSaveBtn.Click += (s, e) => SaveDatabaseDevices();

            var dbDeviceDeleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = Color.IndianRed, Font = new Font("Segoe UI", 10) };
            dbDeviceDeleteBtn.Click += (s, e) => DeleteDatabaseDeviceRow();

            dbDevicePanel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, dbDeviceFilter, dbDeviceSaveBtn, dbDeviceDeleteBtn });

            dbDeviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 80 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 400 });

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(dbDeviceTable);
            container.Controls.Add(dbDevicePanel);
            return new TabPage("DB - Geräte") { Controls = { container } };
        }

        // === DB SOFTWARE TAB ===
        private TabPage CreateDBSoftwareTab()
        {
            Panel dbSoftwarePanel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var dbSoftwareFilter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            dbSoftwareFilter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            dbSoftwareFilter.SelectedIndex = 0;
            dbSoftwareFilter.SelectedIndexChanged += (s, e) => LoadDatabaseSoftware(dbSoftwareFilter.SelectedItem.ToString());

            var dbSoftwareSaveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            dbSoftwareSaveBtn.Click += (s, e) => SaveDatabaseSoftware();

            var dbSoftwareDeleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = Color.IndianRed, Font = new Font("Segoe UI", 10) };
            dbSoftwareDeleteBtn.Click += (s, e) => DeleteDatabaseSoftwareRow();

            dbSoftwarePanel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, dbSoftwareFilter, dbSoftwareSaveBtn, dbSoftwareDeleteBtn });

            dbSoftwareTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "PC Name/IP", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 200 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert am", Width = 120 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Letztes Update", Width = 140 });

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(dbSoftwareTable);
            container.Controls.Add(dbSoftwarePanel);
            return new TabPage("DB - Software") { Controls = { container } };
        }

        // === CUSTOMER/LOCATION TAB ===
        private TabPage CreateCustomerLocationTab()
        {
            var customerLocationTab = new TabPage("Kunden / Standorte");
            var splitContainer = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 300 };

            Panel leftPanel = new Panel { Dock = DockStyle.Fill };
            var treeView = new TreeView { Name = "customerTreeView", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            treeView.AfterSelect += (s, e) => OnTreeNodeSelected(e.Node);

            Panel treeButtonPanel = new Panel { Dock = DockStyle.Top, Height = 40 };
            var addCustomerBtn = new Button { Text = "+ Kunde", Location = new Point(5, 8), Width = 90, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
            addCustomerBtn.Click += (s, e) => AddCustomer();
            var addLocationBtn = new Button { Text = "+ Standort", Location = new Point(100, 8), Width = 90, BackColor = Color.LightBlue, Font = new Font("Segoe UI", 10) };
            addLocationBtn.Click += (s, e) => AddLocation();
            var addDepartmentBtn = new Button { Text = "+ Abteilung", Location = new Point(195, 8), Width = 90, BackColor = Color.LightCyan, Font = new Font("Segoe UI", 10) };
            addDepartmentBtn.Click += (s, e) => AddDepartment();
            var deleteNodeBtn = new Button { Text = "Löschen", Location = new Point(290, 8), Width = 90, BackColor = Color.IndianRed, Font = new Font("Segoe UI", 10) };
            deleteNodeBtn.Click += (s, e) => DeleteNode();
            treeButtonPanel.Controls.AddRange(new Control[] { addCustomerBtn, addLocationBtn, addDepartmentBtn, deleteNodeBtn });

            leftPanel.Controls.Add(treeView);
            leftPanel.Controls.Add(treeButtonPanel);
            splitContainer.Panel1.Controls.Add(leftPanel);

            Panel rightPanel = CreateDetailsPanel();
            splitContainer.Panel2.Controls.Add(rightPanel);
            customerLocationTab.Controls.Add(splitContainer);
            return customerLocationTab;
        }

        // === DETAILS PANEL ===
        private Panel CreateDetailsPanel()
        {
            Panel rightPanel = new Panel { Dock = DockStyle.Fill };

            var detailsLabel = new Label { Name = "detailsLabel", Text = "Details", Location = new Point(10, 10), Font = new Font("Segoe UI", 12, FontStyle.Bold), AutoSize = true };
            var nameLabel = new Label { Text = "Name:", Location = new Point(10, 50), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var nameTextBox = new TextBox { Name = "nameTextBox", Location = new Point(100, 47), Width = 300, Font = new Font("Segoe UI", 10) };
            var addressLabel = new Label { Text = "Adresse:", Location = new Point(10, 80), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var addressTextBox = new TextBox { Name = "addressTextBox", Location = new Point(100, 77), Width = 300, Multiline = true, Height = 60, Font = new Font("Segoe UI", 10) };
            var saveDetailsBtn = new Button { Text = "Details speichern", Location = new Point(100, 145), Width = 150, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
            saveDetailsBtn.Click += (s, e) => SaveNodeDetails();

            var ipLabel = new Label { Text = "Zugeordnete IP-Adressen:", Location = new Point(10, 190), Font = new Font("Segoe UI", 10, FontStyle.Bold), AutoSize = true };
            var ipDataGridView = new DataGridView { Name = "ipDataGridView", Location = new Point(10, 220), Width = 600, Height = 200, AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = true, Font = new Font("Segoe UI", 10) };
            ipDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Arbeitsplatz", Width = 200 });
            ipDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 150 });
            ipDataGridView.CellDoubleClick += (s, e) => EditIPEntry(e.RowIndex);

            var ipInputLabel = new Label { Text = "Arbeitsplatz:", Location = new Point(10, 430), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var workstationInputTextBox = new TextBox { Name = "workstationInputTextBox", Location = new Point(100, 427), Width = 200, Font = new Font("Segoe UI", 10) };
            var ipInputLabel2 = new Label { Text = "IP:", Location = new Point(310, 430), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var ipInputTextBox = new TextBox { Name = "ipInputTextBox", Location = new Point(340, 427), Width = 150, Font = new Font("Segoe UI", 10) };
            var addIpBtn = new Button { Text = "IP hinzufügen", Location = new Point(500, 425), Width = 110, BackColor = Color.LightBlue, Font = new Font("Segoe UI", 10) };
            addIpBtn.Click += (s, e) => AddIPToNode();
            var removeIpBtn = new Button { Text = "IP(s) entfernen", Location = new Point(10, 460), Width = 120, BackColor = Color.IndianRed, Font = new Font("Segoe UI", 10) };
            removeIpBtn.Click += (s, e) => RemoveIPFromNode();
            var moveIpBtn = new Button { Text = "IP verschieben", Location = new Point(280, 460), Width = 150, BackColor = Color.LightCyan, Font = new Font("Segoe UI", 10) };
            moveIpBtn.Click += (s, e) => MoveIPToAnotherLocation();
            var importFromDbBtn = new Button { Text = "Aus DB importieren", Location = new Point(140, 460), Width = 150, BackColor = Color.LightGoldenrodYellow, Font = new Font("Segoe UI", 10) };
            importFromDbBtn.Click += (s, e) => ImportIPsFromDatabase();

            rightPanel.Controls.AddRange(new Control[] { detailsLabel, nameLabel, nameTextBox, addressLabel, addressTextBox, saveDetailsBtn, ipLabel, ipDataGridView, ipInputLabel, workstationInputTextBox, ipInputLabel2, ipInputTextBox, addIpBtn, removeIpBtn, moveIpBtn, importFromDbBtn });
            return rightPanel;
        }

        // === STATUS LABEL ===
        private Label CreateStatusLabel()
        {
            return new Label { Dock = DockStyle.Bottom, Height = 30, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };
        }

        // === SCAN OPERATIONS ===
        private void StartScan()
        {
            if (selectedLocationID <= 0)
            {
                MessageBox.Show("Bitte wähle einen Standort aus oder erstelle einen neuen!", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            scanButton.Enabled = false;
            statusLabel.Text = "Scan läuft...";
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var oldDevices = currentDevices.ToList();
                    lastScanDevices = oldDevices;
                    var result = nmapScanner.Scan(networkTextBox.Text);

                    Invoke(new MethodInvoker(() =>
                    {
                        rawOutputTextBox.Text = result.RawOutput;
                        currentDevices = result.Devices;
                        SyncDevicesToLocation(selectedLocationID, currentDevices);
                        DisplayDevicesWithStatus(currentDevices, oldDevices);
                        dbManager.SaveDevices(currentDevices);
                        LoadDatabaseDevices();
                        statusLabel.Text = "Scan abgeschlossen und mit Standort abgeglichen";
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

        private void DisplayDevicesWithStatus(List<DeviceInfo> currentDevices, List<DeviceInfo> previousDevices)
        {
            deviceTable.Rows.Clear();
            var displayedIPs = new HashSet<string>();
            var oldIPSet = new HashSet<string>(previousDevices.Select(d => d.IP));
            var newIPSet = new HashSet<string>(currentDevices.Select(d => d.IP));

            foreach (var dev in currentDevices)
            {
                if (!displayedIPs.Contains(dev.IP))
                {
                    string statusText = oldIPSet.Contains(dev.IP) ? "AKTIV" : "NEU";
                    deviceTable.Rows.Add(statusText, dev.IP, dev.Hostname, dev.Status, dev.Ports);
                    displayedIPs.Add(dev.IP);
                }
            }

            foreach (var oldDev in previousDevices)
            {
                if (!newIPSet.Contains(oldDev.IP) && !displayedIPs.Contains(oldDev.IP))
                {
                    deviceTable.Rows.Add("OFFLINE", oldDev.IP, oldDev.Hostname, "Down", "-");
                    displayedIPs.Add(oldDev.IP);
                }
            }
        }

        private void SyncDevicesToLocation(int locationID, List<DeviceInfo> devices)
        {
            var location = dbManager.GetLocationsByCustomer(dbManager.GetCustomers().FirstOrDefault().ID).FirstOrDefault(l => l.ID == locationID);
            if (location == null) return;

            var currentIPs = dbManager.GetIPsWithWorkstationByLocation(locationID);
            var currentIPSet = new HashSet<string>(currentIPs.Select(ip => ip.IPAddress));

            foreach (var device in devices)
            {
                if (!currentIPSet.Contains(device.IP))
                    dbManager.AddIPToLocation(locationID, device.IP, device.Hostname ?? "");
            }

            statusLabel.Text = $"Neue Geräte hinzugefügt: {devices.Count - currentIPSet.Count}";
        }

        // === HARDWARE OPERATIONS ===
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

        private void DisplaySoftwareGrid(List<SoftwareInfo> software, string remotePC = "")
        {
            currentRemotePC = remotePC;
            softwareGridView.Rows.Clear();
            var displayedSoftware = new HashSet<string>();

            var softwareTab = tabControl.TabPages["Software"];
            if (softwareTab != null)
            {
                var container = softwareTab.Controls[0] as Panel;
                var infoPanel = container?.Controls.OfType<Panel>().FirstOrDefault();
                if (infoPanel != null)
                {
                    var pcLabel = infoPanel.Controls["softwarePCLabel"] as Label;
                    var updateLabel = infoPanel.Controls["softwareUpdateLabel"] as Label;
                    if (pcLabel != null) pcLabel.Text = string.IsNullOrEmpty(remotePC) ? $"Lokaler PC: {Environment.MachineName}" : $"Remote-PC: {remotePC}";
                    if (updateLabel != null) updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
                }
            }

            foreach (var sw in software.OrderBy(s => s.Name))
            {
                string uniqueKey = sw.Name + "|" + sw.Version;
                if (!displayedSoftware.Contains(uniqueKey))
                {
                    softwareGridView.Rows.Add(sw.Name, sw.Version ?? "N/A", sw.Publisher ?? "", sw.InstallDate ?? "", sw.LastUpdate ?? "-", "Update");
                    displayedSoftware.Add(uniqueKey);
                }
            }
        }

        // === DATABASE OPERATIONS ===
        private void LoadDatabaseDevices(string filter = "Alle")
        {
            dbDeviceTable.Rows.Clear();
            var devices = dbManager.LoadDevices(filter);
            var displayedIPs = new HashSet<string>();
            foreach (var dev in devices)
            {
                if (!displayedIPs.Contains(dev.IP))
                {
                    dbDeviceTable.Rows.Add(dev.ID, dev.Zeitstempel, dev.IP, dev.Hostname, dev.Status, dev.Ports);
                    displayedIPs.Add(dev.IP);
                }
            }
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var software = dbManager.LoadSoftware(filter);
            var displayedSoftware = new HashSet<string>();
            foreach (var sw in software)
            {
                string uniqueKey = sw.PCName + "|" + sw.Name + "|" + sw.Version;
                if (!displayedSoftware.Contains(uniqueKey))
                {
                    dbSoftwareTable.Rows.Add(sw.ID, sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
                    displayedSoftware.Add(uniqueKey);
                }
            }
        }

        private void SaveDatabaseDevices()
        {
            if (!VerifyPassword("Änderungen speichern")) return;
            try
            {
                dbManager.UpdateDevices(GetDevicesFromGrid());
                MessageBox.Show("Änderungen wurden gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseDevices();
            }
            catch (Exception ex) { MessageBox.Show($"Fehler beim Speichern: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void DeleteDatabaseDeviceRow()
        {
            if (dbDeviceTable.SelectedRows.Count == 0) { MessageBox.Show("Bitte wähle eine Zeile zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (!VerifyPassword("Zeile löschen")) return;
            try
            {
                int id = Convert.ToInt32(dbDeviceTable.SelectedRows[0].Cells[0].Value);
                dbManager.DeleteDevice(id);
                MessageBox.Show("Zeile wurde gelöscht!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseDevices();
            }
            catch (Exception ex) { MessageBox.Show($"Fehler beim Löschen: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void SaveDatabaseSoftware()
        {
            if (!VerifyPassword("Änderungen speichern")) return;
            try
            {
                dbManager.UpdateSoftwareEntries(GetSoftwareFromGrid());
                MessageBox.Show("Änderungen wurden gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseSoftware();
            }
            catch (Exception ex) { MessageBox.Show($"Fehler beim Speichern: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void DeleteDatabaseSoftwareRow()
        {
            if (dbSoftwareTable.SelectedRows.Count == 0) { MessageBox.Show("Bitte wähle eine Zeile zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (!VerifyPassword("Zeile löschen")) return;
            try
            {
                int id = Convert.ToInt32(dbSoftwareTable.SelectedRows[0].Cells[0].Value);
                dbManager.DeleteSoftwareEntry(id);
                MessageBox.Show("Zeile wurde gelöscht!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDatabaseSoftware();
            }
            catch (Exception ex) { MessageBox.Show($"Fehler beim Löschen: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        // === EXPORT ===
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

        // === HELPER METHODS ===
        private bool VerifyPassword(string action)
        {
            using (var passwordForm = new PasswordVerificationForm(action))
                return passwordForm.ShowDialog(this) == DialogResult.OK;
        }

        private List<DatabaseDevice> GetDevicesFromGrid()
        {
            var devices = new List<DatabaseDevice>();
            foreach (DataGridViewRow row in dbDeviceTable.Rows)
                if (row.Cells[0].Value != null)
                    devices.Add(new DatabaseDevice { ID = Convert.ToInt32(row.Cells[0].Value), Zeitstempel = row.Cells[1].Value?.ToString(), IP = row.Cells[2].Value?.ToString(), Hostname = row.Cells[3].Value?.ToString(), Status = row.Cells[4].Value?.ToString(), Ports = row.Cells[5].Value?.ToString() });
            return devices;
        }

        private List<DatabaseSoftware> GetSoftwareFromGrid()
        {
            var software = new List<DatabaseSoftware>();
            foreach (DataGridViewRow row in dbSoftwareTable.Rows)
                if (row.Cells[0].Value != null)
                    software.Add(new DatabaseSoftware { ID = Convert.ToInt32(row.Cells[0].Value), Zeitstempel = row.Cells[1].Value?.ToString(), PCName = row.Cells[2].Value?.ToString(), Name = row.Cells[3].Value?.ToString(), Version = row.Cells[4].Value?.ToString(), Publisher = row.Cells[5].Value?.ToString(), InstallDate = row.Cells[6].Value?.ToString(), LastUpdate = row.Cells[7].Value?.ToString() });
            return software;
        }

        private void UpdateHardwareLabels(string pcName, bool isRemote)
        {
            var hardwareTab = tabControl.TabPages["Hardware"];
            var container = hardwareTab?.Controls[0] as Panel;
            var infoPanel = container?.Controls.OfType<Panel>().FirstOrDefault();
            if (infoPanel != null)
            {
                var pcLabel = infoPanel.Controls["hardwarePCLabel"] as Label;
                var updateLabel = infoPanel.Controls["hardwareUpdateLabel"] as Label;
                if (pcLabel != null) pcLabel.Text = isRemote ? $"Remote-PC: {pcName}" : $"Lokaler PC: {pcName}";
                if (updateLabel != null) updateLabel.Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
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

        // === LOCATION COMBO ===
        private void LoadLocationCombo()
        {
            locationComboBox.Items.Clear();
            locationComboBox.Items.Add(new { ID = -1, Display = "-- Alle Standorte --" });
            var customers = dbManager.GetCustomers();
            foreach (var customer in customers)
            {
                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                foreach (var location in locations)
                    locationComboBox.Items.Add(new { ID = location.ID, Display = $"{customer.Name} > {location.Name}" });
            }
            if (locationComboBox.Items.Count > 0) locationComboBox.SelectedIndex = 0;
        }

        private void OnLocationSelected()
        {
            if (locationComboBox.SelectedItem == null) return;
            dynamic selected = locationComboBox.SelectedItem;
            selectedLocationID = selected.ID;
        }

        private void CreateNewLocation()
        {
            var customers = dbManager.GetCustomers();
            if (customers.Count == 0) { MessageBox.Show("Bitte erstelle zuerst einen Kunden!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new CustomerSelectionForm(customers))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    using (var inputForm = new InputDialog("Neuer Standort", "Standortname:", "Adresse:"))
                    {
                        if (inputForm.ShowDialog(this) == DialogResult.OK)
                        {
                            dbManager.AddLocation(form.SelectedCustomerID, inputForm.Value1, inputForm.Value2);
                            LoadLocationCombo();
                            statusLabel.Text = "Standort erstellt";
                        }
                    }
                }
            }
        }

        // === CUSTOMER/LOCATION TREE ===
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
                    AddLocationNodeRecursive(customerNode, location);
                treeView.Nodes.Add(customerNode);
            }
            treeView.ExpandAll();
            LoadLocationCombo();
        }

        private void AddLocationNodeRecursive(TreeNode parentNode, Location location)
        {
            var locationNode = new TreeNode(location.Name) { Tag = new NodeData { Type = "Location", ID = location.ID, Data = location } };
            var children = dbManager.GetChildLocations(location.ID);
            foreach (var child in children)
                AddLocationNodeRecursive(locationNode, child);
            var ips = dbManager.GetIPsWithWorkstationByLocation(location.ID);
            foreach (var ip in ips)
            {
                string displayText = string.IsNullOrEmpty(ip.WorkstationName) ? $"📍 {ip.IPAddress}" : $"💻 {ip.WorkstationName} ({ip.IPAddress})";
                var ipNode = new TreeNode(displayText) { Tag = new NodeData { Type = "IP", Data = ip } };
                locationNode.Nodes.Add(ipNode);
            }
            parentNode.Nodes.Add(locationNode);
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
                nameTextBox.Enabled = addressTextBox.Enabled = true;
            }
            else if (nodeData.Type == "Location")
            {
                var location = (Location)nodeData.Data;
                string levelName = location.Level == 0 ? "Standort" : "Abteilung";
                detailsLabel.Text = $"{levelName} bearbeiten (Level {location.Level})";
                nameTextBox.Text = location.Name;
                addressTextBox.Text = location.Address;
                nameTextBox.Enabled = addressTextBox.Enabled = true;
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
                nameTextBox.Enabled = addressTextBox.Enabled = false;
            }
        }

        private void AddCustomer()
        {
            using (var form = new InputDialog("Neuer Kunde", "Kundenname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("Kunde hinzufügen")) return;
                    dbManager.AddCustomer(form.Value1, form.Value2);
                    LoadCustomerTree();
                    statusLabel.Text = "Kunde hinzugefügt";
                }
        }

        private void AddLocation()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) return;
            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Customer") { MessageBox.Show("Bitte wähle einen Kunden aus, um einen Standort zu erstellen!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new InputDialog("Neuer Top-Level Standort", "Standortname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("Standort hinzufügen")) return;
                    dbManager.AddLocation(nodeData.ID, form.Value1, form.Value2);
                    LoadCustomerTree();
                    statusLabel.Text = "Standort hinzugefügt";
                }
        }

        private void AddDepartment()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) return;
            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location") { MessageBox.Show("Bitte wähle einen Standort oder eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new InputDialog("Neue Abteilung/Ebene", "Abteilungsname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (!VerifyPassword("Abteilung hinzufügen")) return;
                    dbManager.AddChildLocation(nodeData.ID, form.Value1, form.Value2);
                    LoadCustomerTree();
                    statusLabel.Text = "Abteilung hinzugefügt";
                }
        }

        private void DeleteNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView?.SelectedNode == null) { MessageBox.Show("Bitte wähle einen Knoten zum Löschen aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null) return;
            if (!VerifyPassword($"{nodeData.Type} löschen")) return;
            if (nodeData.Type == "Customer") { dbManager.DeleteCustomer(nodeData.ID); statusLabel.Text = "Kunde gelöscht"; }
            else if (nodeData.Type == "Location") { dbManager.DeleteLocation(nodeData.ID); statusLabel.Text = "Standort gelöscht"; }
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
            if (!VerifyPassword("Details speichern")) return;
            if (nodeData.Type == "Customer") { dbManager.UpdateCustomer(nodeData.ID, nameTextBox.Text, addressTextBox.Text); statusLabel.Text = "Kunde aktualisiert"; }
            else if (nodeData.Type == "Location") { dbManager.UpdateLocation(nodeData.ID, nameTextBox.Text, addressTextBox.Text); statusLabel.Text = "Standort aktualisiert"; }
            LoadCustomerTree();
        }

        private void AddIPToNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var ipInputTextBox = FindControl<TextBox>("ipInputTextBox");
            var workstationInputTextBox = FindControl<TextBox>("workstationInputTextBox");
            if (treeView?.SelectedNode == null || ipInputTextBox == null || workstationInputTextBox == null) return;
            var nodeData = treeView.SelectedNode.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location") { MessageBox.Show("Bitte wähle zuerst einen Standort aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            string ip = ipInputTextBox.Text.Trim();
            string workstation = workstationInputTextBox.Text.Trim();
            if (string.IsNullOrEmpty(ip)) { MessageBox.Show("Bitte gib eine IP-Adresse ein!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!VerifyPassword("IP-Adresse hinzufügen")) return;
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
            if (ipDataGridView?.SelectedRows.Count == 0) { MessageBox.Show("Bitte wähle mindestens eine IP-Adresse aus der Liste!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData == null || nodeData.Type != "Location") return;
            if (!VerifyPassword($"{ipDataGridView.SelectedRows.Count} IP-Adresse(n) entfernen")) return;
            foreach (DataGridViewRow row in ipDataGridView.SelectedRows)
            {
                string ip = row.Cells[1].Value?.ToString();
                if (!string.IsNullOrEmpty(ip)) dbManager.RemoveIPFromLocation(nodeData.ID, ip);
            }
            LoadCustomerTree();
            OnTreeNodeSelected(treeView.SelectedNode);
            statusLabel.Text = $"{ipDataGridView.SelectedRows.Count} IP-Adresse(n) entfernt";
        }
        private void MoveIPToAnotherLocation()
        {
            var ipDataGridView = FindControl<DataGridView>("ipDataGridView");

            if (ipDataGridView?.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bitte wähle eine IP-Adresse zum Verschieben aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;

            if (nodeData == null || nodeData.Type != "Location")
            {
                MessageBox.Show("Bitte wähle einen Standort oder eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int sourceLocationID = nodeData.ID;
            string ipToMove = ipDataGridView.SelectedRows[0].Cells[1].Value?.ToString();
            string workstationName = ipDataGridView.SelectedRows[0].Cells[0].Value?.ToString();

            if (string.IsNullOrEmpty(ipToMove))
            {
                MessageBox.Show("Fehler beim Lesen der IP-Adresse!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var form = new LocationSelectionDialog(dbManager, sourceLocationID))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    int targetLocationID = form.SelectedLocationID;

                    if (!VerifyPassword($"IP {ipToMove} verschieben"))
                        return;

                    try
                    {
                        dbManager.RemoveIPFromLocation(sourceLocationID, ipToMove);
                        dbManager.AddIPToLocation(targetLocationID, ipToMove, workstationName);
                        LoadCustomerTree();

                        var targetLocation = dbManager.GetLocationByID(targetLocationID);
                        statusLabel.Text = $"IP {ipToMove} nach '{targetLocation.Name}' verschoben";

                        MessageBox.Show($"IP-Adresse {ipToMove} erfolgreich nach '{targetLocation.Name}' verschoben!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Fehler beim Verschieben: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
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
                    if (!VerifyPassword("IP-Eintrag bearbeiten")) return;
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
            if (nodeData == null || nodeData.Type != "Location") { MessageBox.Show("Bitte wähle zuerst einen Standort aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new IPImportDialog(dbManager))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    if (form.SelectedIPs.Count == 0) { MessageBox.Show("Keine IPs ausgewählt!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                    if (!VerifyPassword($"{form.SelectedIPs.Count} IP(s) importieren")) return;
                    foreach (var ipEntry in form.SelectedIPs)
                        dbManager.AddIPToLocation(nodeData.ID, ipEntry.Item1, ipEntry.Item2);
                    LoadCustomerTree();
                    OnTreeNodeSelected(treeView.SelectedNode);
                    statusLabel.Text = $"{form.SelectedIPs.Count} IP(s) importiert";
                }
            }
        }

        private T FindControl<T>(string name) where T : Control
            => Controls.Find(name, true).FirstOrDefault() as T;
    }

    public class CustomerSelectionForm : Form
    {
        public int SelectedCustomerID { get; private set; }

        public CustomerSelectionForm(List<Customer> customers)
        {
            Text = "Kunde auswählen";
            Width = 400;
            Height = 200;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var label = new Label { Text = "Wähle einen Kunden aus:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 11) };
            var combo = new ComboBox { Location = new Point(20, 50), Width = 350, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            foreach (var customer in customers)
                combo.Items.Add(new { ID = customer.ID, Name = customer.Name });
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;

            var okBtn = new Button { Text = "OK", Location = new Point(150, 120), Width = 100, DialogResult = DialogResult.OK, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 120), Width = 100, DialogResult = DialogResult.Cancel, Font = new Font("Segoe UI", 10) };

            okBtn.Click += (s, e) =>
            {
                if (combo.SelectedItem != null)
                {
                    dynamic selected = combo.SelectedItem;
                    SelectedCustomerID = selected.ID;
                }
            };

            Controls.AddRange(new Control[] { label, combo, okBtn, cancelBtn });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }
    }
    public class LocationSelectionDialog : Form
    {
        public int SelectedLocationID { get; private set; }
        private DatabaseManager dbManager;
        private int sourceLocationID;
        private TreeView locationTree;

        public LocationSelectionDialog(DatabaseManager dbManager, int sourceLocationID)
        {
            this.dbManager = dbManager;
            this.sourceLocationID = sourceLocationID;

            Text = "Ziel-Location auswählen";
            Width = 400;
            Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Segoe UI", 10);

            var label = new Label
            {
                Text = "Wähle die Ziel-Abteilung/Standort aus:",
                Location = new Point(10, 10),
                AutoSize = true,
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };

            locationTree = new TreeView
            {
                Location = new Point(10, 40),
                Width = 360,
                Height = 380,
                Font = new Font("Segoe UI", 10)
            };

            PopulateLocationTree();

            var okBtn = new Button
            {
                Text = "Verschieben",
                Location = new Point(140, 430),
                Width = 110,
                DialogResult = DialogResult.OK,
                BackColor = Color.LightGreen,
                Font = new Font("Segoe UI", 10)
            };

            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(260, 430),
                Width = 110,
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 10)
            };

            okBtn.Click += (s, e) =>
            {
                if (locationTree.SelectedNode?.Tag is NodeData nodeData && nodeData.Type == "Location")
                {
                    SelectedLocationID = nodeData.ID;
                }
                else
                {
                    MessageBox.Show("Bitte wähle einen Standort oder eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                }
            };

            Controls.AddRange(new Control[] { label, locationTree, okBtn, cancelBtn });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void PopulateLocationTree()
        {
            locationTree.Nodes.Clear();
            var customers = dbManager.GetCustomers();

            foreach (var customer in customers)
            {
                var customerNode = new TreeNode(customer.Name)
                {
                    Tag = new NodeData { Type = "Customer", ID = customer.ID }
                };

                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                foreach (var location in locations)
                {
                    AddLocationNodeRecursive(customerNode, location);
                }

                locationTree.Nodes.Add(customerNode);
            }

            locationTree.ExpandAll();
        }

        private void AddLocationNodeRecursive(TreeNode parentNode, Location location)
        {
            if (location.ID == sourceLocationID)
                return;

            var locationNode = new TreeNode(location.Name)
            {
                Tag = new NodeData { Type = "Location", ID = location.ID }
            };

            var children = dbManager.GetChildLocations(location.ID);
            foreach (var child in children)
            {
                AddLocationNodeRecursive(locationNode, child);
            }

            parentNode.Nodes.Add(locationNode);
        }
    }
    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}