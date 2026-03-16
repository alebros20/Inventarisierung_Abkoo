using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace NmapInventory
{
    public class MainForm : Form
    {
        // === UI COMPONENTS ===
        private TabControl tabControl;
        private TextBox networkTextBox;
        private Button scanButton, detailScanButton, hardwareButton, exportButton, remoteHardwareButton;
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
        private string currentDisplayedIP = "";  // aktuell im Detail-Panel angezeigte IP

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

        private void ApplyAutoSizing()
        {
            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int screenHeight = Screen.PrimaryScreen.WorkingArea.Height;
            Size = new Size(Math.Max((int)(screenWidth * 0.9), 1200), Math.Max((int)(screenHeight * 0.9), 800));
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void InitializeComponent()
        {
            Text = "Nmap Inventarisierung";
            Size = new Size(1400, 900);
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Font;

            // Felder explizit befüllen — sonst sind tabControl und statusLabel null!
            tabControl = CreateTabControl();
            var topPanel = CreateTopPanel();
            statusLabel = CreateStatusLabel();
            Controls.AddRange(new Control[] { tabControl, topPanel, statusLabel });
        }

        // =========================================================
        // === TOP PANEL ===
        // =========================================================

        private Panel CreateTopPanel()
        {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = Color.LightGray };

            var networkLabel = new Label { Text = "Netzwerk:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            networkTextBox = new TextBox { Text = "192.168.2.0/24", Location = new Point(100, 8), Width = 200, Font = new Font("Segoe UI", 10) };

            var locationLabel = new Label { Text = "Standort:", Location = new Point(10, 45), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            locationComboBox = new ComboBox { Location = new Point(100, 43), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            locationComboBox.SelectedIndexChanged += (s, e) => OnLocationSelected();

            newLocationButton = new Button { Text = "+ Neuer Standort", Location = new Point(310, 43), Width = 140, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            newLocationButton.Click += (s, e) => CreateNewLocation();

            scanButton = new Button { Text = "🔍 Netzwerk scannen", Location = new Point(460, 8), Width = 150, Height = 28, BackColor = Color.LightBlue, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            scanButton.Click += (s, e) => StartScan();

            detailScanButton = new Button { Text = "🔬 Details scannen", Location = new Point(620, 8), Width = 145, Height = 28, BackColor = Color.LightYellow, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            detailScanButton.Click += (s, e) => StartDetailScan();

            hardwareButton = new Button { Text = "Hardware/Software", Location = new Point(775, 8), Width = 160, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            hardwareButton.Click += (s, e) => StartHardwareQuery();

            exportButton = new Button { Text = "Exportieren", Location = new Point(945, 8), Width = 110, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            exportButton.Click += (s, e) => ExportData();

            remoteHardwareButton = new Button { Text = "Remote Hardware", Location = new Point(1065, 8), Width = 150, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            remoteHardwareButton.Click += (s, e) => StartRemoteHardwareQuery();

            topPanel.Controls.AddRange(new Control[] { networkLabel, networkTextBox, locationLabel, locationComboBox, newLocationButton, scanButton, detailScanButton, hardwareButton, exportButton, remoteHardwareButton });
            return topPanel;
        }

        // =========================================================
        // === TAB CONTROL ===
        // =========================================================

        private TabControl CreateTabControl()
        {
            tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            tabControl.TabPages.Add(CreateDeviceTab());
            tabControl.TabPages.Add(CreateNmapTab());
            tabControl.TabPages.Add(CreateHardwareTab());
            tabControl.TabPages.Add(CreateSoftwareTab());
            tabControl.TabPages.Add(CreateDBDeviceTab());
            tabControl.TabPages.Add(CreateDBSoftwareTab());
            tabControl.TabPages.Add(CreateCustomerLocationTab());
            return tabControl;
        }

        private TabPage CreateDeviceTab()
        {
            Panel devicePanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.WhiteSmoke, Padding = new Padding(10) };
            devicePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Legende:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) },
                new Label { Text = "●", Location = new Point(90, 10), Font = new Font("Arial", 16), ForeColor = Color.Green, AutoSize = true },
                new Label { Text = "Aktiv (online)", Location = new Point(110, 12), AutoSize = true, Font = new Font("Segoe UI", 10) },
                new Label { Text = "●", Location = new Point(230, 10), Font = new Font("Arial", 16), ForeColor = Color.Gold, AutoSize = true },
                new Label { Text = "Neu seit letztem Scan", Location = new Point(250, 12), AutoSize = true, Font = new Font("Segoe UI", 10) },
                new Label { Text = "●", Location = new Point(430, 10), Font = new Font("Arial", 16), ForeColor = Color.Red, AutoSize = true },
                new Label { Text = "Offline (fehlend)", Location = new Point(450, 12), AutoSize = true, Font = new Font("Segoe UI", 10) }
            });

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
                    Color circleColor = status == "NEU" ? Color.Gold : status == "OFFLINE" ? Color.Red : Color.Green;
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

        private TabPage CreateNmapTab()
        {
            rawOutputTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            return new TabPage("Nmap Ausgabe") { Controls = { rawOutputTextBox } };
        }

        private TabPage CreateHardwareTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.LightSteelBlue, Padding = new Padding(10) };
            panel.Controls.AddRange(new Control[] {
                new Label { Name = "hardwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue },
                new Label { Name = "hardwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = Color.DarkSlateGray }
            });
            hardwareInfoTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(hardwareInfoTextBox);
            container.Controls.Add(panel);
            return new TabPage("Hardware") { Name = "Hardware", Controls = { container } };
        }

        private TabPage CreateSoftwareTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.LightSteelBlue, Padding = new Padding(10) };
            panel.Controls.AddRange(new Control[] {
                new Label { Name = "softwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue },
                new Label { Name = "softwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = Color.DarkSlateGray }
            });
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
            container.Controls.Add(panel);
            return new TabPage("Software") { Name = "Software", Controls = { container } };
        }

        private TabPage CreateDBDeviceTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var filter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            filter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            filter.SelectedIndex = 0;
            filter.SelectedIndexChanged += (s, e) => LoadDatabaseDevices(filter.SelectedItem.ToString());
            var saveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            saveBtn.Click += (s, e) => SaveDatabaseDevices();
            var deleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            deleteBtn.Click += (s, e) => DeleteDatabaseDeviceRow();
            panel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, filter, saveBtn, deleteBtn });

            dbDeviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 80 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 400 });
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(dbDeviceTable);
            container.Controls.Add(panel);
            return new TabPage("DB - Geräte") { Controls = { container } };
        }

        private TabPage CreateDBSoftwareTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var filter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            filter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            filter.SelectedIndex = 0;
            filter.SelectedIndexChanged += (s, e) => LoadDatabaseSoftware(filter.SelectedItem.ToString());
            var saveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            saveBtn.Click += (s, e) => SaveDatabaseSoftware();
            var deleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            deleteBtn.Click += (s, e) => DeleteDatabaseSoftwareRow();
            panel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, filter, saveBtn, deleteBtn });

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
            container.Controls.Add(panel);
            return new TabPage("DB - Software") { Controls = { container } };
        }

        // =========================================================
        // === KUNDEN / STANDORTE TAB ===
        // =========================================================

        private TabPage CreateCustomerLocationTab()
        {
            var tab = new TabPage("Kunden / Standorte");
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 320 };

            // --- Linke Seite: TreeView + Buttons ---
            var treeView = new TreeView { Name = "customerTreeView", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            treeView.AfterSelect += (s, e) => OnTreeNodeSelected(e.Node);
            treeView.NodeMouseDoubleClick += (s, e) =>
            {
                var nd = e.Node?.Tag as NodeData;
                if (nd?.Type == "IP" && nd.Data is LocationIP lip)
                    ShowDeviceDetails(lip.IPAddress);
            };

            Panel treeButtonPanel = new Panel { Dock = DockStyle.Top, Height = 40 };
            var addCustomerBtn = CreateButton("+ Kunde", 5, Color.LightGreen, AddCustomer);
            var addLocationBtn = CreateButton("+ Standort", 100, Color.LightBlue, AddLocation);
            var addDeptBtn = CreateButton("+ Abteilung", 195, Color.LightCyan, AddDepartment);
            var deleteBtn = CreateButton("Löschen", 290, Color.IndianRed, DeleteNode);
            treeButtonPanel.Controls.AddRange(new Control[] { addCustomerBtn, addLocationBtn, addDeptBtn, deleteBtn });

            var leftPanel = new Panel { Dock = DockStyle.Fill };
            leftPanel.Controls.Add(treeView);
            leftPanel.Controls.Add(treeButtonPanel);
            split.Panel1.Controls.Add(leftPanel);

            // --- Rechte Seite: Details ---
            split.Panel2.Controls.Add(CreateDetailsPanel());
            tab.Controls.Add(split);
            return tab;
        }

        private Button CreateButton(string text, int x, Color color, Action onClick)
        {
            var btn = new Button { Text = text, Location = new Point(x, 8), Width = 90, Font = new Font("Segoe UI", 10) };
            btn.Click += (s, e) => onClick();
            return btn;
        }

        private Panel CreateDetailsPanel()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

            // --- Kunde / Standort Details ---
            var detailsLabel = new Label { Name = "detailsLabel", Text = "Details", Location = new Point(10, 10), Font = new Font("Segoe UI", 12, FontStyle.Bold), AutoSize = true };
            var nameLabel = new Label { Text = "Name:", Location = new Point(10, 50), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var nameTextBox = new TextBox { Name = "nameTextBox", Location = new Point(100, 47), Width = 300, Font = new Font("Segoe UI", 10) };
            var addressLabel = new Label { Text = "Adresse:", Location = new Point(10, 80), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var addressTextBox = new TextBox { Name = "addressTextBox", Location = new Point(100, 77), Width = 300, Multiline = true, Height = 60, Font = new Font("Segoe UI", 10) };
            var saveBtn = new Button { Text = "Details speichern", Location = new Point(100, 145), Width = 150, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            saveBtn.Click += (s, e) => SaveNodeDetails();

            // --- IP Tabelle (für Location-Knoten) ---
            var ipLabel = new Label { Text = "Zugeordnete IP-Adressen:", Location = new Point(10, 190), Font = new Font("Segoe UI", 10, FontStyle.Bold), AutoSize = true };
            var ipGrid = new DataGridView
            {
                Name = "ipDataGridView",
                Location = new Point(10, 215),
                Width = 620,
                Height = 200,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                Font = new Font("Segoe UI", 10),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Arbeitsplatz", Width = 200 });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 150 });
            ipGrid.CellDoubleClick += (s, e) => EditIPEntry(e.RowIndex);

            // --- Manuelle IP Eingabe ---
            var y = 425;
            var wsLabel = new Label { Text = "Arbeitsplatz:", Location = new Point(10, y + 3), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var wsBox = new TextBox { Name = "workstationInputTextBox", Location = new Point(110, y), Width = 180, Font = new Font("Segoe UI", 10) };
            var ipLbl2 = new Label { Text = "IP:", Location = new Point(300, y + 3), AutoSize = true, Font = new Font("Segoe UI", 10) };
            var ipBox = new TextBox { Name = "ipInputTextBox", Location = new Point(320, y), Width = 140, Font = new Font("Segoe UI", 10) };
            var addIpBtn = new Button { Text = "Manuell hinzufügen", Location = new Point(470, y - 1), Width = 155, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            addIpBtn.Click += (s, e) => AddIPToNode();

            y += 35;
            var removeBtn = new Button { Text = "IP(s) entfernen", Location = new Point(10, y), Width = 130, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            removeBtn.Click += (s, e) => RemoveIPFromNode();
            var moveBtn = new Button { Text = "IP verschieben", Location = new Point(150, y), Width = 130, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            moveBtn.Click += (s, e) => MoveIPToAnotherLocation();
            var importBtn = new Button { Text = "📥 Aus DB importieren", Location = new Point(290, y), Width = 160, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            importBtn.Click += (s, e) => ImportIPsFromDatabase();

            // =========================================================
            // === GERÄTE-DETAILS (sichtbar wenn IP-Knoten gewählt) ===
            // =========================================================
            var devicePanel = new Panel
            {
                Name = "deviceDetailPanel",
                Location = new Point(0, 500),
                Width = 650,
                Height = 500,
                Visible = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var deviceTitleLabel = new Label
            {
                Name = "deviceTitleLabel",
                Text = "💻 Geräteinformationen",
                Location = new Point(10, 5),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true
            };

            var refreshDeviceBtn = new Button
            {
                Text = "🔄 Aktualisieren",
                Location = new Point(480, 2),
                Width = 150,
                Height = 26,
                BackColor = Color.LightSteelBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            refreshDeviceBtn.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(currentDisplayedIP))
                    ShowDeviceDetails(currentDisplayedIP);
            };

            // Gerät Info-Zeile (IP / MAC / Hostname)
            var deviceInfoLabel = new Label
            {
                Name = "deviceInfoLabel",
                Text = "",
                Location = new Point(10, 30),
                Size = new Size(630, 20),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };

            // --- Hardware Tab ---
            var deviceTabs = new TabControl
            {
                Name = "deviceTabControl",
                Location = new Point(10, 55),
                Width = 630,
                Height = 430,
                Font = new Font("Segoe UI", 10)
            };

            // Hardware-Seite
            var hwPage = new TabPage("🖥 Hardware");
            var hwBox = new TextBox
            {
                Name = "deviceHardwareBox",
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9),
                BackColor = Color.WhiteSmoke
            };
            hwPage.Controls.Add(hwBox);

            // Software-Seite
            var swPage = new TabPage("📦 Software");
            var swGrid = new DataGridView
            {
                Name = "deviceSoftwareGrid",
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Software", DataPropertyName = "Name", FillWeight = 40 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", DataPropertyName = "Version", FillWeight = 20 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", DataPropertyName = "Publisher", FillWeight = 25 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert", DataPropertyName = "InstallDate", FillWeight = 15 });
            swPage.Controls.Add(swGrid);

            deviceTabs.TabPages.Add(hwPage);
            deviceTabs.TabPages.Add(swPage);

            devicePanel.Controls.AddRange(new Control[] { deviceTitleLabel, refreshDeviceBtn, deviceInfoLabel, deviceTabs });

            panel.Controls.AddRange(new Control[] {
                detailsLabel, nameLabel, nameTextBox, addressLabel, addressTextBox, saveBtn,
                ipLabel, ipGrid,
                wsLabel, wsBox, ipLbl2, ipBox, addIpBtn,
                removeBtn, moveBtn, importBtn,
                devicePanel
            });
            return panel;
        }

        private Label CreateStatusLabel()
            => new Label { Dock = DockStyle.Bottom, Height = 30, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };

        // =========================================================
        // === SCAN ===
        // =========================================================

        /// <summary>
        /// 🔍 NETZWERK SCANNEN — ARP-Discovery, schnell (~5 Sek)
        /// Findet ALLE Geräte im Subnetz inkl. MAC-Adressen.
        /// Befehl: nmap -sn -PR --send-eth
        /// </summary>
        private void StartScan()
        {
            if (selectedLocationID <= 0) { MessageBox.Show("Bitte wähle einen Standort aus!", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            scanButton.Enabled = false;
            detailScanButton.Enabled = false;
            statusLabel.Text = "🔍 Netzwerk-Discovery läuft (ARP)...";

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var oldDevices = currentDevices.ToList();
                    lastScanDevices = oldDevices;
                    var result = nmapScanner.DiscoveryScan(networkTextBox.Text);
                    Invoke(new MethodInvoker(() =>
                    {
                        rawOutputTextBox.Text = result.RawOutput;
                        currentDevices = result.Devices;
                        dbManager.SaveDevices(currentDevices);
                        SyncDevicesToLocation(selectedLocationID, currentDevices);
                        DisplayDevicesWithStatus(currentDevices, oldDevices);
                        LoadDatabaseDevices();

                        if (!string.IsNullOrEmpty(currentDisplayedIP))
                            ShowDeviceDetails(currentDisplayedIP);

                        int found = currentDevices.Count;
                        int withMac = currentDevices.Count(d => !string.IsNullOrEmpty(d.MacAddress));
                        statusLabel.Text = $"✅ Discovery abgeschlossen — {found} Geräte gefunden, {withMac} mit MAC-Adresse";
                        scanButton.Enabled = true;
                        detailScanButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new MethodInvoker(() =>
                    {
                        MessageBox.Show($"Fehler: {ex.Message}");
                        statusLabel.Text = "Fehler beim Scan";
                        scanButton.Enabled = true;
                        detailScanButton.Enabled = true;
                    }));
                }
            });
        }

        /// <summary>
        /// 🔬 DETAILS SCANNEN — Service-Versionen, OS, Ports, Banner
        /// Läuft auf bereits bekannten Geräten aus dem letzten Discovery-Scan.
        /// Befehl: nmap -sV -O -sC -T4 --open
        /// </summary>
        private void StartDetailScan()
        {
            var targets = currentDevices.Select(d => d.IP).ToList();
            if (targets.Count == 0)
            {
                MessageBox.Show("Bitte zuerst einen Netzwerk-Scan durchführen!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            scanButton.Enabled = false;
            detailScanButton.Enabled = false;
            statusLabel.Text = $"🔬 Detail-Scan läuft für {targets.Count} Geräte (kann mehrere Minuten dauern)...";

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    // Alle bekannten IPs als Ziel übergeben
                    string targetList = string.Join(" ", targets);
                    var result = nmapScanner.DetailScan(targetList);

                    Invoke(new MethodInvoker(() =>
                    {
                        rawOutputTextBox.Text = result.RawOutput;

                        // Detail-Infos (OS, Ports, Banner) in DB speichern
                        // Discovery-Geräte mit Detail-Infos zusammenführen
                        foreach (var detailed in result.Devices)
                        {
                            dbManager.SaveNmapDetails(detailed);

                            // Hostname aktualisieren falls jetzt bekannt
                            var existing = currentDevices.FirstOrDefault(d => d.IP == detailed.IP);
                            if (existing != null)
                            {
                                existing.OS = detailed.OS;
                                existing.OSDetails = detailed.OSDetails;
                                existing.OpenPorts = detailed.OpenPorts;
                                existing.Ports = detailed.Ports;
                            }
                        }

                        dbManager.SaveDevices(currentDevices);
                        LoadDatabaseDevices();

                        if (!string.IsNullOrEmpty(currentDisplayedIP))
                            ShowDeviceDetails(currentDisplayedIP);

                        int withOS = result.Devices.Count(d => !string.IsNullOrEmpty(d.OS));
                        int withPorts = result.Devices.Count(d => d.OpenPorts?.Count > 0);
                        statusLabel.Text = $"✅ Detail-Scan abgeschlossen — {withOS} OS erkannt, {withPorts} Geräte mit offenen Ports";
                        scanButton.Enabled = true;
                        detailScanButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new MethodInvoker(() =>
                    {
                        MessageBox.Show($"Fehler: {ex.Message}");
                        statusLabel.Text = "Fehler beim Detail-Scan";
                        scanButton.Enabled = true;
                        detailScanButton.Enabled = true;
                    }));
                }
            });
        }

        private void SyncDevicesToLocation(int locationID, List<DeviceInfo> devices)
        {
            var location = dbManager.GetLocationByID(locationID);
            if (location == null) { statusLabel.Text = "Fehler: Standort nicht gefunden!"; return; }

            // Bereits zugewiesene IPs und MACs ermitteln
            var currentIPs = new HashSet<string>(
                dbManager.GetIPsWithWorkstationByLocation(locationID).Select(ip => ip.IPAddress));

            // Alle MACs die bereits in Devices-Tabelle unter dieser Location stecken
            var allLocationDevices = dbManager.GetAllIPsFromDevices();
            var assignedMACs = new HashSet<string>(
                allLocationDevices
                    .Join(dbManager.GetIPsWithWorkstationByLocation(locationID),
                          d => d.IP, l => l.IPAddress, (d, l) => d)
                    .Select(d => dbManager.LoadDevices("Alle")
                        .FirstOrDefault(x => x.IP == d.IP)?.MacAddress ?? "")
                    .Where(m => !string.IsNullOrEmpty(m)));

            int added = 0, skippedIP = 0, skippedMAC = 0;

            foreach (var device in devices)
            {
                // Gleiche IP bereits vorhanden → überspringen
                if (currentIPs.Contains(device.IP))
                {
                    skippedIP++;
                    continue;
                }

                // Gleiche MAC bereits vorhanden → nur LastSeen aktualisieren, nicht neu anlegen
                if (!string.IsNullOrEmpty(device.MacAddress) && assignedMACs.Contains(device.MacAddress))
                {
                    skippedMAC++;
                    continue;
                }

                dbManager.AddIPToLocation(locationID, device.IP, device.Hostname ?? "");
                added++;
            }

            var msg = $"{added} neue Geräte zu '{location.Name}' hinzugefügt";
            if (skippedIP > 0) msg += $", {skippedIP} IP bereits vorhanden";
            if (skippedMAC > 0) msg += $", {skippedMAC} MAC bereits bekannt (nur LastSeen aktualisiert)";
            statusLabel.Text = msg;
        }

        private void DisplayDevicesWithStatus(List<DeviceInfo> currentDevices, List<DeviceInfo> previousDevices)
        {
            deviceTable.Rows.Clear();
            var displayed = new HashSet<string>();
            var locationIPs = dbManager.GetIPsWithWorkstationByLocation(selectedLocationID);
            var lastScanIPs = new HashSet<string>(dbManager.LoadDevices("Alle").Where(d => locationIPs.Any(l => l.IPAddress == d.IP)).GroupBy(d => d.IP).Select(g => g.First().IP));
            var currentIPs = new HashSet<string>(currentDevices.Select(d => d.IP));
            foreach (var dev in currentDevices)
                if (displayed.Add(dev.IP))
                    deviceTable.Rows.Add(lastScanIPs.Contains(dev.IP) ? "AKTIV" : "NEU", dev.IP, dev.Hostname, dev.Status, dev.Ports);
            foreach (var last in dbManager.LoadDevices("Alle").Where(d => locationIPs.Any(l => l.IPAddress == d.IP)).GroupBy(d => d.IP).Select(g => g.First()))
                if (!currentIPs.Contains(last.IP) && displayed.Add(last.IP))
                    deviceTable.Rows.Add("OFFLINE", last.IP, last.Hostname ?? "", "Down", "-");
        }

        // =========================================================
        // === HARDWARE / SOFTWARE ===
        // =========================================================

        private void StartHardwareQuery()
        {
            hardwareButton.Enabled = false;
            statusLabel.Text = "Daten werden gesammelt...";
            System.Threading.Tasks.Task.Run(() =>
            {
                string hw = hardwareManager.GetHardwareInfo();
                var sw = softwareManager.GetInstalledSoftware();
                string pcName = Environment.MachineName;
                foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
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
                            var sw = softwareManager.GetRemoteSoftware(form.ComputerIP, form.Username, form.Password);
                            string pcName = form.ComputerIP;
                            foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
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
                            Invoke(new MethodInvoker(() => { MessageBox.Show($"Fehler: {ex.Message}"); statusLabel.Text = "Fehler bei Remote Abfrage"; remoteHardwareButton.Enabled = true; }));
                        }
                    });
                }
            }
        }

        private void DisplaySoftwareGrid(List<SoftwareInfo> software, string remotePC = "")
        {
            currentRemotePC = remotePC;
            softwareGridView.Rows.Clear();
            var displayed = new HashSet<string>();
            var softwareTab = tabControl.TabPages["Software"];
            var infoPanel = (softwareTab?.Controls[0] as Panel)?.Controls.OfType<Panel>().FirstOrDefault();
            if (infoPanel != null)
            {
                (infoPanel.Controls["softwarePCLabel"] as Label).Text = string.IsNullOrEmpty(remotePC) ? $"Lokaler PC: {Environment.MachineName}" : $"Remote-PC: {remotePC}";
                (infoPanel.Controls["softwareUpdateLabel"] as Label).Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
            }
            foreach (var sw in software.OrderBy(s => s.Name))
            {
                string key = sw.Name + "|" + sw.Version;
                if (displayed.Add(key))
                    softwareGridView.Rows.Add(sw.Name, sw.Version ?? "N/A", sw.Publisher ?? "", sw.InstallDate ?? "", sw.LastUpdate ?? "-", "Update");
            }
        }

        // =========================================================
        // === DB LADEN / SPEICHERN ===
        // =========================================================

        private void LoadDatabaseDevices(string filter = "Alle")
        {
            dbDeviceTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var dev in dbManager.LoadDevices(filter))
                if (displayed.Add(dev.IP))
                    dbDeviceTable.Rows.Add(dev.ID, dev.Zeitstempel, dev.IP, dev.Hostname, dev.MacAddress ?? "", dev.Status, dev.Ports);
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var sw in dbManager.LoadSoftware(filter))
            {
                string key = sw.PCName + "|" + sw.Name + "|" + sw.Version;
                if (displayed.Add(key))
                    dbSoftwareTable.Rows.Add(sw.ID, sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
            }
        }

        private void SaveDatabaseDevices()
        {
            try { dbManager.UpdateDevices(GetDevicesFromGrid()); MessageBox.Show("Gespeichert!", "Erfolg"); LoadDatabaseDevices(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void DeleteDatabaseDeviceRow()
        {
            if (dbDeviceTable.SelectedRows.Count == 0) return;
            try { dbManager.DeleteDevice(Convert.ToInt32(dbDeviceTable.SelectedRows[0].Cells[0].Value)); LoadDatabaseDevices(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void SaveDatabaseSoftware()
        {
            try { dbManager.UpdateSoftwareEntries(GetSoftwareFromGrid()); MessageBox.Show("Gespeichert!", "Erfolg"); LoadDatabaseSoftware(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void DeleteDatabaseSoftwareRow()
        {
            if (dbSoftwareTable.SelectedRows.Count == 0) return;
            try { dbManager.DeleteSoftwareEntry(Convert.ToInt32(dbSoftwareTable.SelectedRows[0].Cells[0].Value)); LoadDatabaseSoftware(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void ExportData()
        {
            using (var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json|XML (*.xml)|*.xml" })
                if (sfd.ShowDialog() == DialogResult.OK)
                    try { new DataExporter().Export(sfd.FileName, currentDevices, currentSoftware, hardwareInfoTextBox.Text); MessageBox.Show("Exportiert!", "Erfolg"); }
                    catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        // =========================================================
        // === CUSTOMER / LOCATION TREE ===
        // =========================================================

        private void LoadCustomerTree()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            if (treeView == null) return;
            treeView.Nodes.Clear();
            foreach (var customer in dbManager.GetCustomers())
            {
                var customerNode = new TreeNode($"👤 {customer.Name}")
                {
                    Tag = new NodeData { Type = "Customer", ID = customer.ID, Data = customer }
                };
                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                for (int i = 0; i < locations.Count; i++)
                    AddLocationNodeRecursive(customerNode, locations[i], $"S{i + 1}");
                treeView.Nodes.Add(customerNode);
            }
            treeView.ExpandAll();
            LoadLocationCombo();
        }

        private void AddLocationNodeRecursive(TreeNode parentNode, Location location, string shortID)
        {
            string icon = location.Level == 0 ? "🏢" : "📂";
            var locationNode = new TreeNode($"[{shortID}] {icon} {location.Name}")
            {
                Tag = new NodeData { Type = "Location", ID = location.ID, Data = location }
            };

            var children = dbManager.GetChildLocations(location.ID);
            for (int i = 0; i < children.Count; i++)
            {
                // Standort:       S1, S2, S3
                // Abteilung:      S1.A1, S1.A2
                // Unterabteilung: S1.A1.1, S1.A1.2
                string childID = location.Level == 0
                    ? $"{shortID}.A{i + 1}"
                    : $"{shortID}.{i + 1}";
                AddLocationNodeRecursive(locationNode, children[i], childID);
            }

            foreach (var ip in dbManager.GetIPsWithWorkstationByLocation(location.ID))
            {
                string display = string.IsNullOrEmpty(ip.WorkstationName)
                    ? $"📍 {ip.IPAddress}"
                    : $"💻 {ip.WorkstationName} ({ip.IPAddress})";
                locationNode.Nodes.Add(new TreeNode(display)
                {
                    Tag = new NodeData { Type = "IP", ID = ip.ID, Data = ip }
                });
            }
            parentNode.Nodes.Add(locationNode);
        }

        private void OnTreeNodeSelected(TreeNode node)
        {
            if (node?.Tag == null) return;
            var nodeData = node.Tag as NodeData;
            if (nodeData == null) return;

            var nameBox = FindControl<TextBox>("nameTextBox");
            var addressBox = FindControl<TextBox>("addressTextBox");
            var ipGrid = FindControl<DataGridView>("ipDataGridView");
            var detailsLabel = FindControl<Label>("detailsLabel");
            if (nameBox == null || addressBox == null || ipGrid == null || detailsLabel == null) return;

            ipGrid.Rows.Clear();
            HideDeviceDetails();

            if (nodeData.Type == "Customer")
            {
                // Data aus DB neu laden statt aus dem Tag-Cache lesen
                var customers = dbManager.GetCustomers();
                var c = customers.FirstOrDefault(x => x.ID == nodeData.ID);
                if (c == null) return;
                detailsLabel.Text = "Kunde bearbeiten";
                nameBox.Text = c.Name;
                addressBox.Text = c.Address ?? "";
                nameBox.Enabled = addressBox.Enabled = true;
            }
            else if (nodeData.Type == "Location")
            {
                // Data aus DB neu laden — verhindert veraltete Daten nach Tree-Rebuild
                var loc = dbManager.GetLocationByID(nodeData.ID);
                if (loc == null) return;
                detailsLabel.Text = $"{(loc.Level == 0 ? "Standort" : "Abteilung")} bearbeiten (Level {loc.Level})";
                nameBox.Text = loc.Name;
                addressBox.Text = loc.Address ?? "";
                nameBox.Enabled = addressBox.Enabled = true;
                foreach (var ip in dbManager.GetIPsWithWorkstationByLocation(loc.ID))
                    ipGrid.Rows.Add(ip.WorkstationName ?? "", ip.IPAddress);
            }
            else if (nodeData.Type == "IP")
            {
                if (nodeData.Data is LocationIP ip)
                {
                    detailsLabel.Text = "IP-Adresse";
                    nameBox.Text = ip.WorkstationName ?? "";
                    addressBox.Text = ip.IPAddress ?? "";
                    nameBox.Enabled = addressBox.Enabled = false;

                    // Geräte-Panel befüllen und einblenden
                    ShowDeviceDetails(ip.IPAddress);
                }
            }
        }

        /// <summary>
        /// Lädt Hardware- und Software-Infos für eine IP aus der DB
        /// und zeigt sie im deviceDetailPanel an.
        /// </summary>
        private void ShowDeviceDetails(string ipAddress)
        {
            currentDisplayedIP = ipAddress;

            var devicePanel = FindControl<Panel>("deviceDetailPanel");
            var titleLabel = FindControl<Label>("deviceTitleLabel");
            var infoLabel = FindControl<Label>("deviceInfoLabel");
            var hwBox = FindControl<TextBox>("deviceHardwareBox");
            var swGrid = FindControl<DataGridView>("deviceSoftwareGrid");
            if (devicePanel == null || hwBox == null || swGrid == null) return;

            var device = dbManager.LoadDevices("Alle").FirstOrDefault(d => d.IP == ipAddress);
            if (device == null) { devicePanel.Visible = false; return; }

            // Info-Zeile
            infoLabel.Text = $"IP: {device.IP}   |   MAC: {device.MacAddress ?? "unbekannt"}   |   " +
                             $"Hostname: {device.Hostname ?? "-"}   |   Zuletzt gesehen: {device.Zeitstempel}";

            // ── Hardware / Nmap ──────────────────────────────────
            var hw = new System.Text.StringBuilder();
            hw.AppendLine("=== Geräteinformationen ===");
            hw.AppendLine($"IP-Adresse  : {device.IP}");
            hw.AppendLine($"MAC-Adresse : {device.MacAddress ?? "nicht ermittelt"}");
            hw.AppendLine($"Hostname    : {device.Hostname ?? "-"}");
            hw.AppendLine($"Status      : {device.Status ?? "-"}");
            hw.AppendLine($"Letzte Scan : {device.Zeitstempel}");

            // OS aus letztem Nmap-Detail
            var nmapDetail = dbManager.GetLatestNmapDetail(ipAddress);
            if (nmapDetail != null)
            {
                if (!string.IsNullOrEmpty(nmapDetail.OS))
                {
                    hw.AppendLine();
                    hw.AppendLine("=== Betriebssystem ===");
                    hw.AppendLine($"OS          : {nmapDetail.OS}");
                    if (!string.IsNullOrEmpty(nmapDetail.OSDetails))
                        hw.AppendLine($"Details     : {nmapDetail.OSDetails}");
                }
                if (!string.IsNullOrEmpty(nmapDetail.Vendor))
                    hw.AppendLine($"Hersteller  : {nmapDetail.Vendor}");
                hw.AppendLine($"Letzter Scan: {nmapDetail.ScanTime:dd.MM.yyyy HH:mm}");
            }

            // Ports aus DB
            var ports = dbManager.GetPortsByDevice(ipAddress);
            if (ports.Count > 0)
            {
                hw.AppendLine();
                hw.AppendLine("=== Offene Ports ===");
                hw.AppendLine($"{"Port",-10} {"Protokoll",-10} {"Status",-10} {"Service",-20} {"Version"}");
                hw.AppendLine(new string('-', 75));
                foreach (var p in ports)
                {
                    hw.AppendLine($"{p.Port}/{p.Protocol,-7} {"open",-10} {p.Service,-20} {p.Version}");
                    if (!string.IsNullOrEmpty(p.Banner))
                        hw.AppendLine($"  └ Banner: {p.Banner}");
                }
            }
            else if (!string.IsNullOrEmpty(device.Ports) && device.Ports != "-")
            {
                hw.AppendLine();
                hw.AppendLine("=== Ports (kompakt) ===");
                foreach (var p in device.Ports.Split(','))
                    hw.AppendLine($"  {p.Trim()}");
            }

            // MAC-History
            var macHistory = dbManager.GetDeviceMacHistory(device.ID);
            if (macHistory.Count > 0)
            {
                hw.AppendLine();
                hw.AppendLine("=== MAC-History ===");
                foreach (var entry in macHistory)
                    hw.AppendLine($"  {entry.Item3:dd.MM.yyyy HH:mm}   {entry.Item1,-20}  {entry.Item2}");
            }

            hwBox.Text = hw.ToString();

            // ── Software ─────────────────────────────────────────
            swGrid.Rows.Clear();
            var software = dbManager.LoadSoftware("Alle")
                .Where(s => s.PCName == device.Hostname || s.PCName == device.IP)
                .OrderBy(s => s.Name)
                .ToList();

            if (software.Count > 0)
            {
                foreach (var sw in software)
                    swGrid.Rows.Add(sw.Name, sw.Version ?? "", sw.Publisher ?? "", sw.InstallDate ?? "");
                titleLabel.Text = $"💻 Geräteinformationen  ({ports.Count} Ports, {software.Count} Software-Einträge)";
            }
            else
            {
                swGrid.Rows.Add("Keine Software-Einträge vorhanden", "", "", "");
                titleLabel.Text = $"💻 Geräteinformationen  ({ports.Count} Ports)";
            }

            devicePanel.Visible = true;
        }

        private void HideDeviceDetails()
        {
            var devicePanel = FindControl<Panel>("deviceDetailPanel");
            if (devicePanel != null) devicePanel.Visible = false;
        }

        // =========================================================
        // === TREE AKTIONEN ===
        // =========================================================

        private void AddCustomer()
        {
            using (var form = new InputDialog("Neuer Kunde", "Kundenname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                { dbManager.AddCustomer(form.Value1, form.Value2); LoadCustomerTree(); statusLabel.Text = "Kunde hinzugefügt"; }
        }

        private void AddLocation()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData?.Type != "Customer") { MessageBox.Show("Bitte wähle einen Kunden aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new InputDialog("Neuer Standort", "Standortname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                { dbManager.AddLocation(nodeData.ID, form.Value1, form.Value2); LoadCustomerTree(); statusLabel.Text = "Standort hinzugefügt"; }
        }

        private void AddDepartment()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData?.Type != "Location") { MessageBox.Show("Bitte wähle einen Standort oder eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using (var form = new InputDialog("Neue Abteilung", "Abteilungsname:", "Adresse:"))
                if (form.ShowDialog(this) == DialogResult.OK)
                { dbManager.AddChildLocation(nodeData.ID, form.Value1, form.Value2); LoadCustomerTree(); statusLabel.Text = "Abteilung hinzugefügt"; }
        }

        private void DeleteNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData == null) { MessageBox.Show("Bitte wähle einen Knoten aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            string confirmMsg = "";
            if (nodeData.Type == "Customer") confirmMsg = "Kunden und alle zugehörigen Standorte/Geräte wirklich löschen?";
            else if (nodeData.Type == "Location") confirmMsg = "Diesen Standort/diese Abteilung und alle Untereinträge wirklich löschen?";
            else if (nodeData.Type == "IP") confirmMsg = "Dieses Gerät aus der Abteilung entfernen?";

            if (MessageBox.Show(confirmMsg, "Löschen bestätigen",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            if (nodeData.Type == "Customer")
            {
                dbManager.DeleteCustomer(nodeData.ID);
                statusLabel.Text = "Kunde gelöscht";
            }
            else if (nodeData.Type == "Location")
            {
                dbManager.DeleteLocation(nodeData.ID);
                statusLabel.Text = "Standort/Abteilung gelöscht";
            }
            else if (nodeData.Type == "IP")
            {
                // IP aus der übergeordneten Location entfernen
                var parentNodeData = treeView.SelectedNode?.Parent?.Tag as NodeData;
                if (parentNodeData?.Type == "Location" && nodeData.Data is LocationIP ip)
                {
                    dbManager.RemoveIPFromLocation(parentNodeData.ID, ip.IPAddress);
                    statusLabel.Text = $"Gerät {ip.IPAddress} entfernt";
                }
                else
                {
                    MessageBox.Show("Fehler: Übergeordneter Knoten ist keine Location!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            HideDeviceDetails();
            currentDisplayedIP = "";
            LoadCustomerTree();
        }

        private void SaveNodeDetails()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            var nameBox = FindControl<TextBox>("nameTextBox");
            var addressBox = FindControl<TextBox>("addressTextBox");
            if (nodeData == null || nameBox == null || addressBox == null) return;

            if (nodeData.Type == "Customer")
            {
                dbManager.UpdateCustomer(nodeData.ID, nameBox.Text, addressBox.Text);
                statusLabel.Text = "Kunde aktualisiert";
            }
            else if (nodeData.Type == "Location")
            {
                dbManager.UpdateLocation(nodeData.ID, nameBox.Text, addressBox.Text);
                statusLabel.Text = "Standort aktualisiert";
            }
            LoadCustomerTree();
            SelectTreeNodeById(treeView, nodeData.ID);
        }

        // =========================================================
        // === IP VERWALTUNG ===
        // =========================================================

        private void AddIPToNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            if (nodeData?.Type != "Location") { MessageBox.Show("Bitte wähle einen Standort oder eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            string ip = FindControl<TextBox>("ipInputTextBox")?.Text.Trim() ?? "";
            string ws = FindControl<TextBox>("workstationInputTextBox")?.Text.Trim() ?? "";
            if (string.IsNullOrEmpty(ip)) { MessageBox.Show("Bitte gib eine IP-Adresse ein!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            int locationId = nodeData.ID;
            dbManager.AddIPToLocation(locationId, ip, ws);
            FindControl<TextBox>("ipInputTextBox").Clear();
            FindControl<TextBox>("workstationInputTextBox").Clear();
            LoadCustomerTree();
            SelectTreeNodeById(treeView, locationId);
            statusLabel.Text = "IP manuell hinzugefügt";
        }

        private void RemoveIPFromNode()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            var ipGrid = FindControl<DataGridView>("ipDataGridView");
            if (nodeData?.Type != "Location" || ipGrid?.SelectedRows.Count == 0) return;

            int locationId = nodeData.ID;
            int removedCount = 0;

            foreach (DataGridViewRow row in ipGrid.SelectedRows)
            {
                string ip = row.Cells[1].Value?.ToString();
                if (!string.IsNullOrEmpty(ip))
                {
                    dbManager.RemoveIPFromLocation(locationId, ip);
                    removedCount++;
                }
            }

            LoadCustomerTree();
            SelectTreeNodeById(treeView, locationId);
            statusLabel.Text = $"{removedCount} IP(s) entfernt";
        }

        private void MoveIPToAnotherLocation()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            var ipGrid = FindControl<DataGridView>("ipDataGridView");

            if (nodeData == null)
            {
                MessageBox.Show("Bitte wähle einen Standort oder eine IP im Baum aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string ipToMove = null;
            string workstation = null;
            int sourceLocationID = -1;

            if (nodeData.Type == "IP")
            {
                // IP-Knoten direkt im Baum angeklickt:
                // Elternknoten liefert die LocationID, IP steht im Tag.ID (LocationIP.ID)
                // Wir lesen die echten Daten frisch aus der DB
                var parentNodeData = treeView.SelectedNode?.Parent?.Tag as NodeData;
                if (parentNodeData?.Type != "Location")
                {
                    MessageBox.Show("Fehler: Übergeordneter Knoten ist keine Location!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                sourceLocationID = parentNodeData.ID;

                // IP frisch aus DB laden anhand der ID im Tag
                var ipsInParent = dbManager.GetIPsWithWorkstationByLocation(sourceLocationID);
                var ipEntry = ipsInParent.FirstOrDefault(x => x.ID == nodeData.ID);

                if (ipEntry == null)
                {
                    MessageBox.Show("IP-Eintrag konnte nicht gefunden werden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ipToMove = ipEntry.IPAddress;
                workstation = ipEntry.WorkstationName ?? "";
            }
            else if (nodeData.Type == "Location")
            {
                // Location-Knoten gewählt — IP muss aus der Tabelle rechts ausgewählt sein
                if (ipGrid == null || ipGrid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Bitte wähle eine IP-Adresse aus der Tabelle aus\noder klicke direkt auf eine IP im Baum!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ipToMove = ipGrid.SelectedRows[0].Cells[1].Value?.ToString();
                workstation = ipGrid.SelectedRows[0].Cells[0].Value?.ToString() ?? "";
                sourceLocationID = nodeData.ID;

                if (string.IsNullOrEmpty(ipToMove))
                {
                    MessageBox.Show("Fehler beim Lesen der IP-Adresse!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Bitte wähle einen Standort oder eine IP im Baum aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Ziel-Location auswählen
            using (var form = new LocationSelectionDialog(dbManager, sourceLocationID))
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        dbManager.RemoveIPFromLocation(sourceLocationID, ipToMove);
                        dbManager.AddIPToLocation(form.SelectedLocationID, ipToMove, workstation);

                        var target = dbManager.GetLocationByID(form.SelectedLocationID);
                        LoadCustomerTree();
                        SelectTreeNodeById(treeView, form.SelectedLocationID);

                        statusLabel.Text = $"IP {ipToMove} nach '{target?.Name}' verschoben";
                        MessageBox.Show($"IP {ipToMove} erfolgreich nach '{target?.Name}' verschoben!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;
            var ipGrid = FindControl<DataGridView>("ipDataGridView");
            if (nodeData?.Type != "Location" || ipGrid == null) return;
            string currentWS = ipGrid.Rows[rowIndex].Cells[0].Value?.ToString();
            string currentIP = ipGrid.Rows[rowIndex].Cells[1].Value?.ToString();
            using (var form = new InputDialog("IP-Eintrag bearbeiten", "Arbeitsplatz:", "IP-Adresse:"))
            {
                form.Value1 = currentWS; form.Value2 = currentIP;
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    dbManager.RemoveIPFromLocation(nodeData.ID, currentIP);
                    dbManager.AddIPToLocation(nodeData.ID, form.Value2, form.Value1);
                    LoadCustomerTree();
                    OnTreeNodeSelected(treeView.SelectedNode);
                    statusLabel.Text = "IP-Eintrag aktualisiert";
                }
            }
        }

        // =========================================================
        // === IP IMPORT AUS DATENBANK ===
        // =========================================================

        /// <summary>
        /// Öffnet den IP-Import-Dialog, der alle gescannten Geräte aus der DB anzeigt.
        /// Der Benutzer kann IPs auswählen und direkt der gewählten Abteilung zuweisen.
        /// </summary>
        private void ImportIPsFromDatabase()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;

            if (nodeData?.Type != "Location")
            {
                MessageBox.Show("Bitte wähle zuerst einen Standort oder eine Abteilung im Baum aus!\n\nDie importierten IPs werden dieser Abteilung zugewiesen.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Frisch aus DB laden — nicht dem möglicherweise veralteten Tag.Data vertrauen
            var location = dbManager.GetLocationByID(nodeData.ID);
            if (location == null)
            {
                MessageBox.Show("Standort nicht gefunden!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var dialog = new IPImportDialog(dbManager, location.ID, location.Name))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK || dialog.SelectedIPs.Count == 0)
                    return;

                int imported = 0;
                var alreadyAssigned = new HashSet<string>(
                    dbManager.GetIPsWithWorkstationByLocation(location.ID).Select(ip => ip.IPAddress));

                foreach (var entry in dialog.SelectedIPs)
                {
                    // entry ist Tuple<string,string>: Item1 = IP, Item2 = Hostname
                    if (!alreadyAssigned.Contains(entry.Item1))
                    {
                        dbManager.AddIPToLocation(location.ID, entry.Item1, entry.Item2);
                        imported++;
                    }
                }

                int skipped = dialog.SelectedIPs.Count - imported;
                LoadCustomerTree();

                // Nach Tree-Rebuild den Node neu suchen und selektieren
                SelectTreeNodeById(treeView, nodeData.ID);

                statusLabel.Text = $"✅ {imported} IP(s) nach '{location.Name}' importiert"
                                 + (skipped > 0 ? $" ({skipped} bereits vorhanden)" : "");

                MessageBox.Show($"{imported} IP-Adresse(n) erfolgreich nach '{location.Name}' importiert!",
                    "Import erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Sucht nach dem Tree-Rebuild einen Location-Knoten anhand der ID und selektiert ihn.
        /// </summary>
        private void SelectTreeNodeById(TreeView treeView, int locationId)
        {
            foreach (TreeNode root in treeView.Nodes)
            {
                var found = FindNodeById(root, locationId);
                if (found != null)
                {
                    treeView.SelectedNode = found;
                    found.EnsureVisible();
                    OnTreeNodeSelected(found);
                    return;
                }
            }
        }

        private TreeNode FindNodeById(TreeNode parent, int locationId)
        {
            var nodeData = parent.Tag as NodeData;
            if (nodeData?.Type == "Location" && nodeData.ID == locationId)
                return parent;

            foreach (TreeNode child in parent.Nodes)
            {
                var found = FindNodeById(child, locationId);
                if (found != null) return found;
            }
            return null;
        }

        // =========================================================
        // === LOCATION COMBO ===
        // =========================================================

        private void LoadLocationCombo()
        {
            locationComboBox.Items.Clear();
            locationComboBox.Items.Add(new ComboItem { ID = -1, Display = "-- Alle Standorte --" });
            foreach (var customer in dbManager.GetCustomers())
                foreach (var location in dbManager.GetLocationsByCustomer(customer.ID))
                    AddLocationComboRecursive(customer.Name, location);
            if (locationComboBox.Items.Count > 0) locationComboBox.SelectedIndex = 0;
        }

        private void AddLocationComboRecursive(string customerName, Location location, string prefix = "")
        {
            locationComboBox.Items.Add(new ComboItem { ID = location.ID, Display = $"{customerName} > {prefix}{location.Name}" });
            foreach (var child in dbManager.GetChildLocations(location.ID))
                AddLocationComboRecursive(customerName, child, prefix + location.Name + " > ");
        }

        private void OnLocationSelected()
        {
            if (locationComboBox.SelectedItem is ComboItem item)
                selectedLocationID = item.ID;
        }

        private void CreateNewLocation()
        {
            var customers = dbManager.GetCustomers();
            if (customers.Count == 0) { MessageBox.Show("Bitte erstelle zuerst einen Kunden!", "Info"); return; }
            using (var form = new CustomerSelectionForm(customers))
                if (form.ShowDialog(this) == DialogResult.OK)
                    using (var input = new InputDialog("Neuer Standort", "Standortname:", "Adresse:"))
                        if (input.ShowDialog(this) == DialogResult.OK)
                        { dbManager.AddLocation(form.SelectedCustomerID, input.Value1, input.Value2); LoadCustomerTree(); statusLabel.Text = "Standort erstellt"; }
        }

        // =========================================================
        // === HILFSMETHODEN ===
        // =========================================================

        private void UpdateHardwareLabels(string pcName, bool isRemote)
        {
            var tab = tabControl.TabPages["Hardware"];
            var infoPanel = (tab?.Controls[0] as Panel)?.Controls.OfType<Panel>().FirstOrDefault();
            if (infoPanel == null) return;
            (infoPanel.Controls["hardwarePCLabel"] as Label).Text = isRemote ? $"Remote-PC: {pcName}" : $"Lokaler PC: {pcName}";
            (infoPanel.Controls["hardwareUpdateLabel"] as Label).Text = $"Letzte Aktualisierung: {DateTime.Now:dd.MM.yyyy HH:mm:ss}";
        }

        private void OnSoftwareUpdateClick(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == softwareGridView.Columns.Count - 1)
            {
                string name = softwareGridView.Rows[e.RowIndex].Cells[0].Value?.ToString();
                if (!string.IsNullOrEmpty(name)) softwareManager.UpdateSoftware(name, currentRemotePC, statusLabel);
            }
        }

        private List<DatabaseDevice> GetDevicesFromGrid()
        {
            var list = new List<DatabaseDevice>();
            foreach (DataGridViewRow row in dbDeviceTable.Rows)
                if (row.Cells[0].Value != null)
                    list.Add(new DatabaseDevice { ID = Convert.ToInt32(row.Cells[0].Value), Zeitstempel = row.Cells[1].Value?.ToString(), IP = row.Cells[2].Value?.ToString(), Hostname = row.Cells[3].Value?.ToString(), MacAddress = row.Cells[4].Value?.ToString(), Status = row.Cells[5].Value?.ToString(), Ports = row.Cells[6].Value?.ToString() });
            return list;
        }

        private List<DatabaseSoftware> GetSoftwareFromGrid()
        {
            var list = new List<DatabaseSoftware>();
            foreach (DataGridViewRow row in dbSoftwareTable.Rows)
                if (row.Cells[0].Value != null)
                    list.Add(new DatabaseSoftware { ID = Convert.ToInt32(row.Cells[0].Value), Zeitstempel = row.Cells[1].Value?.ToString(), PCName = row.Cells[2].Value?.ToString(), Name = row.Cells[3].Value?.ToString(), Version = row.Cells[4].Value?.ToString(), Publisher = row.Cells[5].Value?.ToString(), InstallDate = row.Cells[6].Value?.ToString(), LastUpdate = row.Cells[7].Value?.ToString() });
            return list;
        }

        private T FindControl<T>(string name) where T : Control
            => Controls.Find(name, true).FirstOrDefault() as T;
    }

    // =========================================================
    // === HILFSKLASSE: ComboBox-Item mit ID ===
    // =========================================================
    public class ComboItem
    {
        public int ID { get; set; }
        public string Display { get; set; }
        public override string ToString() => Display;
    }

    // =========================================================
    // === LOCATION SELECTION DIALOG ===
    // =========================================================
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
            Width = 400; Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Segoe UI", 10);

            var label = new Label { Text = "Ziel-Abteilung auswählen:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold) };
            locationTree = new TreeView { Location = new Point(10, 40), Width = 360, Height = 380, Font = new Font("Segoe UI", 10) };
            PopulateLocationTree();

            var okBtn = new Button { Text = "Verschieben", Location = new Point(140, 430), Width = 110, DialogResult = DialogResult.OK, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 430), Width = 110, DialogResult = DialogResult.Cancel, Font = new Font("Segoe UI", 10) };
            okBtn.Click += (s, e) =>
            {
                var nodeData = locationTree.SelectedNode?.Tag as NodeData;
                if (nodeData?.Type == "Location") SelectedLocationID = nodeData.ID;
                else { MessageBox.Show("Bitte wähle eine Abteilung aus!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); DialogResult = DialogResult.None; }
            };
            Controls.AddRange(new Control[] { label, locationTree, okBtn, cancelBtn });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }

        private void PopulateLocationTree()
        {
            locationTree.Nodes.Clear();
            foreach (var customer in dbManager.GetCustomers())
            {
                var customerNode = new TreeNode($"👤 {customer.Name}") { Tag = new NodeData { Type = "Customer", ID = customer.ID } };
                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                for (int i = 0; i < locations.Count; i++)
                    AddLocationNode(customerNode, locations[i], $"S{i + 1}");
                locationTree.Nodes.Add(customerNode);
            }
            locationTree.ExpandAll();
        }

        private void AddLocationNode(TreeNode parent, Location location, string shortID)
        {
            var children = dbManager.GetChildLocations(location.ID);

            // Quell-Knoten selbst ausblenden — Kinder aber trotzdem an parent hängen
            if (location.ID == sourceLocationID)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    string cid = location.Level == 0 ? $"{shortID}.A{i + 1}" : $"{shortID}.{i + 1}";
                    AddLocationNode(parent, children[i], cid);
                }
                return;
            }

            string icon = location.Level == 0 ? "🏢" : "📂";
            var node = new TreeNode($"[{shortID}] {icon} {location.Name}")
            {
                Tag = new NodeData { Type = "Location", ID = location.ID }
            };

            for (int i = 0; i < children.Count; i++)
            {
                // location.Level kommt direkt aus GetChildLocations — ist korrekt befüllt
                string cid = location.Level == 0 ? $"{shortID}.A{i + 1}" : $"{shortID}.{i + 1}";
                AddLocationNode(node, children[i], cid);
            }

            parent.Nodes.Add(node);
        }
    }

    // =========================================================
    // === MAIN ENTRY ===
    // =========================================================
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