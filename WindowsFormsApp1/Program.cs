using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using ClosedXML.Excel;

namespace NmapInventory
{
    public class MainForm : Form
    {
        // === UI COMPONENTS ===
        private TabControl tabControl;
        private TextBox networkTextBox;
        private Button scanButton, remoteHardwareButton, exportButton, snmpScanButton, snmpSettingsButton, refreshButton;
        private Label statusLabel;
        private ProgressBar scanProgressBar;
        private System.Windows.Forms.Timer scanAnimTimer;
        private DataGridView deviceTable, softwareGridView, dbDeviceTable, dbSoftwareTable;
        private TextBox rawOutputTextBox, hardwareInfoTextBox;
        private ComboBox locationComboBox;
        private Button newLocationButton;

        // === MANAGERS ===
        private NmapScanner nmapScanner;
        private DatabaseManager dbManager;
        private HardwareManager hardwareManager;
        private SoftwareManager softwareManager;
        private LinuxManager linuxManager;
        private AdbManager adbManager;
        private IosManager iosManager;
        private SnmpManager snmpManager;
        private SnmpSettings snmpSettings = new SnmpSettings();

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
            if (!nmapScanner.IsNmapAvailable)
                MessageBox.Show(
                    "Nmap wurde nicht gefunden!\n\n" +
                    "Bitte eine der folgenden Optionen wählen:\n" +
                    "  1. nmap.exe in denselben Ordner wie NmapInventory.exe kopieren\n" +
                    "  2. Nmap installieren: https://nmap.org/download.html\n\n" +
                    "Der Scan-Button wird nicht funktionieren bis Nmap verfügbar ist.",
                    "Nmap nicht gefunden",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            dbManager = new DatabaseManager();
            hardwareManager = new HardwareManager();
            softwareManager = new SoftwareManager();
            snmpManager     = new SnmpManager();
            linuxManager    = new LinuxManager();
            adbManager      = new AdbManager();
            iosManager      = new IosManager();

            InitializeComponent();
            dbManager.InitializeDatabase();
            // Inventar_* sind jetzt VIEWs — kein Refresh nötig
            LoadCustomerTree();
            ApplyAutoSizing();

            this.Load += (s, e) => RestoreSplitterPosition();
        }

        private void RestoreSplitterPosition()
        {
            var split = Controls.Find("customerSplitContainer", true)
                                .FirstOrDefault() as SplitContainer;
            if (split == null) return;

            // Mittig — halbe Breite des SplitContainers
            split.SplitterDistance = split.Width / 2;
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

            // Ladebalken — erscheint nur während Scan, sonst versteckt
            scanProgressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 6,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 25,
                Visible = false
            };

            // Sanduhr-Animation im Tab-Icon während Scan
            scanAnimTimer = new System.Windows.Forms.Timer { Interval = 600 };
            int sandFrame = 0;
            string[] sandFrames = { "⏳", "⌛" };
            scanAnimTimer.Tick += (s, e) =>
            {
                sandFrame = (sandFrame + 1) % sandFrames.Length;
                statusLabel.Text = statusLabel.Text.TrimStart('⏳', '⌛', ' ');
                statusLabel.Text = sandFrames[sandFrame] + "  " + statusLabel.Text;
            };

            Controls.AddRange(new Control[] { tabControl, topPanel, scanProgressBar, statusLabel });
        }

        // =========================================================
        // === TOP PANEL ===
        // =========================================================

        private Panel CreateTopPanel()
        {
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = SystemColors.Control };

            var networkLabel = new Label { Text = "Netzwerk:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            networkTextBox = new TextBox { Text = "192.168.2.0/24", Location = new Point(100, 8), Width = 200, Font = new Font("Segoe UI", 10) };

            var locationLabel = new Label { Text = "Standort:", Location = new Point(10, 45), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            locationComboBox = new ComboBox { Location = new Point(100, 43), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            locationComboBox.SelectedIndexChanged += (s, e) => OnLocationSelected();

            newLocationButton = new Button { Location = new Point(310, 43), Width = 140, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            SetButtonIcon(newLocationButton, "+ Neuer Standort", 3);    // imageres: Ordner
            newLocationButton.Click += (s, e) => CreateNewLocation();

            scanButton = new Button { Location = new Point(460, 8), Width = 165, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            SetButtonIcon(scanButton, "Netzwerk scannen", 168);  // imageres: Lupe/Suche
            scanButton.Click += (s, e) => StartScan();

            remoteHardwareButton = new Button { Location = new Point(634, 8), Width = 150, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            SetButtonIcon(remoteHardwareButton, "Remote Scan", 109);  // imageres: Computer/Monitor
            remoteHardwareButton.Click += (s, e) => StartRemoteHardwareQuery();

            exportButton = new Button { Location = new Point(793, 8), Width = 130, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            SetButtonIcon(exportButton, "Exportieren", 162);  // imageres: Speichern/Export
            exportButton.Click += (s, e) => ExportData();

            snmpScanButton = new Button { Location = new Point(460, 42), Width = 165, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            snmpScanButton.Text = "📡 SNMP-Scan";
            snmpScanButton.Click += (s, e) => StartSnmpScan();

            snmpSettingsButton = new Button { Location = new Point(634, 42), Width = 150, Height = 28, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            snmpSettingsButton.Text = "⚙️ SNMP-Einstellungen";
            snmpSettingsButton.Click += (s, e) => OpenSnmpSettings();

            refreshButton = new Button { Location = new Point(793, 42), Width = 130, Height = 28, BackColor = Color.FromArgb(220, 240, 220), Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            refreshButton.Text = "🔄 Aktualisieren";
            refreshButton.Click += (s, e) => RefreshAllViews();

            topPanel.Controls.AddRange(new Control[] { networkLabel, networkTextBox, locationLabel, locationComboBox, newLocationButton, scanButton, remoteHardwareButton, exportButton, snmpScanButton, snmpSettingsButton, refreshButton });
            return topPanel;
        }

        // =========================================================
        // === TAB CONTROL ===
        // =========================================================

        private TabControl CreateTabControl()
        {
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10),
                DrawMode = TabDrawMode.OwnerDrawFixed,
                ItemSize = new Size(48, 40),
                SizeMode = TabSizeMode.Fixed
            };

            // ImageList mit Icons aus imageres.dll (Windows 7+, zuverlässige Indizes)
            var tabImages = new ImageList { ImageSize = new Size(24, 24), ColorDepth = ColorDepth.Depth32Bit };
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 109)); // 0: Geräte        → Computer/Monitor
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 168)); // 1: Nmap Ausgabe  → Lupe/Suche
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 25)); // 2: Hardware      → Chip/CPU
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 2)); // 3: Software      → Anwendungsfenster
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 27)); // 4: DB Geräte     → Datenbank/Server
            tabImages.Images.Add(ExtractDllIcon("shell32.dll", 20)); // 5: DB Software   → Diskette
            tabImages.Images.Add(GetStandardUserIcon(24));              // 6: Kunden        → Standard-Benutzer-Icon
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 8)); // 7: Auswertung    → Diagramm/Tabelle
            tabControl.ImageList = tabImages;

            // Tooltip für Mouse-Over
            var tabToolTip = new ToolTip { ShowAlways = true, InitialDelay = 300, AutoPopDelay = 5000 };
            string[] tabNames = { "Geräte", "Nmap Ausgabe", "Hardware", "Software", "DB - Geräte", "DB - Software", "Kunden / Standorte", "Auswertung" };

            tabControl.DrawItem += (s, e) =>
            {
                var tab = tabControl.TabPages[e.Index];
                var rect = e.Bounds;
                bool sel = (e.Index == tabControl.SelectedIndex);

                // Hintergrund
                e.Graphics.FillRectangle(sel
                    ? SystemBrushes.Window
                    : SystemBrushes.Control, rect);

                // Icon zentriert
                if (tabControl.ImageList != null && e.Index < tabNames.Length)
                {
                    var img = tabControl.ImageList.Images[e.Index];
                    if (img != null)
                    {
                        int ix = rect.Left + (rect.Width - img.Width) / 2;
                        int iy = rect.Top + (rect.Height - img.Height) / 2;
                        e.Graphics.DrawImage(img, ix, iy);
                    }
                }
            };

            tabControl.MouseMove += (s, e) =>
            {
                for (int i = 0; i < tabControl.TabPages.Count; i++)
                {
                    if (tabControl.GetTabRect(i).Contains(e.Location))
                    {
                        string tip = i < tabNames.Length ? tabNames[i] : tabControl.TabPages[i].Text;
                        if (tabToolTip.GetToolTip(tabControl) != tip)
                            tabToolTip.SetToolTip(tabControl, tip);
                        return;
                    }
                }
                tabToolTip.SetToolTip(tabControl, "");
            };

            tabControl.TabPages.Add(CreateDeviceTab());
            tabControl.TabPages.Add(CreateNmapTab());
            tabControl.TabPages.Add(CreateHardwareTab());
            tabControl.TabPages.Add(CreateSoftwareTab());
            tabControl.TabPages.Add(CreateDBDeviceTab());
            tabControl.TabPages.Add(CreateDBSoftwareTab());
            tabControl.TabPages.Add(CreateCustomerLocationTab());
            tabControl.TabPages.Add(CreateAuswertungTab());

            // ImageIndex pro Tab setzen
            for (int i = 0; i < tabControl.TabPages.Count; i++)
                tabControl.TabPages[i].ImageIndex = i;

            return tabControl;
        }

        private TabPage CreateDeviceTab()
        {
            Panel devicePanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = SystemColors.Control, Padding = new Padding(10) };
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
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status",        Width = 50  });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP",            Width = 120 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname",      Width = 150 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Gerätetyp",    Width = 130 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller",   Width = 130 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC-Adresse",  Width = 140 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Kommentar",    Width = 200 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Online Status", Width = 80  });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 400 });
            // Nur Kommentar-Spalte (Index 6) editierbar — alle anderen per CellBeginEdit sperren
            deviceTable.CellBeginEdit += (s, e) =>
            {
                if (e.ColumnIndex != 6) e.Cancel = true;
            };
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

            // Kommentar-Spalte: Änderung direkt in DB speichern (Spalte 6)
            deviceTable.CellEndEdit += (s, e) =>
            {
                if (e.ColumnIndex != 6 || e.RowIndex < 0) return;
                string ip      = deviceTable.Rows[e.RowIndex].Cells[1].Value?.ToString();
                string comment = deviceTable.Rows[e.RowIndex].Cells[6].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(ip))
                    dbManager.SaveComment(ip, comment);
            };

            // Rechtsklick auf Zeile → Gerätetyp manuell setzen
            var ctxDeviceTable = new ContextMenuStrip();
            deviceTable.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Right) return;
                var hit = deviceTable.HitTest(e.X, e.Y);
                if (hit.RowIndex < 0) return;
                deviceTable.Rows[hit.RowIndex].Selected = true;

                ctxDeviceTable.Items.Clear();
                string rowIp = deviceTable.Rows[hit.RowIndex].Cells[1].Value?.ToString();
                if (string.IsNullOrEmpty(rowIp)) return;

                ctxDeviceTable.Items.Add(new ToolStripLabel($"Gerätetyp setzen für {rowIp}") { Font = new Font("Segoe UI", 9, FontStyle.Bold), Enabled = false });
                ctxDeviceTable.Items.Add(new ToolStripSeparator());

                void AddType(string label, DeviceType dt)
                {
                    ctxDeviceTable.Items.Add(label, null, (_, __) =>
                    {
                        dbManager.SetDeviceType(rowIp, dt);
                        LoadCustomerTree();
                        // Zeile direkt aktualisieren
                        string icon = $"{DeviceTypeHelper.GetIcon(dt)} {DeviceTypeHelper.GetLabel(dt)}";
                        deviceTable.Rows[hit.RowIndex].Cells[3].Value = icon;
                    });
                }

                AddType("🖥  Windows PC",      DeviceType.WindowsPC);
                AddType("🗄  Windows Server",   DeviceType.WindowsServer);
                AddType("💻 Laptop",            DeviceType.Laptop);
                AddType("🐧 Linux",             DeviceType.Linux);
                AddType("🍎 macOS / iOS",       DeviceType.MacOS);
                AddType("📱 Smartphone",        DeviceType.Smartphone);
                AddType("📟 Tablet",            DeviceType.Tablet);
                AddType("🔀 Netzwerkgerät",     DeviceType.NetzwerkGeraet);
                AddType("💾 NAS",               DeviceType.NAS);
                AddType("📺 Smart TV",          DeviceType.SmartTV);
                AddType("💡 IoT-Gerät",         DeviceType.IoT);
                AddType("🖨  Drucker",           DeviceType.Drucker);
                AddType("❓ Unbekannt",          DeviceType.Unbekannt);

                ctxDeviceTable.Show(deviceTable, e.Location);
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
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = SystemColors.Control, Padding = new Padding(10) };
            panel.Controls.AddRange(new Control[] {
                new Label { Name = "hardwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue },
                new Label { Name = "hardwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = SystemColors.ControlText }
            });
            hardwareInfoTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(hardwareInfoTextBox);
            container.Controls.Add(panel);
            return new TabPage("Hardware") { Name = "Hardware", Controls = { container } };
        }

        private TabPage CreateSoftwareTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = SystemColors.Control, Padding = new Padding(10) };
            panel.Controls.AddRange(new Control[] {
                new Label { Name = "softwarePCLabel", Text = "Lokaler PC: " + Environment.MachineName, Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.DarkBlue },
                new Label { Name = "softwareUpdateLabel", Text = "Letzte Aktualisierung: -", Location = new Point(400, 8), AutoSize = true, Font = new Font("Segoe UI", 10), ForeColor = SystemColors.ControlText }
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
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 350 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Gerätetyp", Width = 140, ReadOnly = true });
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
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "GeräteID", Width = 80, ReadOnly = true });
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
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 320,
                SplitterWidth = 5,
                IsSplitterFixed = false,
                Name = "customerSplitContainer"
            };

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
            var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(0) };

            // ── Abschnitt: Kunde / Standort ───────────────────
            var secLabel = new Label
            {
                Name = "detailsLabel",
                Text = "Details",
                Location = new Point(12, 12),
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                AutoSize = true,
                ForeColor = SystemColors.ControlText
            };

            // Trennlinie
            var sep1 = new Panel
            {
                Location = new Point(12, 36),
                Size = new Size(620, 1),
                BackColor = SystemColors.ControlDark
            };

            // Name-Zeile
            var nameLabel = new Label
            {
                Text = "Name:",
                Location = new Point(12, 48),
                AutoSize = true,
                Font = new Font("Segoe UI", 10),
                ForeColor = SystemColors.ControlText
            };
            var nameTextBox = new TextBox
            {
                Name = "nameTextBox",
                Location = new Point(90, 45),
                Width = 340,
                Font = new Font("Segoe UI", 10)
            };

            // Adresse-Zeile
            var addressLabel = new Label
            {
                Text = "Adresse:",
                Location = new Point(12, 78),
                AutoSize = true,
                Font = new Font("Segoe UI", 10),
                ForeColor = SystemColors.ControlText
            };
            var addressTextBox = new TextBox
            {
                Name = "addressTextBox",
                Location = new Point(90, 75),
                Width = 340,
                Multiline = true,
                Height = 52,
                Font = new Font("Segoe UI", 10)
            };

            // Speichern-Button — Shell32 Icon Diskette
            var saveBtn = new Button
            {
                Location = new Point(90, 134),
                Width = 160,
                Height = 28,
                Font = new Font("Segoe UI", 10)
            };
            SetButtonIcon(saveBtn, "Speichern", 20, "shell32.dll");
            saveBtn.Click += (s, e) => SaveNodeDetails();

            // ── Abschnitt: IP-Adressen ────────────────────────
            var sep2 = new Panel
            {
                Location = new Point(12, 174),
                Size = new Size(620, 1),
                BackColor = SystemColors.ControlDark
            };
            var ipSecLabel = new Label
            {
                Text = "Zugeordnete IP-Adressen",
                Location = new Point(12, 180),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true,
                ForeColor = SystemColors.ControlText
            };

            var ipGrid = new DataGridView
            {
                Name = "ipDataGridView",
                Location = new Point(12, 204),
                Width = 620,
                Height = 190,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                Font = new Font("Segoe UI", 9),
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = SystemColors.ControlLight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                EnableHeadersVisualStyles = false
            };
            ipGrid.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            ipGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Arbeitsplatz", Width = 220 });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 140 });
            ipGrid.CellDoubleClick += (s, e) => EditIPEntry(e.RowIndex);

            // ── Eingabe-Zeile: neue IP ────────────────────────
            int iy = 404;
            var wsLabel = new Label
            {
                Text = "Arbeitsplatz:",
                Location = new Point(12, iy + 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText
            };
            var wsBox = new TextBox
            {
                Name = "workstationInputTextBox",
                Location = new Point(95, iy),
                Width = 165,
                Font = new Font("Segoe UI", 9)
            };
            var ipLbl2 = new Label
            {
                Text = "IP:",
                Location = new Point(268, iy + 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText
            };
            var ipBox = new TextBox
            {
                Name = "ipInputTextBox",
                Location = new Point(285, iy),
                Width = 130,
                Font = new Font("Segoe UI", 9)
            };
            var addIpBtn = new Button
            {
                Location = new Point(424, iy - 1),
                Width = 140,
                Height = 26,
                Font = new Font("Segoe UI", 9)
            };
            SetButtonIcon(addIpBtn, "Hinzufügen", 319, "shell32.dll");
            addIpBtn.Click += (s, e) => AddIPToNode();

            // ── Button-Leiste: Aktionen ───────────────────────
            iy += 34;
            var sep3 = new Panel
            {
                Location = new Point(12, iy - 4),
                Size = new Size(620, 1),
                BackColor = SystemColors.ControlDark
            };

            int bx = 12;
            var removeBtn = new Button
            {
                Location = new Point(bx, iy),
                Width = 130,
                Height = 28,
                Font = new Font("Segoe UI", 9)
            };
            SetButtonIcon(removeBtn, "Entfernen", 131, "shell32.dll");
            removeBtn.Click += (s, e) => RemoveIPFromNode();

            bx += 138;
            var moveBtn = new Button
            {
                Location = new Point(bx, iy),
                Width = 130,
                Height = 28,
                Font = new Font("Segoe UI", 9)
            };
            SetButtonIcon(moveBtn, "Verschieben", 46, "shell32.dll");
            moveBtn.Click += (s, e) => MoveIPToAnotherLocation();

            bx += 138;
            var importBtn = new Button
            {
                Location = new Point(bx, iy),
                Width = 155,
                Height = 28,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            SetButtonIcon(importBtn, "Aus DB importieren", 165, "imageres.dll");
            importBtn.Click += (s, e) => ImportIPsFromDatabase();

            // ── Geräte-Detail-Panel ───────────────────────────
            int dpY = iy + 42;
            var sep4 = new Panel
            {
                Location = new Point(0, dpY - 6),
                Size = new Size(660, 2),
                BackColor = SystemColors.ControlDark
            };

            var devicePanel = new Panel
            {
                Name = "deviceDetailPanel",
                Location = new Point(0, dpY),
                Width = 660,
                Height = 580,
                Visible = false,
                BackColor = SystemColors.Control,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            // Titel-Leiste im Device-Panel
            var titleBar = new Panel
            {
                Location = new Point(0, 0),
                Width = 660,
                Height = 36,
                BackColor = SystemColors.Control
            };
            var deviceTitleLabel = new Label
            {
                Name = "deviceTitleLabel",
                Text = "  Geräteinformationen",
                Location = new Point(4, 0),
                Size = new Size(450, 36),
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                TextAlign = ContentAlignment.MiddleLeft
            };
            var refreshDeviceBtn = new Button
            {
                Text = "Aktualisieren",
                Location = new Point(476, 5),
                Width = 120,
                Height = 26,
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText,
                BackColor = SystemColors.Control
            };
            refreshDeviceBtn.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(currentDisplayedIP))
                    ShowDeviceDetails(currentDisplayedIP);
            };
            titleBar.Controls.AddRange(new Control[] { deviceTitleLabel, refreshDeviceBtn });

            // Info-Leiste: IP / MAC / Typ
            var deviceInfoLabel = new Label
            {
                Name = "deviceInfoLabel",
                Text = "",
                Location = new Point(10, 42),
                Size = new Size(640, 18),
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText
            };

            // Hostname-Zeile
            var hostnameLabel = new Label
            {
                Text = "Hostname:",
                Location = new Point(10, 66),
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = SystemColors.ControlText
            };
            var hostnameBox = new TextBox
            {
                Name = "deviceHostnameBox",
                Location = new Point(86, 63),
                Width = 280,
                Font = new Font("Segoe UI", 9)
            };

            var saveHostnameBtn = new Button
            {
                Location = new Point(374, 62),
                Width = 110,
                Height = 25,
                Font = new Font("Segoe UI", 9)
            };
            SetButtonIcon(saveHostnameBtn, "Speichern", 20, "shell32.dll");
            saveHostnameBtn.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(hostnameBox.Text) ||
                    string.IsNullOrEmpty(currentDisplayedIP)) return;
                dbManager.UpdateDeviceHostname(currentDisplayedIP, hostnameBox.Text.Trim());
                statusLabel.Text = $"Hostname gesetzt: {hostnameBox.Text.Trim()}";
                ShowDeviceDetails(currentDisplayedIP);
                LoadCustomerTree();
            };
            var resetHostnameBtn = new Button
            {
                Location = new Point(492, 62),
                Width = 80,
                Height = 25,
                Font = new Font("Segoe UI", 9)
            };
            SetButtonIcon(resetHostnameBtn, "Auto", 238, "shell32.dll");
            resetHostnameBtn.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(currentDisplayedIP)) return;
                dbManager.ResetCustomHostname(currentDisplayedIP);
                statusLabel.Text = "Hostname wird beim nächsten Scan automatisch gesetzt";
            };

            // Hardware/Software-Scan-Button + Status
            var scanHwSwBtn = new Button
            {
                Name = "scanHwSwBtn",
                Location = new Point(10, 96),
                Width = 200,
                Height = 28,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            SetButtonIcon(scanHwSwBtn, "Hardware/SW abfragen", 109, "imageres.dll");
            scanHwSwBtn.Click += (s, e) => StartHwSwScanForDevice();

            // Rechtsklick-Menü: Scan-Typ manuell erzwingen
            var ctxScan = new ContextMenuStrip();
            ctxScan.Items.Add("🖥  Windows (WMI)",  null, (s, e) => StartHwSwScanForDevice(DeviceType.WindowsPC));
            ctxScan.Items.Add("🐧 Linux (SSH)",     null, (s, e) => StartHwSwScanForDevice(DeviceType.Linux));
            ctxScan.Items.Add("📱 Android (ADB)",   null, (s, e) => StartHwSwScanForDevice(DeviceType.Smartphone));
            ctxScan.Items.Add("🍎 iOS (USB)",       null, (s, e) => StartHwSwScanForDevice(DeviceType.MacOS));
            scanHwSwBtn.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                    ctxScan.Show(scanHwSwBtn, e.Location);
            };

            var scanHwSwStatus = new Label
            {
                Name = "scanHwSwStatus",
                Text = "",
                Location = new Point(218, 100),
                Size = new Size(420, 18),
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText
            };

            // Tabs: Hardware + Software
            var deviceTabs = new TabControl
            {
                Name = "deviceTabControl",
                Location = new Point(10, 130),
                Width = 638,
                Height = 435,
                Font = new Font("Segoe UI", 10)
            };

            var hwPage = new TabPage("  Hardware");
            var hwBox = new TextBox
            {
                Name = "deviceHardwareBox",
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9),
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.None
            };
            hwPage.Controls.Add(hwBox);

            var swPage = new TabPage("  Software");
            var swGrid = new DataGridView
            {
                Name = "deviceSoftwareGrid",
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                GridColor = Color.FromArgb(225, 230, 240),
                EnableHeadersVisualStyles = false
            };
            swGrid.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            swGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            swGrid.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Window;
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Software", DataPropertyName = "Name", FillWeight = 40 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", DataPropertyName = "Version", FillWeight = 18 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", DataPropertyName = "Publisher", FillWeight = 27 });
            swGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert", DataPropertyName = "InstallDate", FillWeight = 15 });
            swPage.Controls.Add(swGrid);

            deviceTabs.TabPages.Add(hwPage);
            deviceTabs.TabPages.Add(swPage);

            devicePanel.Controls.AddRange(new Control[] {
                titleBar, deviceInfoLabel,
                hostnameLabel, hostnameBox, saveHostnameBtn, resetHostnameBtn,
                scanHwSwBtn, scanHwSwStatus,
                deviceTabs
            });

            panel.Controls.AddRange(new Control[] {
                secLabel, sep1,
                nameLabel, nameTextBox, addressLabel, addressTextBox, saveBtn,
                sep2, ipSecLabel, ipGrid,
                wsLabel, wsBox, ipLbl2, ipBox, addIpBtn,
                sep3, removeBtn, moveBtn, importBtn,
                sep4, devicePanel
            });
            return panel;
        }


        // Setzt Icon + Text auf einen Button (Icon links, Text rechts)
        private static void SetButtonIcon(Button btn, string text, int iconIndex, string dll = "imageres.dll")
        {
            try
            {
                var bmp = ExtractDllIcon(dll, iconIndex);
                var img = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
                img.Images.Add(new Bitmap(bmp, 16, 16));
                btn.ImageList = img;
                btn.ImageIndex = 0;
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextAlign = ContentAlignment.MiddleRight;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            }
            catch { }
            btn.Text = text;
        }

        // Zeichnet ein einfaches Männchen-Icon als Bitmap
        private static Bitmap DrawPersonIcon(int size)
        {
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                var color = Color.FromArgb(70, 130, 180);
                using (var brush = new SolidBrush(color))
                using (var pen = new Pen(color, 1.5f))
                {
                    float s = size;
                    // Kopf
                    float headR = s * 0.18f;
                    float headX = s * 0.5f - headR;
                    float headY = s * 0.05f;
                    g.FillEllipse(brush, headX, headY, headR * 2, headR * 2);
                    // Körper
                    float bodyTop = headY + headR * 2 + s * 0.02f;
                    float bodyBottom = s * 0.72f;
                    float bodyW = s * 0.28f;
                    float bodyX = s * 0.5f - bodyW / 2;
                    g.FillRectangle(brush, bodyX, bodyTop,
                        bodyW, bodyBottom - bodyTop);
                    // Arme
                    pen.Width = size * 0.09f;
                    pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    float armY = bodyTop + (bodyBottom - bodyTop) * 0.25f;
                    g.DrawLine(pen, bodyX, armY,
                        s * 0.08f, armY + s * 0.18f);             // linker Arm
                    g.DrawLine(pen, bodyX + bodyW, armY,
                        s * 0.92f, armY + s * 0.18f);             // rechter Arm
                    // Beine
                    float legTop = bodyBottom;
                    g.DrawLine(pen, s * 0.5f - bodyW * 0.2f, legTop,
                        s * 0.28f, s * 0.97f);                    // linkes Bein
                    g.DrawLine(pen, s * 0.5f + bodyW * 0.2f, legTop,
                        s * 0.72f, s * 0.97f);                    // rechtes Bein
                }
            }
            return bmp;
        }

        // Standard Windows-Benutzer-Icon — nutzt das offizielle Benutzerkonten-Bild
        private static Bitmap GetStandardUserIcon(int size)
        {
            try
            {
                // Weg 1: Offizielles Benutzerbild aus Windows (user.bmp / user.png)
                string userPic = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    @"Microsoft\User Account Pictures\user.bmp");
                if (File.Exists(userPic))
                {
                    using (var orig = new Bitmap(userPic))
                    {
                        var bmp = new Bitmap(size, size);
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.DrawImage(orig, 0, 0, size, size);
                        }
                        return bmp;
                    }
                }

                // Weg 2: Shell32 Icon 265 = Standard-Benutzer (seit XP unverändert)
                return ExtractDllIcon("shell32.dll", 265);
            }
            catch
            {
                return ExtractDllIcon("shell32.dll", 265);
            }
        }

        // Extrahiert ein Icon aus einer beliebigen DLL (shell32.dll, imageres.dll etc.)
        [System.Runtime.InteropServices.DllImport("shell32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        private static Bitmap ExtractDllIcon(string dllName, int index)
        {
            try
            {
                string path = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System), dllName);
                IntPtr hIcon = ExtractIcon(IntPtr.Zero, path, index);
                if (hIcon != IntPtr.Zero)
                {
                    using (var icon = System.Drawing.Icon.FromHandle(hIcon))
                    {
                        var bmp = new Bitmap(24, 24);
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.Clear(Color.Transparent);
                            g.DrawIcon(icon, new Rectangle(0, 0, 24, 24));
                        }
                        return bmp;
                    }
                }
            }
            catch { }
            return new Bitmap(24, 24);
        }

        // Kompatibilitäts-Alias
        private static Bitmap ExtractShell32Icon(int index)
            => ExtractDllIcon("shell32.dll", index);

        private Label CreateStatusLabel()
            => new Label { Dock = DockStyle.Bottom, Height = 30, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };


        // =========================================================
        // === AUSWERTUNG / PIVOT TAB ===
        // =========================================================

        private Button calcExportButton;

        private TabPage CreateAuswertungTab()
        {
            var tab = new TabPage("Auswertung");

            var infoLabel = new Label
            {
                Text = "Exportiert Geräte, Software und Ports als XLSX und öffnet sie in LibreOffice Calc.",
                Dock = DockStyle.Top, Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Font = new Font("Segoe UI", 9),
                ForeColor = SystemColors.ControlText
            };

            calcExportButton = new Button
            {
                Text = "📊 In LibreOffice Calc öffnen",
                Dock = DockStyle.Top, Height = 44,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(220, 235, 255),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            calcExportButton.Click += (s, e) => ExportAndOpenInCalc();

            // ── Trennlinie ──
            var separator = new Label { Dock = DockStyle.Top, Height = 20 };

            var backupLabel = new Label
            {
                Text = "Datenbank sichern und wiederherstellen:",
                Dock = DockStyle.Top, Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = SystemColors.ControlText
            };

            var btnPanel = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(6, 4, 6, 4) };

            var btnExportDb = new Button
            {
                Text = "📥 Datenbank exportieren",
                Location = new Point(6, 8), Width = 220, Height = 34,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(220, 240, 220),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportDb.Click += (s, e) => ExportDatabase();

            var btnImportDb = new Button
            {
                Text = "📤 Datenbank importieren",
                Location = new Point(236, 8), Width = 220, Height = 34,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(255, 235, 210),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnImportDb.Click += (s, e) => ImportDatabase();

            btnPanel.Controls.Add(btnExportDb);
            btnPanel.Controls.Add(btnImportDb);

            // Dock-Reihenfolge: zuletzt hinzugefügt = oben
            tab.Controls.Add(btnPanel);
            tab.Controls.Add(backupLabel);
            tab.Controls.Add(separator);
            tab.Controls.Add(calcExportButton);
            tab.Controls.Add(infoLabel);
            return tab;
        }

        private void ExportAndOpenInCalc()
        {
            calcExportButton.Enabled = false;
            string originalText = calcExportButton.Text;

            try
            {
                // Sheet 1: Geräte
                calcExportButton.Text = "Exportiere... (1/3) Geräte";
                Application.DoEvents();
                var dtGeraete = dbManager.GetViewData("Inventar_Geraete");

                // Sheet 2: Software
                calcExportButton.Text = "Exportiere... (2/3) Software";
                Application.DoEvents();
                var dtSoftware = dbManager.GetViewData("Inventar_Software");

                // Sheet 3: Ports
                calcExportButton.Text = "Exportiere... (3/3) Ports";
                Application.DoEvents();
                var dtPorts = dbManager.GetViewData("Inventar_Ports");

                if (dtGeraete.Rows.Count == 0 && dtSoftware.Rows.Count == 0 && dtPorts.Rows.Count == 0)
                {
                    MessageBox.Show("Keine Daten vorhanden.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string filePath = Path.Combine(Path.GetTempPath(),
                    $"Inventarisierung_{DateTime.Now:yyyy-MM-dd_HHmmss}.xlsx");

                using (var wb = new XLWorkbook())
                {
                    AddSheet(wb, "Geräte", dtGeraete);
                    AddSheet(wb, "Software", dtSoftware);
                    AddSheet(wb, "Ports", dtPorts);
                    wb.SaveAs(filePath);
                }

                // LibreOffice Calc suchen und öffnen
                string scalc = FindLibreOfficeCalc();
                if (scalc != null)
                {
                    System.Diagnostics.Process.Start(scalc, "\"" + filePath + "\"");
                    statusLabel.Text = "✔ Export geöffnet in LibreOffice Calc";
                }
                else
                {
                    MessageBox.Show(
                        "LibreOffice wurde nicht gefunden.\n\nDie Datei wurde gespeichert unter:\n" + filePath,
                        "LibreOffice nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    statusLabel.Text = "Export gespeichert: " + filePath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Export:\n" + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Export fehlgeschlagen: " + ex.Message;
            }
            finally
            {
                calcExportButton.Enabled = true;
                calcExportButton.Text = originalText;
            }
        }

        private void AddSheet(XLWorkbook wb, string sheetName, System.Data.DataTable dt)
        {
            var ws = wb.Worksheets.Add(sheetName);

            // Header
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = dt.Columns[c].ColumnName;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(34, 84, 150);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // Daten
            for (int r = 0; r < dt.Rows.Count; r++)
                for (int c = 0; c < dt.Columns.Count; c++)
                    ws.Cell(r + 2, c + 1).Value = dt.Rows[r][c]?.ToString() ?? "";

            // Als Tabelle formatieren
            if (dt.Rows.Count > 0)
            {
                var tbl = ws.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count)
                            .CreateTable(sheetName.Replace(" ", ""));
                tbl.Theme = XLTableTheme.TableStyleMedium2;
            }

            ws.Columns().AdjustToContents(1, 50);
            ws.SheetView.FreezeRows(1);
        }

        private static string FindLibreOfficeCalc()
        {
            string[] searchPaths = {
                @"C:\Program Files\LibreOffice\program\scalc.exe",
                @"C:\Program Files (x86)\LibreOffice\program\scalc.exe",
                @"C:\Program Files\LibreOffice 7\program\scalc.exe",
                @"C:\Program Files\LibreOffice 24\program\scalc.exe",
                @"C:\Program Files\LibreOffice 25\program\scalc.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    @"Programs\LibreOffice\program\scalc.exe"),
            };

            foreach (var p in searchPaths)
                if (File.Exists(p)) return p;

            // Fallback: im PATH suchen
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo("where", "scalc.exe")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (var proc = System.Diagnostics.Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadLine();
                    proc.WaitForExit();
                    if (!string.IsNullOrEmpty(output) && File.Exists(output))
                        return output;
                }
            }
            catch { }

            return null;
        }

        // =========================================================
        // === DATENBANK BACKUP / RESTORE ===
        // =========================================================

        private void ExportDatabase()
        {
            var dbFiles = dbManager.GetAllDatabaseFiles();
            if (dbFiles.Count == 0)
            {
                MessageBox.Show("Keine Datenbanken gefunden.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var fbd = new FolderBrowserDialog { Description = "Zielordner für Datenbank-Backup wählen:" })
            {
                if (fbd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string backupDir = Path.Combine(fbd.SelectedPath, $"Inventarisierung_Backup_{DateTime.Now:yyyy-MM-dd_HHmmss}");
                    Directory.CreateDirectory(backupDir);

                    int count = 0;
                    foreach (var dbFile in dbFiles)
                    {
                        string destFile = Path.Combine(backupDir, Path.GetFileName(dbFile));
                        File.Copy(dbFile, destFile, true);
                        count++;
                    }

                    MessageBox.Show(
                        $"{count} Datenbank(en) exportiert nach:\n{backupDir}",
                        "Backup erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = $"✔ Backup: {count} DB(s) → {backupDir}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Export:\n" + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ImportDatabase()
        {
            var result = MessageBox.Show(
                "Beim Import werden die aktuellen Datenbanken überschrieben!\n\n" +
                "Es wird empfohlen, vorher ein Backup zu erstellen.\n\nFortfahren?",
                "Datenbank importieren", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result != DialogResult.Yes) return;

            using (var ofd = new OpenFileDialog
            {
                Title = "Datenbank-Backup-Ordner oder einzelne DB wählen",
                Filter = "SQLite Datenbank (*.db)|*.db",
                Multiselect = true
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string targetDir = Path.GetDirectoryName(dbManager.GetMainDatabasePath());
                    int count = 0;

                    foreach (var srcFile in ofd.FileNames)
                    {
                        string fileName = Path.GetFileName(srcFile);
                        if (!fileName.StartsWith("nmap_"))
                        {
                            statusLabel.Text = $"Übersprungen (kein nmap_*): {fileName}";
                            continue;
                        }
                        string destFile = Path.Combine(targetDir, fileName);
                        File.Copy(srcFile, destFile, true);
                        count++;
                    }

                    MessageBox.Show(
                        $"{count} Datenbank(en) importiert.\n\nDas Programm wird jetzt neu gestartet, um die Daten zu laden.",
                        "Import erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Neustart
                    Application.Restart();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Import:\n" + ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

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
            remoteHardwareButton.Enabled = false;
            BeginScanUI("Netzwerk-Discovery läuft (ARP-Scan)...");

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
                        // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                        DisplayDevicesWithStatus(currentDevices, oldDevices);
                        LoadDatabaseDevices();

                        if (!string.IsNullOrEmpty(currentDisplayedIP))
                            ShowDeviceDetails(currentDisplayedIP);

                        int found = currentDevices.Count;
                        int withMac = currentDevices.Count(d => !string.IsNullOrEmpty(d.MacAddress));
                        EndScanUI($"✅ Discovery abgeschlossen — {found} Geräte, {withMac} mit MAC");
                        scanButton.Enabled = true;
                        remoteHardwareButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new MethodInvoker(() =>
                    {
                        MessageBox.Show($"Fehler: {ex.Message}");
                        EndScanUI("Fehler beim Scan");
                        scanButton.Enabled = true;
                        remoteHardwareButton.Enabled = true;
                    }));
                }
            });
        }

        // Ladebalken + Sanduhr-Animation starten
        private void BeginScanUI(string message)
        {
            scanProgressBar.Visible = true;
            scanProgressBar.Value = 0;
            scanAnimTimer.Start();
            statusLabel.Text = "⏳  " + message;
            Cursor = Cursors.AppStarting;
        }

        // Ladebalken + Sanduhr-Animation stoppen
        private void EndScanUI(string message)
        {
            scanProgressBar.Visible = false;
            scanAnimTimer.Stop();
            statusLabel.Text = message;
            Cursor = Cursors.Default;
        }

        /// <summary>
        /// 🔬 FERNWARTUNGS-SCAN — scannt jedes Gerät einzeln mit Fortschrittsanzeige.
        /// Ports: RDP 3389, SSH 22, VNC 5900, WinRM 5985/5986, SMB 445, HTTP/S

        private void SyncDevicesToLocation(int locationID, List<DeviceInfo> devices)
        {
            var location = dbManager.GetLocationByID(locationID);
            if (location == null) { statusLabel.Text = "Fehler: Standort nicht gefunden!"; return; }

            // IPs die bereits dieser Location zugewiesen sind
            var assignedIPs = new HashSet<string>(
                dbManager.GetIPsWithWorkstationByLocation(locationID).Select(x => x.IPAddress));

            // MACs die bereits in dieser Location vorhanden sind (verhindert Doppelaufnahme bei IP-Wechsel)
            var assignedMACs = new HashSet<string>(
                dbManager.GetIPsWithWorkstationByLocation(locationID)
                    .Select(x => x.IPAddress)
                    .SelectMany(ip => dbManager.LoadDevices("Alle")
                        .Where(d => d.IP == ip && !string.IsNullOrEmpty(d.MacAddress))
                        .Select(d => d.MacAddress)));

            int added = 0, skipped = 0;

            foreach (var device in devices)
            {
                // IP bereits in dieser Location → nur LastSeen aktualisieren, nicht neu anlegen
                if (assignedIPs.Contains(device.IP))
                {
                    skipped++;
                    continue;
                }

                // MAC bereits in dieser Location (Gerät hat nur IP gewechselt) → überspringen
                if (!string.IsNullOrEmpty(device.MacAddress) && assignedMACs.Contains(device.MacAddress))
                {
                    skipped++;
                    continue;
                }

                // Wirklich neues Gerät → aufnehmen
                dbManager.AddIPToLocation(locationID, device.IP, device.Hostname ?? "");
                try { dbManager.SyncDeviceToCustomerDb(device, location.CustomerID); } catch { }
                added++;
            }

            statusLabel.Text = added > 0
                ? $"{added} neue Geräte zu '{location.Name}' hinzugefügt, {skipped} bereits vorhanden"
                : $"Alle {skipped} Geräte bereits in '{location.Name}' vorhanden";
        }

        /// Bereinigt Hostnamen die noch im alten "Vendor (MAC)" oder reinen MAC-Format gespeichert sind.
        /// Gibt einen leeren String zurück wenn kein echter Name vorhanden → Anzeige zeigt "-"
        /// Extrahiert "Hersteller : XYZ" aus HW-Info-Text (WMI/Linux/ADB).
        /// Gibt null zurück wenn nichts gefunden.
        private static string ExtractVendorFromHwInfo(string hwInfo)
        {
            if (string.IsNullOrEmpty(hwInfo)) return null;
            foreach (var line in hwInfo.Split('\n'))
            {
                var t = line.Trim();
                // WMI:   "Hersteller       : Dell Inc."
                // Linux: "  Hersteller : LENOVO"
                if (t.StartsWith("Hersteller") && t.Contains(":"))
                {
                    var val = t.Substring(t.IndexOf(':') + 1).Trim();
                    if (!string.IsNullOrEmpty(val) &&
                        val != "To Be Filled By O.E.M." &&
                        val != "Default string" &&
                        val.Length > 1)
                        return val;
                }
            }
            return null;
        }

        private static string CleanHostname(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            // Muster: "Vendor (AA:BB:CC:DD:EE:FF)" → ""  (Vendor kommt separat als Gerätetyp)
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^.+\s+\([0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}\)$")) return "";
            // Reines MAC-Format → ""
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^[0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}$")) return "";
            // IP-Adresse als Hostname → ""
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^\d{1,3}(\.\d{1,3}){3}$")) return "";
            return raw;
        }

        private void DisplayDevicesWithStatus(List<DeviceInfo> currentDevices, List<DeviceInfo> previousDevices)
        {
            deviceTable.Rows.Clear();
            var displayed   = new HashSet<string>();
            var locationIPs = dbManager.GetIPsWithWorkstationByLocation(selectedLocationID);

            // Einmal laden und für beide Zwecke verwenden
            var dbDevices   = dbManager.LoadDevices("Alle")
                                       .Where(d => locationIPs.Any(l => l.IPAddress == d.IP))
                                       .GroupBy(d => d.IP)
                                       .Select(g => g.First())
                                       .ToList();

            var lastScanIPs = new HashSet<string>(dbDevices.Select(d => d.IP));

            // Manuell vergebene Hostnamen: IP → Name aus DB
            var dbHostnames = dbDevices
                .Where(d => !string.IsNullOrEmpty(d.Hostname))
                .ToDictionary(d => d.IP, d => d.Hostname);

            // Gespeicherter Gerätetyp: IP → DeviceType aus DB
            var dbTypes = dbDevices
                .Where(d => d.DeviceType != DeviceType.Unbekannt)
                .ToDictionary(d => d.IP, d => d.DeviceType);

            // Vendor und Kommentar aus DB
            var dbVendors   = dbDevices.Where(d => !string.IsNullOrEmpty(d.Vendor))
                                       .ToDictionary(d => d.IP, d => d.Vendor);
            var dbComments  = dbDevices.ToDictionary(d => d.IP, d => d.Comment ?? "");

            var currentIPs  = new HashSet<string>(currentDevices.Select(d => d.IP));

            foreach (var dev in currentDevices)
                if (displayed.Add(dev.IP))
                {
                    string raw      = dbHostnames.TryGetValue(dev.IP, out string dbName) && !string.IsNullOrEmpty(dbName)
                                      ? dbName : dev.Hostname;
                    string hostname = CleanHostname(raw);
                    var dt          = dbTypes.TryGetValue(dev.IP, out DeviceType dbDt) ? dbDt : dev.DeviceType;
                    string vendor   = dbVendors.TryGetValue(dev.IP, out string dbVendor) && !string.IsNullOrEmpty(dbVendor)
                                      ? dbVendor : (dev.Vendor ?? "");
                    string comment  = dbComments.TryGetValue(dev.IP, out string c) ? c : "";
                    string mac      = !string.IsNullOrEmpty(dev.MacAddress) ? dev.MacAddress
                                      : dbDevices.FirstOrDefault(d => d.IP == dev.IP)?.MacAddress ?? "";
                    deviceTable.Rows.Add(
                        lastScanIPs.Contains(dev.IP) ? "AKTIV" : "NEU",
                        dev.IP,
                        string.IsNullOrEmpty(hostname) ? "-" : hostname,
                        $"{DeviceTypeHelper.GetIcon(dt)} {DeviceTypeHelper.GetLabel(dt)}",
                        vendor, mac, comment, dev.Status, dev.Ports);
                }

            foreach (var last in dbDevices)
                if (!currentIPs.Contains(last.IP) && displayed.Add(last.IP))
                {
                    string hostname = CleanHostname(last.Hostname ?? "");
                    deviceTable.Rows.Add(
                        "OFFLINE",
                        last.IP,
                        string.IsNullOrEmpty(hostname) ? "-" : hostname,
                        $"{DeviceTypeHelper.GetIcon(last.DeviceType)} {DeviceTypeHelper.GetLabel(last.DeviceType)}",
                        last.Vendor ?? "", last.MacAddress ?? "", last.Comment ?? "", "Down", "-");
                }
        }

        // =========================================================
        // === HARDWARE / SOFTWARE ===
        // =========================================================


        private void StartRemoteHardwareQuery()
        {
            using (var form = new RemoteConnectionForm(dbManager))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;

                // Einzelgerät oder Mehrfachauswahl aus Picker?
                var targets = new List<(string IP, string Username, string Password)>();

                if (form.SelectedTargets.Count > 0)
                    targets.AddRange(form.SelectedTargets);
                else if (!string.IsNullOrEmpty(form.ComputerIP))
                    targets.Add((form.ComputerIP, form.Username, form.Password));

                if (targets.Count == 0) return;

                remoteHardwareButton.Enabled = false;
                int total = targets.Count;
                int current = 0;

                BeginScanUI(total == 1
                    ? $"Verbinde zu {targets[0].IP}..."
                    : $"Starte Remote-Scan für {total} Geräte...");

                System.Threading.Tasks.Task.Run(() =>
                {
                    var errors = new List<string>();
                    int success = 0;

                    foreach (var target in targets)
                    {
                        current++;
                        Invoke(new MethodInvoker(() =>
                            statusLabel.Text = $"⏳  Gerät {current}/{total}: {target.IP}..."));

                        try
                        {
                            string hw = hardwareManager.GetRemoteHardwareInfo(target.IP, target.Username, target.Password);
                            var sw = softwareManager.GetRemoteSoftware(target.IP, target.Username, target.Password);
                            string pcName = target.IP;
                            foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                            dbManager.CheckForUpdates(sw, pcName);

                            Invoke(new MethodInvoker(() =>
                            {
                                hardwareInfoTextBox.Text = hw;
                                DisplaySoftwareGrid(sw, pcName);
                                dbManager.SaveSoftware(sw);
                                UpdateHardwareLabels(pcName, true);
                            }));
                            success++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{target.IP}: {ex.Message}");
                        }
                    }

                    Invoke(new MethodInvoker(() =>
                    {
                        remoteHardwareButton.Enabled = true;
                        if (errors.Count == 0)
                            EndScanUI(total == 1
                                ? "Remote-Abfrage abgeschlossen"
                                : $"✅ {success}/{total} Geräte erfolgreich abgefragt");
                        else
                        {
                            EndScanUI($"⚠️ {success}/{total} erfolgreich — {errors.Count} Fehler");
                            MessageBox.Show(
                                "Fehler bei folgenden Geräten:\n\n" + string.Join("\n", errors),
                                "Teilweise Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }));
                });
            }
        }

        /// <summary>
        /// Hardware/Software für das aktuell im Detail-Panel angezeigte Gerät abfragen.
        /// Lokal wenn IP = eigene IP, sonst Remote mit Anmeldedialog.
        /// Ergebnis wird in DB dem Gerät zugeordnet und im Detail-Panel angezeigt.
        /// </summary>
        private void StartHwSwScanForDevice(DeviceType? forceType = null)
        {
            if (string.IsNullOrEmpty(currentDisplayedIP))
            {
                MessageBox.Show("Bitte zuerst ein Gerät im Baum auswählen!",
                    "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string ip = currentDisplayedIP;

            // Prüfen ob es der lokale PC ist
            bool isLocal = IsLocalIP(ip);

            var scanBtn = FindControl<Button>("scanHwSwBtn");
            var scanStatus = FindControl<Label>("scanHwSwStatus");
            if (scanBtn != null) scanBtn.Enabled = false;
            if (scanStatus != null) scanStatus.Text = "Abfrage läuft...";

            if (isLocal)
            {
                // Lokaler Scan — kein Login nötig
                System.Threading.Tasks.Task.Run(() =>
                {
                    string hw = hardwareManager.GetHardwareInfo();
                    var sw = softwareManager.GetInstalledSoftware();
                    string pcName = ip; // per IP speichern für eindeutige Zuordnung
                    foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                    dbManager.CheckForUpdates(sw, pcName);
                    dbManager.SaveSoftware(sw);

                    // Auch Hostname in Devices aktualisieren
                    var device = dbManager.LoadDevices("Alle").FirstOrDefault(d => d.IP == ip);

                    if (device != null)
                    {
                        var hwInfo = $"=== Lokal abgefragt: {DateTime.Now:dd.MM.yyyy HH:mm} ===\n{hw}";
                        dbManager.SaveHardwareInfo(ip, hwInfo);
                    }
                    dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                    // Inventar_* sind jetzt VIEWs — kein Refresh nötig

                    Invoke(new MethodInvoker(() =>
                    {
                        hardwareInfoTextBox.Text = hw;
                        currentSoftware = sw;
                        DisplaySoftwareGrid(sw, ip);
                        UpdateHardwareLabels(Environment.MachineName, false);
                        ShowDeviceDetails(ip);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = $"✅ Abgefragt: {DateTime.Now:HH:mm}";
                        statusLabel.Text = $"Hardware/Software für {ip} gespeichert";
                    }));
                });
            }
            else
            {
                // Gerätetyp: manuell erzwungen oder automatisch aus DB
                var knownDevice = dbManager.LoadDevices("Alle").FirstOrDefault(d => d.IP == ip);
                var deviceType  = forceType ?? knownDevice?.DeviceType ?? DeviceType.Unbekannt;

                // iOS wird als MacOS-Override signalisiert (Rechtsklick-Menü)
                bool forcediOS = forceType == DeviceType.MacOS;
                bool isLinux   = deviceType == DeviceType.Linux;
                bool isAndroid = deviceType == DeviceType.Smartphone && !forcediOS && !IsAppleDevice(knownDevice);
                bool isApple   = forcediOS || (IsAppleDevice(knownDevice) && deviceType == DeviceType.Smartphone);

                // ── Linux via SSH ─────────────────────────────
                if (isLinux)
                {
                    using (var form = new LinuxSshConnectionForm())
                    {
                        form.SetIP(ip);
                        if (form.ShowDialog(this) != DialogResult.OK)
                        { if (scanBtn != null) scanBtn.Enabled = true; if (scanStatus != null) scanStatus.Text = ""; return; }

                        string sshHost = form.Host; string sshUser = form.Username;
                        string sshPass = form.Password; string sshKey = form.KeyFilePath;
                        int    sshPort = form.Port;

                        System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                string hw = linuxManager.GetHardwareInfo(sshHost, sshUser, sshPass, sshKey, sshPort);
                                var sw    = linuxManager.GetInstalledSoftware(sshHost, sshUser, sshPass, sshKey, sshPort);
                                foreach (var s in sw) s.PCName = ip;
                                dbManager.SaveHardwareInfo(ip, hw);
                                dbManager.SaveSoftware(sw);
                                dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                                // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                                Invoke(new MethodInvoker(() =>
                                {
                                    hardwareInfoTextBox.Text = hw; DisplaySoftwareGrid(sw, ip);
                                    ShowDeviceDetails(ip);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = $"✅ Linux SSH: {DateTime.Now:HH:mm}";
                                    statusLabel.Text = $"Linux-Abfrage für {ip} gespeichert";
                                }));
                            }
                            catch (Exception ex)
                            {
                                Invoke(new MethodInvoker(() =>
                                {
                                    MessageBox.Show($"SSH-Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                                }));
                            }
                        });
                    }
                }
                // ── Android via ADB ───────────────────────────
                else if (isAndroid || deviceType == DeviceType.Smartphone)
                {
                    if (!adbManager.IsAdbAvailable)
                    {
                        MessageBox.Show(
                            "ADB (Android Debug Bridge) wurde nicht gefunden.\n\n" +
                            "Bitte Android Platform Tools installieren:\n" +
                            "https://developer.android.com/tools/releases/platform-tools\n\n" +
                            "Dann adb.exe in denselben Ordner wie Inventarisierung.exe kopieren.",
                            "ADB nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = "";
                        return;
                    }

                    // Dialog: IP und Port vom Nutzer bestätigen lassen (Android 11+ hat anderen Port)
                    var adbDlg = new AdbConnectionDialog(ip);
                    if (adbDlg.ShowDialog(this) != System.Windows.Forms.DialogResult.OK)
                    {
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = "";
                        return;
                    }

                    string ipPort = adbDlg.IpPort;
                    bool doPair   = adbDlg.DoPairing;
                    string pairIp = $"{ip}:{adbDlg.PairPort}";
                    string pairCode = adbDlg.PairCode;

                    if (scanStatus != null) scanStatus.Text = $"Verbinde mit {ipPort}...";

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            // Ggf. erst koppeln (einmalig für Android 11+)
                            if (doPair)
                            {
                                string pairResult = adbManager.PairDevice(pairIp, pairCode);
                                Invoke(new MethodInvoker(() =>
                                    MessageBox.Show(pairResult, "Kopplung", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                            }

                            string hw = adbManager.GetDeviceInfo(ipPort);
                            var sw    = adbManager.GetInstalledApps(ipPort);
                            foreach (var s in sw) s.PCName = ip;
                            dbManager.SaveHardwareInfo(ip, hw);
                            dbManager.SaveSoftware(sw);
                            dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));

                            // Gerätenamen (z.B. "Samsung Galaxy A54") automatisch als Hostname setzen,
                            // aber nur wenn noch kein manueller Name vergeben wurde
                            string adbName = AdbManager.ExtractAdbDeviceName(hw);
                            if (!string.IsNullOrEmpty(adbName))
                                dbManager.SetHostnameIfEmpty(ip, adbName);

                            // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                            Invoke(new MethodInvoker(() =>
                            {
                                hardwareInfoTextBox.Text = hw; DisplaySoftwareGrid(sw, ip);
                                ShowDeviceDetails(ip);
                                LoadCustomerTree();
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = $"✅ Android ADB: {DateTime.Now:HH:mm}";
                                statusLabel.Text = $"Android-Abfrage für {ip} gespeichert";
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new MethodInvoker(() =>
                            {
                                MessageBox.Show($"ADB-Fehler:\n\n{ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                            }));
                        }
                    });
                }
                // ── iOS via USB ───────────────────────────────
                else if (isApple && deviceType == DeviceType.Smartphone)
                {
                    if (!iosManager.IsAvailable)
                    {
                        MessageBox.Show(
                            "libimobiledevice wurde nicht gefunden.\n\n" +
                            "Bitte ideviceinfo.exe herunterladen und neben Inventarisierung.exe ablegen:\n" +
                            "https://github.com/libimobiledevice-win32/imobiledevice-net/releases\n\n" +
                            "Außerdem: iPhone per USB anschließen und 'Vertrauen' antippen.",
                            "libimobiledevice nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = "";
                        return;
                    }

                    if (scanStatus != null) scanStatus.Text = "iOS-Abfrage via USB...";

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            string hw = iosManager.GetDeviceInfo();
                            var sw    = new List<SoftwareInfo> { new SoftwareInfo
                                { Name = "iOS: Softwareliste nicht verfügbar (kein Jailbreak)", Source = "iOS", PCName = ip } };
                            dbManager.SaveHardwareInfo(ip, hw);
                            dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                            // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                            Invoke(new MethodInvoker(() =>
                            {
                                hardwareInfoTextBox.Text = hw; DisplaySoftwareGrid(sw, ip);
                                ShowDeviceDetails(ip);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = $"✅ iOS: {DateTime.Now:HH:mm}";
                                statusLabel.Text = $"iOS-Abfrage gespeichert";
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new MethodInvoker(() =>
                            {
                                MessageBox.Show($"iOS-Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                            }));
                        }
                    });
                }
                // ── Windows WMI (Standard) ────────────────────
                else
                {
                    using (var form = new RemoteConnectionForm(dbManager))
                    {
                        form.SetIP(ip);
                        if (form.ShowDialog(this) != DialogResult.OK)
                        { if (scanBtn != null) scanBtn.Enabled = true; if (scanStatus != null) scanStatus.Text = ""; return; }

                        string username = form.Username;
                        string password = form.Password;

                        System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                string hw = hardwareManager.GetRemoteHardwareInfo(ip, username, password);
                                var sw    = softwareManager.GetRemoteSoftware(ip, username, password);
                                string pcName = ip;
                                foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                                dbManager.CheckForUpdates(sw, pcName);
                                dbManager.SaveSoftware(sw);
                                dbManager.SaveHardwareInfo(ip, $"=== Remote abgefragt: {DateTime.Now:dd.MM.yyyy HH:mm} ===\n{hw}");
                                dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                                // Inventar_* sind jetzt VIEWs — kein Refresh nötig

                                Invoke(new MethodInvoker(() =>
                                {
                                    hardwareInfoTextBox.Text = hw; DisplaySoftwareGrid(sw, ip);
                                    UpdateHardwareLabels(ip, true); ShowDeviceDetails(ip);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = $"✅ Remote abgefragt: {DateTime.Now:HH:mm}";
                                    statusLabel.Text = $"Hardware/Software für {ip} gespeichert";
                                }));
                            }
                            catch (Exception ex)
                            {
                                Invoke(new MethodInvoker(() =>
                                {
                                    MessageBox.Show($"Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                                }));
                            }
                        });
                    }
                }
            }
        }

        private static bool IsAppleDevice(DatabaseDevice d)
        {
            if (d == null) return false;
            string vendor = (d.MacAddress ?? "").ToLower();
            string host   = (d.Hostname ?? "").ToLower();
            return host.Contains("iphone") || host.Contains("ipad") ||
                   vendor.Contains("apple");
        }

        private static bool IsAppleVendor(DatabaseDevice d)
            => d != null && (d.Hostname ?? "").ToLower().Contains("apple");

        private bool IsLocalIP(string ip)
        {
            try
            {
                var localIPs = System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName())
                                    .Select(a => a.ToString());
                return ip == "127.0.0.1" || ip == "::1" || localIPs.Contains(ip);
            }
            catch { return false; }
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

        private void RefreshAllViews()
        {
            refreshButton.Enabled = false;
            refreshButton.Text = "⏳ Lädt...";
            try
            {
                // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                LoadDatabaseDevices();
                LoadDatabaseSoftware();
                LoadCustomerTree();
                statusLabel.Text = "✔ Daten aktualisiert — " + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Fehler beim Aktualisieren: " + ex.Message;
            }
            finally
            {
                refreshButton.Enabled = true;
                refreshButton.Text = "🔄 Aktualisieren";
            }
        }

        private void LoadDatabaseDevices(string filter = "Alle")
        {
            dbDeviceTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var dev in dbManager.LoadDevices(filter))
                if (displayed.Add(dev.IP))
                    dbDeviceTable.Rows.Add(dev.ID, dev.Zeitstempel, dev.IP, dev.Hostname, dev.MacAddress ?? "", dev.Status, dev.Ports,
                        $"{DeviceTypeHelper.GetIcon(dev.DeviceType)} {DeviceTypeHelper.GetLabel(dev.DeviceType)}");
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var sw in dbManager.LoadSoftware(filter))
            {
                string key = sw.PCName + "|" + sw.Name + "|" + sw.Version;
                if (displayed.Add(key))
                    dbSoftwareTable.Rows.Add(sw.ID, sw.DeviceID > 0 ? (object)sw.DeviceID : "", sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
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

            // Einmal laden — nicht für jede Location neu
            var allDevices = dbManager.LoadDevices("Alle");

            foreach (var customer in dbManager.GetCustomers())
            {
                var customerNode = new TreeNode($"👤 {customer.Name}")
                {
                    Tag = new NodeData { Type = "Customer", ID = customer.ID, Data = customer }
                };
                var locations = dbManager.GetLocationsByCustomer(customer.ID);
                for (int i = 0; i < locations.Count; i++)
                    AddLocationNodeRecursive(customerNode, locations[i], $"S{i + 1}", allDevices);
                treeView.Nodes.Add(customerNode);
            }
            treeView.ExpandAll();
            LoadLocationCombo();
        }

        private void AddLocationNodeRecursive(TreeNode parentNode, Location location, string shortID,
            List<DatabaseDevice> allDevices)
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
                AddLocationNodeRecursive(locationNode, children[i], childID, allDevices);
            }

            foreach (var ip in dbManager.GetIPsWithWorkstationByLocation(location.ID))
            {
                // Bevorzuge den in der Location eingetragenen Arbeitsplatznamen (selbst vergeben),
                // falls vorhanden. Sonst Hostname aus Devices-Tabelle verwenden.
                var device = allDevices?.FirstOrDefault(d => d.IP == ip.IPAddress);
                string label = !string.IsNullOrEmpty(ip.WorkstationName) ? ip.WorkstationName
                             : !string.IsNullOrEmpty(device?.Hostname)    ? device.Hostname
                             : "";

                // MAC und Hersteller aus Device-DB (enthält WMI-Hersteller) oder NmapDetail
                var nmapDetail  = dbManager.GetLatestNmapDetail(ip.IPAddress);
                string mac      = device?.MacAddress ?? nmapDetail?.MacAddress ?? "";
                string vendor   = !string.IsNullOrEmpty(device?.Vendor)      ? device.Vendor
                                : !string.IsNullOrEmpty(nmapDetail?.Vendor)  ? nmapDetail.Vendor
                                : OUIHelper.Lookup(mac) ?? "";

                // Gerätetyp für Icon
                var detectedType = device?.DeviceType ?? DeviceType.Unbekannt;
                if (detectedType == DeviceType.Unbekannt)
                {
                    var ports = dbManager.GetPortsByDevice(ip.IPAddress) ?? new List<NmapPort>();
                    detectedType = DeviceTypeHelper.Detect(new DeviceInfo
                    {
                        IP = ip.IPAddress, Hostname = label,
                        Vendor = vendor, OS = nmapDetail?.OS ?? nmapDetail?.OSDetails,
                        OpenPorts = ports
                    });
                }
                string typeIcon = DeviceTypeHelper.GetIcon(detectedType);

                // Format: Icon Name - IP - MAC - Hersteller
                // Name: manuell gesetzt > DB-Hostname > "Unbekannt"
                string name = !string.IsNullOrEmpty(label) ? label : "Unbekannt";

                var parts = new List<string> { $"{typeIcon} {name}", ip.IPAddress };
                if (!string.IsNullOrEmpty(mac))    parts.Add(mac);
                if (!string.IsNullOrEmpty(vendor)) parts.Add(vendor);

                string display = string.Join(" - ", parts);

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

            // Info-Zeile + Hostname-Box befüllen
            infoLabel.Text = $"IP: {device.IP}   |   MAC: {device.MacAddress ?? "unbekannt"}   |   " +
                             $"Hostname: {device.Hostname ?? "-"}   |   Zuletzt gesehen: {device.Zeitstempel}";
            var hostnameBox = FindControl<TextBox>("deviceHostnameBox");
            if (hostnameBox != null) hostnameBox.Text = device.Hostname ?? "";

            // ── Hardware / Nmap ──────────────────────────────────
            var hw = new System.Text.StringBuilder();
            hw.AppendLine("=== Geräteinformationen ===");
            hw.AppendLine($"IP-Adresse  : {device.IP}");
            hw.AppendLine($"MAC-Adresse : {device.MacAddress ?? "nicht ermittelt"}");
            hw.AppendLine($"Hostname    : {device.Hostname ?? "-"}");
            hw.AppendLine($"Status      : {device.Status ?? "-"}");
            hw.AppendLine($"Letzte Scan : {device.Zeitstempel}");

            // WMI Hardware-Info falls vorhanden (von Hardware/Software-Abfrage)
            string wmiInfo = dbManager.GetLatestHardwareInfo(ipAddress);
            if (!string.IsNullOrEmpty(wmiInfo))
            {
                hw.AppendLine();
                hw.AppendLine(wmiInfo);
            }

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
                // Fernwartungs-Übersicht zuerst
                hw.AppendLine();
                hw.AppendLine("=== Fernwartungs-Zugänge ===");
                var remoteMap = new Dictionary<int, string>
                {
                    {22,   "SSH"}, {23, "Telnet ⚠️"}, {3389, "RDP 🖥"},
                    {5900, "VNC"}, {5985, "WinRM"},   {5986, "WinRM-SSL"}
                };
                bool anyRemote = false;
                foreach (var p in ports)
                {
                    if (remoteMap.ContainsKey(p.Port))
                    {
                        hw.AppendLine($"  ✅ {remoteMap[p.Port],-15} Port {p.Port}  {p.Version}");
                        if (!string.IsNullOrEmpty(p.Banner))
                            hw.AppendLine($"     └ {p.Banner}");
                        anyRemote = true;
                    }
                }
                if (!anyRemote)
                    hw.AppendLine("  ❌ Keine Fernwartungs-Ports gefunden");

                // Alle Ports tabellarisch
                hw.AppendLine();
                hw.AppendLine("=== Alle offenen Ports ===");
                hw.AppendLine($"  {"Port",-10} {"Protokoll",-6} {"Service",-20} {"Version"}");
                hw.AppendLine($"  {new string('-', 65)}");
                foreach (var p in ports)
                {
                    hw.AppendLine($"  {p.Port + "/" + p.Protocol,-10} {"open",-6} {p.Service,-20} {p.Version}");
                    if (!string.IsNullOrEmpty(p.Banner))
                        hw.AppendLine($"     └ {p.Banner}");
                }
            }
            else if (!string.IsNullOrEmpty(device.Ports) && device.Ports != "-")
            {
                hw.AppendLine();
                hw.AppendLine("=== Ports (aus letztem Scan) ===");
                foreach (var portStr in device.Ports.Split(','))
                    hw.AppendLine($"  {portStr.Trim()}");
            }
            else
            {
                hw.AppendLine();
                hw.AppendLine("ℹ️  Noch kein Fernwartungs-Scan durchgeführt.");
                hw.AppendLine("   → '🔬 Fernwartungs-Scan' klicken für Port-Details.");
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

            // SNMP-Daten (aus aktuellem Scan-Speicher — nicht persistiert)
            var snmpDevice = currentDevices.FirstOrDefault(d => d.IP == ipAddress);
            if (snmpDevice?.SnmpData != null)
            {
                hw.AppendLine();
                if (snmpDevice.SnmpData.Success)
                {
                    string ver = snmpDevice.SnmpData.SnmpVersion ?? "?";
                    hw.AppendLine($"=== SNMP-Informationen (SNMP {ver}) ===");
                    hw.AppendLine($"  sysName     : {snmpDevice.SnmpData.SysName ?? "-"}");
                    hw.AppendLine($"  sysDescr    : {snmpDevice.SnmpData.SysDescr ?? "-"}");
                    hw.AppendLine($"  sysLocation : {snmpDevice.SnmpData.SysLocation ?? "-"}");
                    hw.AppendLine($"  sysContact  : {snmpDevice.SnmpData.SysContact ?? "-"}");
                    hw.AppendLine($"  sysUpTime   : {snmpDevice.SnmpData.SysUpTime ?? "-"}");
                    hw.AppendLine($"  sysObjectID : {snmpDevice.SnmpData.SysObjectID ?? "-"}");
                    hw.AppendLine($"  Abfragetime : {snmpDevice.SnmpData.QueryTime:dd.MM.yyyy HH:mm}");
                }
                else
                {
                    hw.AppendLine("=== SNMP — Keine Antwort ===");
                    hw.AppendLine($"  Fehler: {snmpDevice.SnmpData.ErrorMessage}");
                }
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

        // ─────────────────────────────────────────────────────────
        // 📡 SNMP-SCAN — fragt alle aktuell bekannten Geräte ab
        // ─────────────────────────────────────────────────────────
        private void StartSnmpScan()
        {
            if (currentDevices.Count == 0)
            {
                MessageBox.Show("Bitte zuerst einen Netzwerk-Scan durchführen.", "Kein Scan vorhanden",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            snmpScanButton.Enabled = false;
            string snmpVerLabel = snmpSettings.Version == 3 ? "v3" : snmpSettings.Version == 1 ? "v1" : "v2c";
            BeginScanUI($"SNMP-Scan läuft ({snmpVerLabel}, Community: {snmpSettings.Community})...");

            var devicesToQuery = currentDevices.ToList();

            System.Threading.Tasks.Task.Run(() =>
            {
                int success = 0;
                snmpManager.QueryDevices(devicesToQuery, snmpSettings, msg =>
                {
                    Invoke(new MethodInvoker(() => statusLabel.Text = "⏳  " + msg));
                });

                foreach (var d in devicesToQuery)
                    if (d.SnmpData?.Success == true) success++;

                Invoke(new MethodInvoker(() =>
                {
                    EndScanUI($"✅ SNMP-Scan abgeschlossen — {success}/{devicesToQuery.Count} Geräte antworteten");
                    snmpScanButton.Enabled = true;
                    if (!string.IsNullOrEmpty(currentDisplayedIP))
                        ShowDeviceDetails(currentDisplayedIP);
                }));
            });
        }

        private void OpenSnmpSettings()
        {
            using (var dlg = new SnmpSettingsDialog(snmpSettings))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    snmpSettings = dlg.Settings;
            }
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
                    list.Add(new DatabaseSoftware { ID = Convert.ToInt32(row.Cells[0].Value), DeviceID = row.Cells[1].Value is int did ? did : 0, Zeitstempel = row.Cells[2].Value?.ToString(), PCName = row.Cells[3].Value?.ToString(), Name = row.Cells[4].Value?.ToString(), Version = row.Cells[5].Value?.ToString(), Publisher = row.Cells[6].Value?.ToString(), InstallDate = row.Cells[7].Value?.ToString(), LastUpdate = row.Cells[8].Value?.ToString() });
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
        public string FilePath { get; set; }
        public override string ToString() => Display;
    }

    // OUI helper: simple vendor lookup from MAC OUI (first 3 bytes)
    public static class OUIHelper
    {
        private static readonly Dictionary<string, string> builtIn = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "00:1A:2B", "Apple, Inc." },
            { "00:1A:79", "Samsung Electronics" },
            { "00:0C:29", "VMware, Inc." },
            { "84:38:35", "Google, Inc." },
            { "3C:5A:B4", "Xiaomi Communications" }
        };

        public static string Lookup(string mac)
        {
            if (string.IsNullOrWhiteSpace(mac)) return null;
            try
            {
                var clean = mac.Trim().ToUpperInvariant().Replace('-', ':').Replace('.', ':');
                clean = new string(clean.Where(c => c == ':' || (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F')).ToArray());
                if (!clean.Contains(":") && clean.Length >= 12)
                    clean = string.Join(":", Enumerable.Range(0, 6).Select(i => clean.Substring(i * 2, 2)));
                var parts = clean.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 3) return null;
                var key = string.Join(":", parts.Take(3));
                if (builtIn.TryGetValue(key, out var vendor)) return vendor;
                return null;
            }
            catch { return null; }
        }
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
            // SQLite.Interop.dll aus eingebetteter Resource ins App-Verzeichnis extrahieren,
            // damit es von System.Data.SQLite über den Standard-DLL-Suchpfad gefunden wird.
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string arch = Environment.Is64BitProcess ? "x64" : "x86";
            string rootDll = Path.Combine(baseDir, "SQLite.Interop.dll");
            string archDir = Path.Combine(baseDir, arch);
            string archDll = Path.Combine(archDir, "SQLite.Interop.dll");
            string resourceName = "SQLite.Interop." + arch + ".dll";
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            // In App-Root extrahieren (Standard-DLL-Suchpfad)
            if (!File.Exists(rootDll))
            {
                using (var stream = asm.GetManifestResourceStream(resourceName))
                using (var fs = File.Create(rootDll))
                    stream.CopyTo(fs);
            }
            // Auch in arch-spezifischen Unterordner extrahieren (Fallback)
            if (!File.Exists(archDll))
            {
                Directory.CreateDirectory(archDir);
                using (var stream = asm.GetManifestResourceStream(resourceName))
                using (var fs = File.Create(archDll))
                    stream.CopyTo(fs);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}