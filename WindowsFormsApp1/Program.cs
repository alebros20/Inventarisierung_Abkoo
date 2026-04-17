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
    public partial class MainForm : Form
    {
        // === UI COMPONENTS ===
        private TabControl tabControl;
        private TextBox networkTextBox;
        private Button scanButton, remoteHardwareButton, exportButton, refreshButton;
        private Label statusLabel;
        private ProgressBar scanProgressBar;
        private System.Windows.Forms.Timer scanAnimTimer;
        private DataGridView deviceTable, dbDeviceTable, dbSoftwareTable;
        private TextBox rawOutputTextBox;
        private ComboBox locationComboBox;
        private Button newLocationButton;

        // === MANAGERS ===
        private NmapScanner nmapScanner;
        private DatabaseManager dbManager;
        private LinuxManager linuxManager;
        private AdbManager adbManager;
        private IosManager iosManager;
        private SnmpManager snmpManager;
        private SnmpSettings snmpSettings = new SnmpSettings();

        // === DATA ===
        private List<DeviceInfo> currentDevices = new List<DeviceInfo>();
        private List<DeviceInfo> lastScanDevices = new List<DeviceInfo>();
        private int selectedLocationID = -1;
        private string currentDisplayedIP = "";

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
                    "Nmap nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            dbManager   = new DatabaseManager();
            snmpManager = new SnmpManager();
            linuxManager = new LinuxManager();
            adbManager  = new AdbManager();
            iosManager  = new IosManager();

            InitializeComponent();
            dbManager.InitializeDatabase();
            LoadCustomerTree();
            ApplyAutoSizing();
            this.Load += (s, e) => RestoreSplitterPosition();
        }

        private void RestoreSplitterPosition()
        {
            var split = Controls.Find("customerSplitContainer", true).FirstOrDefault() as SplitContainer;
            if (split == null) return;
            split.SplitterDistance = split.Width / 2;
        }

        private void ApplyAutoSizing()
        {
            int screenWidth  = Screen.PrimaryScreen.WorkingArea.Width;
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

            tabControl = CreateTabControl();
            var topPanel = CreateTopPanel();
            statusLabel = CreateStatusLabel();

            scanProgressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 6,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 25,
                Visible = false
            };

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

            newLocationButton = new Button { Location = new Point(310, 43), Width = 140, Height = 28, Font = new Font("Segoe UI", 10) };
            SetButtonIcon(newLocationButton, "+ Neuer Standort", 3);
            newLocationButton.Click += (s, e) => CreateNewLocation();

            scanButton = new Button { Location = new Point(460, 8), Width = 165, Height = 28, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            SetButtonIcon(scanButton, "Scan starten", 168);
            scanButton.Click += (s, e) => OpenScanDialog();

            remoteHardwareButton = new Button();

            exportButton = new Button { Location = new Point(634, 8), Width = 130, Height = 28, Font = new Font("Segoe UI", 10) };
            SetButtonIcon(exportButton, "Exportieren", 162);
            exportButton.Click += (s, e) => ExportData();

            refreshButton = new Button { Location = new Point(773, 8), Width = 140, Height = 28, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            refreshButton.Text = "🔄 Aktualisieren";
            refreshButton.Click += (s, e) => RefreshAllViews();

            var credentialButton = new Button { Location = new Point(460, 42), Width = 160, Height = 28, Font = new Font("Segoe UI", 10) };
            credentialButton.Text = "🔑 Passwort-Vorlagen";
            credentialButton.Click += (s, e) => { using (var dlg = new CredentialTemplateDialog(dbManager)) dlg.ShowDialog(this); };

            var snmpSettingsBtn = new Button { Location = new Point(634, 42), Width = 160, Height = 28, Font = new Font("Segoe UI", 10) };
            snmpSettingsBtn.Text = "⚙️ SNMP-Einstellungen";
            snmpSettingsBtn.Click += (s, e) => OpenSnmpSettings();

            topPanel.Controls.AddRange(new Control[] { networkLabel, networkTextBox, locationLabel, locationComboBox, newLocationButton, scanButton, exportButton, refreshButton, credentialButton, snmpSettingsBtn });
            return topPanel;
        }

        private void OpenScanDialog()
        {
            using (var dlg = new ScanDialog())
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                switch (dlg.SelectedMode)
                {
                    case ScanDialog.ScanMode.NetworkDiscovery: StartScan(); break;
                    case ScanDialog.ScanMode.RemoteScan:       StartRemoteHardwareQuery(); break;
                    case ScanDialog.ScanMode.SnmpScan:         StartSnmpScan(); break;
                }
            }
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

            var tabImages = new ImageList { ImageSize = new Size(24, 24), ColorDepth = ColorDepth.Depth32Bit };
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 109)); // 0: Geräte
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 168)); // 1: Nmap Ausgabe
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 27));  // 2: DB Geräte
            tabImages.Images.Add(ExtractDllIcon("shell32.dll",  20));  // 3: DB Software
            tabImages.Images.Add(GetStandardUserIcon(24));              // 4: Kunden
            tabImages.Images.Add(ExtractDllIcon("imageres.dll", 8));   // 5: Auswertung
            tabControl.ImageList = tabImages;

            var tabToolTip = new ToolTip { ShowAlways = true, InitialDelay = 300, AutoPopDelay = 5000 };
            string[] tabNames = { "Geräte", "Nmap Ausgabe", "DB - Geräte", "DB - Software", "Kunden / Standorte", "Auswertung" };

            tabControl.DrawItem += (s, e) =>
            {
                var rect = e.Bounds;
                bool sel = (e.Index == tabControl.SelectedIndex);
                e.Graphics.FillRectangle(sel ? SystemBrushes.Window : SystemBrushes.Control, rect);
                if (tabControl.ImageList != null && e.Index < tabNames.Length)
                {
                    var img = tabControl.ImageList.Images[e.Index];
                    if (img != null)
                    {
                        int ix = rect.Left + (rect.Width - img.Width) / 2;
                        int iy = rect.Top  + (rect.Height - img.Height) / 2;
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
            tabControl.TabPages.Add(CreateDBDeviceTab());
            tabControl.TabPages.Add(CreateDBSoftwareTab());
            tabControl.TabPages.Add(CreateCustomerLocationTab());
            tabControl.TabPages.Add(CreateAuswertungTab());

            for (int i = 0; i < tabControl.TabPages.Count; i++)
                tabControl.TabPages[i].ImageIndex = i;

            return tabControl;
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
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string arch    = Environment.Is64BitProcess ? "x64" : "x86";
            string rootDll = Path.Combine(baseDir, "SQLite.Interop.dll");
            string archDir = Path.Combine(baseDir, arch);
            string archDll = Path.Combine(archDir, "SQLite.Interop.dll");
            string resourceName = "SQLite.Interop." + arch + ".dll";
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            if (!File.Exists(rootDll))
            {
                using (var stream = asm.GetManifestResourceStream(resourceName))
                using (var fs = File.Create(rootDll))
                    stream.CopyTo(fs);
            }
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
