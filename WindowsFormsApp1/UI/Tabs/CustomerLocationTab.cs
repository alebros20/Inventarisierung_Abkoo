using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public partial class MainForm
    {
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

            var treeActionPanel = new Panel { Dock = DockStyle.Top, Height = 36 };
            var btnTestCreds = new Button
            {
                Text = "🔑 Zugangsdaten testen",
                Location = new Point(5, 4), Width = 170, Height = 26,
                Font = new Font("Segoe UI", 8, FontStyle.Bold)
            };
            btnTestCreds.Click += (s, e) => StartCredentialScan();
            var btnResetCreds = new Button
            {
                Text = "↺ Zugangsdaten zurücksetzen",
                Location = new Point(180, 4), Width = 190, Height = 26,
                Font = new Font("Segoe UI", 8)
            };
            btnResetCreds.Click += (s, e) => ResetSelectedDeviceCredential();
            treeActionPanel.Controls.AddRange(new Control[] { btnTestCreds, btnResetCreds });

            var leftPanel = new Panel { Dock = DockStyle.Fill };
            leftPanel.Controls.Add(treeView);
            leftPanel.Controls.Add(treeActionPanel);
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
                ForeColor = SystemColors.ControlText
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

        private void StartCredentialScan()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;

            List<DatabaseDevice> devices;
            if (nodeData?.Type == "Customer")
                devices = dbManager.GetDevicesByCustomer(nodeData.ID);
            else if (nodeData?.Type == "Location")
                devices = dbManager.GetDevicesByLocationRecursive(nodeData.ID);
            else
            {
                MessageBox.Show("Bitte wähle einen Kunden oder Standort im Baum aus.",
                    "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (devices.Count == 0)
            {
                MessageBox.Show("Keine Geräte an diesem Standort/Kunden gefunden.",
                    "Keine Geräte", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dlg = new CredentialScanDialog(dbManager, devices))
                dlg.ShowDialog(this);
        }

        private void ResetSelectedDeviceCredential()
        {
            var treeView = FindControl<TreeView>("customerTreeView");
            var nodeData = treeView?.SelectedNode?.Tag as NodeData;

            List<DatabaseDevice> devices = null;
            if (nodeData?.Type == "Customer")
                devices = dbManager.GetDevicesByCustomer(nodeData.ID);
            else if (nodeData?.Type == "Location")
                devices = dbManager.GetDevicesByLocationRecursive(nodeData.ID);
            else
            {
                MessageBox.Show("Bitte wähle einen Kunden oder Standort im Baum aus.", "Hinweis");
                return;
            }

            int count = devices.Count(d => d.CredentialTemplateID.HasValue);
            if (count == 0)
            {
                MessageBox.Show("Keine Geräte mit Zugangsdaten gefunden.", "Info");
                return;
            }

            if (MessageBox.Show($"Zugangsdaten für {count} Geräte zurücksetzen?", "Bestätigen",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            foreach (var d in devices.Where(d => d.CredentialTemplateID.HasValue))
                dbManager.ResetDeviceCredential(d.ID);

            statusLabel.Text = $"✔ Zugangsdaten für {count} Geräte zurückgesetzt";
        }

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
    }
}
