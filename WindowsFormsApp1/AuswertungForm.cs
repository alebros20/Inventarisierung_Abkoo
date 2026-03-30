using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
// iTextSharp wird vollqualifiziert verwendet um Konflikte mit System.Drawing zu vermeiden
using itextPdf = iTextSharp.text.pdf;
using itextText = iTextSharp.text;

namespace NmapInventory
{
    /// <summary>
    /// Pivot-Auswertung: Geräte × Software × Hardware quer über alle Kunden/Standorte.
    /// Zeilen   = Geräte (gruppierbar nach Kunde / Standort / Gerätetyp / OS)
    /// Spalten  = frei wählbar: Software-Namen, Hardware-Werte, oder beides
    /// Werte    = ✔ vorhanden / – nicht vorhanden / Versions-String
    /// </summary>
    public class AuswertungForm : Form
    {
        private readonly DatabaseManager _db;
        private readonly List<string> _dbPaths;

        // Standard-Konstruktor — nur Haupt-DB, zeigt das Fenster
        public AuswertungForm(DatabaseManager db) : this(db, null) { }

        // Konstruktor mit ausgewählten DB-Pfaden
        public AuswertungForm(DatabaseManager db, List<string> dbPaths)
        {
            _db = db;
            _dbPaths = dbPaths ?? new List<string>();
            Text = "Auswertung / Pivot";
            Size = new Size(1200, 750);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(900, 500);
            Font = new Font("Segoe UI", 9);
            BuildUI();
        }

        // Direkt exportieren + Calc öffnen ohne Fenster anzuzeigen
        public void ExportAndOpenCalc()
        {
            // Controls müssen existieren damit lblStatus befüllt werden kann
            if (lblStatus == null) BuildUI();
            ExportAllDataToExcel();
        }

        // ── UI-Controls ───────────────────────────────────────
        private ComboBox cmbZeilenGruppe;
        private ComboBox cmbSpaltenTyp;
        private TextBox txtSpaltenFilter;
        private TextBox txtZeilenFilter;
        private CheckBox chkNurVorhanden;
        private Button btnAuswerten;
        private Button btnExport;
        private Label lblStatus;
        private DataGridView grid;
        private Panel topPanel;

        // ── UI-Aufbau ─────────────────────────────────────────
        private void BuildUI()
        {
            // ── Top-Panel ────────────────────────────────────
            topPanel = new Panel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(8) };

            // Zeile 1: Gruppierung + Spaltentyp
            var lbl1 = new Label { Text = "Zeilen gruppieren nach:", Location = new Point(8, 10), AutoSize = true };
            cmbZeilenGruppe = new ComboBox
            {
                Location = new Point(175, 7),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbZeilenGruppe.Items.AddRange(new[] { "Gerät (IP)", "Hostname", "Kunde", "Gerätetyp", "Betriebssystem" });
            cmbZeilenGruppe.SelectedIndex = 1;

            var lbl2 = new Label { Text = "Spalten:", Location = new Point(340, 10), AutoSize = true };
            cmbSpaltenTyp = new ComboBox
            {
                Location = new Point(395, 7),
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSpaltenTyp.Items.AddRange(new[] {
                "Software (installiert?)",
                "Software (Version)",
                "Hardware: RAM",
                "Hardware: OS",
                "Hardware: CPU",
                "Hardware: Disk"
            });
            cmbSpaltenTyp.SelectedIndex = 0;

            // Zeile 2: Filter-Felder
            var lbl3 = new Label { Text = "Spalten-Filter (z.B. Word):", Location = new Point(8, 45), AutoSize = true };
            txtSpaltenFilter = new TextBox { Location = new Point(175, 42), Width = 200, ForeColor = Color.Gray, Text = "Alle Spalten" };
            txtSpaltenFilter.GotFocus += (s, e) => { if (txtSpaltenFilter.ForeColor == Color.Gray) { txtSpaltenFilter.Text = ""; txtSpaltenFilter.ForeColor = SystemColors.WindowText; } };
            txtSpaltenFilter.LostFocus += (s, e) => { if (string.IsNullOrEmpty(txtSpaltenFilter.Text)) { txtSpaltenFilter.Text = "Alle Spalten"; txtSpaltenFilter.ForeColor = Color.Gray; } };

            var lbl4 = new Label { Text = "Zeilen-Filter:", Location = new Point(390, 45), AutoSize = true };
            txtZeilenFilter = new TextBox { Location = new Point(465, 42), Width = 200, ForeColor = Color.Gray, Text = "Alle Geräte" };
            txtZeilenFilter.GotFocus += (s, e) => { if (txtZeilenFilter.ForeColor == Color.Gray) { txtZeilenFilter.Text = ""; txtZeilenFilter.ForeColor = SystemColors.WindowText; } };
            txtZeilenFilter.LostFocus += (s, e) => { if (string.IsNullOrEmpty(txtZeilenFilter.Text)) { txtZeilenFilter.Text = "Alle Geräte"; txtZeilenFilter.ForeColor = Color.Gray; } };

            chkNurVorhanden = new CheckBox
            {
                Text = "Nur Geräte mit Treffer",
                Location = new Point(680, 44),
                AutoSize = true
            };

            // Zeile 3: Buttons + Status
            btnAuswerten = new Button
            {
                Text = "▶ Auswerten",
                Location = new Point(8, 72),
                Width = 120,
                Height = 24
            };
            btnAuswerten.Click += (s, e) => RunPivot();

            btnExport = new Button
            {
                Text = "Exportieren (Calc / PDF / CSV)",
                Location = new Point(136, 72),
                Width = 220,
                Height = 24
            };
            btnExport.Click += (s, e) => ExportCsv();

            var btnOpenXml = new Button
            {
                Text = "Alle DB-Daten in Calc öffnen",
                Location = new Point(364, 72),
                Width = 210,
                Height = 24,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnOpenXml.Click += (s, e) => ExportAllDataToExcel();

            lblStatus = new Label
            {
                Location = new Point(582, 76),
                AutoSize = true,
                ForeColor = Color.DarkSlateGray
            };

            topPanel.Controls.AddRange(new Control[] {
                lbl1, cmbZeilenGruppe, lbl2, cmbSpaltenTyp,
                lbl3, txtSpaltenFilter, lbl4, txtZeilenFilter, chkNurVorhanden,
                btnAuswerten, btnExport, btnOpenXml, lblStatus
            });

            // ── DataGridView ──────────────────────────────────
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                BackgroundColor = SystemColors.Window,
                GridColor = Color.FromArgb(220, 220, 220),
                Font = new Font("Segoe UI", 9),
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                EnableHeadersVisualStyles = false
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 252);

            // Zellen einfärben: ✔ grün, – grau, Versionen normal
            grid.CellFormatting += Grid_CellFormatting;

            Controls.Add(grid);
            Controls.Add(topPanel);

            // Enter in Filter-Box triggert Auswertung
            txtSpaltenFilter.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) RunPivot(); };
            txtZeilenFilter.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) RunPivot(); };
        }

        // ── Pivot-Logik ───────────────────────────────────────
        private void RunPivot()
        {
            btnAuswerten.Enabled = false;
            lblStatus.Text = "Lade Daten...";
            Application.DoEvents();

            try
            {
                string spaltenTyp = cmbSpaltenTyp.SelectedItem?.ToString() ?? "";
                string zeilenGruppe = cmbZeilenGruppe.SelectedItem?.ToString() ?? "";
                string spaltenFilter = txtSpaltenFilter.Text.Trim().ToLower();
                string zeilenFilter = txtZeilenFilter.Text.Trim().ToLower();

                // ── Daten laden ───────────────────────────────
                var devices = _db.LoadDevices("Alle");
                var software = _db.LoadSoftware("Alle");
                var kunden = _db.GetCustomers();

                // Kunden-Lookup per Hostname/IP (über LocationIPs)
                var kundenMap = BuildKundenMap(kunden);

                // ── Zeilen bestimmen ──────────────────────────
                // Jede Zeile = ein Eintrag nach der gewählten Gruppierung
                var rows = BuildRows(devices, zeilenGruppe, kundenMap);

                // Zeilen-Filter anwenden
                if (!string.IsNullOrEmpty(zeilenFilter) && zeilenFilter != "alle geräte")
                    rows = rows.Where(r => r.Label.ToLower().Contains(zeilenFilter)).ToList();

                // ── Spalten bestimmen ─────────────────────────
                List<string> columns;
                if (spaltenTyp.StartsWith("Software"))
                {
                    // Alle eindeutigen Software-Namen als Spalten
                    columns = software
                        .Select(s => s.Name ?? "")
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(n => n)
                        .ToList();
                }
                else
                {
                    // Hardware-Spalten sind fix
                    columns = new List<string> { "OS", "RAM", "CPU", "Disk C:" };
                }

                // Spalten-Filter
                if (!string.IsNullOrEmpty(spaltenFilter) && spaltenFilter != "alle spalten")
                    columns = columns.Where(c => c.ToLower().Contains(spaltenFilter)).ToList();

                if (columns.Count > 200)
                    columns = columns.Take(200).ToList(); // Limit für Performance

                // ── DataTable aufbauen ────────────────────────
                var dt = new DataTable();
                dt.Columns.Add(GetZeilenHeader(zeilenGruppe), typeof(string));
                dt.Columns.Add("IP", typeof(string));

                foreach (var col in columns)
                    dt.Columns.Add(col, typeof(string));

                // Software nach Gerät (PCName) indexieren
                var swByDevice = software
                    .GroupBy(s => s.PCName ?? "", StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

                // Hardware aus HardwareInfo laden
                var hwByDevice = LoadHardwareValues(devices);

                // ── Zeilen befüllen ───────────────────────────
                foreach (var row in rows)
                {
                    // Software-Lookup: PCName = Hostname oder IP
                    swByDevice.TryGetValue(row.Hostname ?? "", out var devSw);
                    if (devSw == null) swByDevice.TryGetValue(row.IP ?? "", out devSw);
                    devSw = devSw ?? new List<DatabaseSoftware>();

                    hwByDevice.TryGetValue(row.IP ?? "", out var hw);
                    hw = hw ?? new HardwareValues();

                    bool hatTreffer = false;
                    var dr = dt.NewRow();
                    dr[0] = row.Label;
                    dr[1] = row.IP;

                    for (int i = 0; i < columns.Count; i++)
                    {
                        string col = columns[i];
                        string val = "";

                        if (spaltenTyp.StartsWith("Software"))
                        {
                            var match = devSw.FirstOrDefault(s =>
                                string.Equals(s.Name, col, StringComparison.OrdinalIgnoreCase));
                            if (match != null)
                            {
                                val = spaltenTyp.Contains("Version") ? (match.Version ?? "✔") : "✔";
                                hatTreffer = true;
                            }
                            else val = "–";
                        }
                        else
                        {
                            // Hardware
                            switch (col)
                            {
                                case "OS": val = hw.OS; break;
                                case "RAM": val = hw.RAM; break;
                                case "CPU": val = hw.CPU; break;
                                case "Disk C:": val = hw.DiskC; break;
                                default: val = ""; break;
                            }
                            if (!string.IsNullOrEmpty(val)) hatTreffer = true;
                        }
                        dr[i + 2] = val;
                    }

                    if (chkNurVorhanden.Checked && !hatTreffer) continue;
                    dt.Rows.Add(dr);
                }

                // ── Grid befüllen ─────────────────────────────
                grid.DataSource = dt;

                // Spaltenbreiten
                if (grid.Columns.Count > 0) grid.Columns[0].Width = 180;
                if (grid.Columns.Count > 1) grid.Columns[1].Width = 110;
                for (int i = 2; i < grid.Columns.Count; i++)
                    grid.Columns[i].Width = spaltenTyp.Contains("Version") ? 120 : 50;

                lblStatus.Text = $"{dt.Rows.Count} Geräte · {columns.Count} Spalten";
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"Fehler: {ex.Message}";
            }
            finally
            {
                btnAuswerten.Enabled = true;
            }
        }

        // ── Zellen einfärben ──────────────────────────────────
        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex < 2 || e.Value == null) return;
            string val = e.Value.ToString();
            if (val == "✔")
            {
                e.CellStyle.ForeColor = Color.FromArgb(30, 130, 80);
                e.CellStyle.BackColor = Color.FromArgb(230, 248, 238);
                e.CellStyle.SelectionForeColor = Color.FromArgb(20, 100, 60);
            }
            else if (val == "–")
            {
                e.CellStyle.ForeColor = Color.FromArgb(190, 190, 190);
            }
            else if (!string.IsNullOrEmpty(val))
            {
                e.CellStyle.ForeColor = Color.FromArgb(30, 90, 160);
            }
        }

        // ── Hilfsmethoden ─────────────────────────────────────
        private string GetZeilenHeader(string gruppe)
        {
            switch (gruppe)
            {
                case "Kunde": return "Kunde";
                case "Gerätetyp": return "Gerätetyp";
                case "Betriebssystem": return "Betriebssystem";
                case "Gerät (IP)": return "IP";
                default: return "Hostname";
            }
        }

        private class PivotRow
        {
            public string Label { get; set; }
            public string IP { get; set; }
            public string Hostname { get; set; }
        }

        private List<PivotRow> BuildRows(List<DatabaseDevice> devices, string gruppe,
            Dictionary<string, string> kundenMap)
        {
            var rows = new List<PivotRow>();
            foreach (var d in devices)
            {
                string label;
                switch (gruppe)
                {
                    case "Gerät (IP)": label = d.IP; break;
                    case "Kunde": label = kundenMap.TryGetValue(d.IP, out var k) ? k : "Unbekannt"; break;
                    case "Gerätetyp": label = DeviceTypeHelper.GetLabel(d.DeviceType); break;
                    case "Betriebssystem": label = !string.IsNullOrEmpty(d.OS) ? d.OS : "Unbekannt"; break;
                    default: label = d.Hostname ?? d.IP; break;
                }
                rows.Add(new PivotRow { Label = label, IP = d.IP, Hostname = d.Hostname });
            }
            return rows;
        }

        private Dictionary<string, string> BuildKundenMap(List<Customer> kunden)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kunde in kunden)
            {
                var locs = _db.GetLocationsByCustomer(kunde.ID);
                foreach (var loc in locs)
                {
                    var ips = _db.GetIPsWithWorkstationByLocation(loc.ID);
                    foreach (var ip in ips)
                        if (!map.ContainsKey(ip.IPAddress))
                            map[ip.IPAddress] = kunde.Name;
                }
            }
            return map;
        }

        private class HardwareValues
        {
            public string OS { get; set; } = "";
            public string RAM { get; set; } = "";
            public string CPU { get; set; } = "";
            public string DiskC { get; set; } = "";
        }

        private Dictionary<string, HardwareValues> LoadHardwareValues(List<DatabaseDevice> devices)
        {
            var result = new Dictionary<string, HardwareValues>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in devices)
            {
                var hw = new HardwareValues { OS = d.OS ?? "" };

                // Hardware-Info aus DB laden
                try
                {
                    string hwText = _db.GetLatestHardwareInfo(d.IP);
                    if (!string.IsNullOrEmpty(hwText))
                    {
                        foreach (var line in hwText.Split('\n'))
                        {
                            if (line.Contains("RAM:") || line.Contains("Physischer RAM"))
                            {
                                var part = line.Split(':').LastOrDefault()?.Trim();
                                if (!string.IsNullOrEmpty(part)) hw.RAM = part;
                            }
                            else if (line.Contains("CPU:") || line.Contains("Prozessor"))
                            {
                                var part = line.Split(':').LastOrDefault()?.Trim();
                                if (!string.IsNullOrEmpty(part)) hw.CPU = part;
                            }
                            else if (line.Contains("Disk C:") || line.Contains("Laufwerk") && line.Contains("C:"))
                            {
                                var part = line.Split(':').LastOrDefault()?.Trim();
                                if (!string.IsNullOrEmpty(part)) hw.DiskC = part;
                            }
                        }
                    }
                }
                catch { }

                result[d.IP] = hw;
            }
            return result;
        }

        // ── LibreOffice Calc öffnen ───────────────────────────
        // Sucht LibreOffice automatisch an allen typischen Installationspfaden.
        // Falls nicht gefunden: Windows-Standard (was auch immer .xlsx öffnet).
        private static void OpenWithLibreOffice(string filePath)
        {
            string[] searchPaths = {
                // Standard-Installation 64-bit
                @"C:\Program Files\LibreOffice\program\scalc.exe",
                // Standard-Installation 32-bit
                @"C:\Program Files (x86)\LibreOffice\program\scalc.exe",
                // Ältere Versionen
                @"C:\Program Files\LibreOffice 7\program\scalc.exe",
                @"C:\Program Files\LibreOffice 6\program\scalc.exe",
                // Portable / benutzerdefiniert
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    @"Programs\LibreOffice\program\scalc.exe"),
            };

            string sCalc = null;
            foreach (var p in searchPaths)
                if (File.Exists(p)) { sCalc = p; break; }

            if (sCalc != null)
            {
                // LibreOffice Calc direkt mit der Datei starten
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sCalc,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false
                });
            }
            else
            {
                // Fallback: Windows-Standard-Programm für .xlsx
                System.Diagnostics.Process.Start(filePath);

                MessageBox.Show(
                    "LibreOffice Calc wurde nicht gefunden.\n" +
                    "Die Datei wird mit dem Standard-Programm geöffnet.\n\n" +
                    "LibreOffice herunterladen: https://www.libreoffice.org",
                    "LibreOffice nicht gefunden",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ── Export: CSV / Calc / PDF ──────────────────────────
        private void ExportCsv()
        {
            if (grid.DataSource == null)
            {
                MessageBox.Show("Erst Auswertung starten!", "Hinweis",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dlg = new SaveFileDialog
            {
                Title = "Auswertung exportieren",
                Filter = "Excel-Arbeitsmappe (*.xlsx)|*.xlsx|PDF-Dokument (*.pdf)|*.pdf|CSV-Datei (*.csv)|*.csv",
                FileName = $"Auswertung_{DateTime.Now:yyyyMMdd_HHmm}"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string ext = Path.GetExtension(dlg.FileName).ToLower();
                    var dt = (DataTable)grid.DataSource;

                    if (ext == ".xlsx") ExportExcel(dlg.FileName, dt);
                    else if (ext == ".pdf") ExportPdf(dlg.FileName, dt);
                    else ExportCsvFile(dlg.FileName, dt);

                    lblStatus.Text = $"✅ Exportiert: {Path.GetFileName(dlg.FileName)}";

                    // LibreOffice Calc direkt öffnen
                    if (ext == ".xlsx")
                    {
                        OpenWithLibreOffice(dlg.FileName);
                        lblStatus.Text += "  — LibreOffice Calc wird geöffnet...";
                    }
                    else if (MessageBox.Show("Datei jetzt öffnen?", "Export",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        System.Diagnostics.Process.Start(dlg.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export fehlgeschlagen:\n{ex.Message}", "Fehler",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ── Alle DB-Daten direkt in LibreOffice Calc öffnen ──
        // Exportiert alle Tabellen als separate Sheets und öffnet Calc.
        // Kein Filter nötig — komplette Datenbank als OpenXML-Arbeitsmappe.

        // ── Direkte SQLite-Lesemethoden für externe DB-Dateien ─
        private static List<DatabaseDevice> ReadDevicesFromDb(string dbPath)
        {
            var list = new List<DatabaseDevice>();
            using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                using (var cmd = new System.Data.SQLite.SQLiteCommand(
                    "SELECT ID,IP,Hostname,MacAddress,LastSeen,'' AS Status,'' AS Ports,'' AS OS FROM Devices ORDER BY LastSeen DESC", conn))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new DatabaseDevice
                        {
                            ID = r.GetInt32(0),
                            IP = r[1]?.ToString() ?? "",
                            Hostname = r[2]?.ToString() ?? "",
                            MacAddress = r[3]?.ToString() ?? "",
                            Zeitstempel = r[4]?.ToString() ?? "",
                            OS = r[7]?.ToString() ?? ""
                        });
            }
            return list;
        }

        private static List<DatabaseSoftware> ReadSoftwareFromDb(string dbPath)
        {
            var list = new List<DatabaseSoftware>();
            using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                try
                {
                    using (var cmd = new System.Data.SQLite.SQLiteCommand(@"
                        SELECT ds.ID, d.Hostname, ds.Name, ds.Version, ds.Publisher, ds.InstallDate, ds.QueryTime
                        FROM DeviceSoftware ds JOIN Devices d ON ds.DeviceID=d.ID
                        ORDER BY ds.QueryTime DESC", conn))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            list.Add(new DatabaseSoftware
                            {
                                ID = r.GetInt32(0),
                                PCName = r[1]?.ToString() ?? "",
                                Name = r[2]?.ToString() ?? "",
                                Version = r[3]?.ToString() ?? "",
                                Publisher = r[4]?.ToString() ?? "",
                                InstallDate = r[5]?.ToString() ?? "",
                                LastUpdate = r[6]?.ToString() ?? ""
                            });
                }
                catch { }
            }
            return list;
        }

        private static List<Customer> ReadCustomersFromDb(string dbPath)
        {
            var list = new List<Customer>();
            using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                conn.Open();
                try
                {
                    using (var cmd = new System.Data.SQLite.SQLiteCommand(
                        "SELECT ID,Name,Address FROM Customers ORDER BY Name", conn))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            list.Add(new Customer
                            {
                                ID = r.GetInt32(0),
                                Name = r[1]?.ToString() ?? "",
                                Address = r[2]?.ToString() ?? ""
                            });
                }
                catch { }
            }
            return list;
        }

        private void ExportAllDataToExcel()
        {
            lblStatus.Text = "Lade Daten aus ausgewählten Datenbanken...";
            Application.DoEvents();

            try
            {
                string path = Path.Combine(Path.GetTempPath(),
                    $"NmapInventory_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

                var wb = new XLWorkbook();

                // ── Daten aus allen ausgewählten DBs sammeln ──
                var allDevices = new List<DatabaseDevice>();
                var allSoftware = new List<DatabaseSoftware>();
                var allKunden = new List<Customer>();

                var paths = (_dbPaths != null && _dbPaths.Count > 0)
                    ? _dbPaths
                    : new List<string> { Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory, "nmap_inventory.db") };

                foreach (var dbPath in paths.Where(p => !string.IsNullOrEmpty(p) && File.Exists(p)))
                {
                    try
                    {
                        // Direkt per SQLite lesen — funktioniert mit jeder .db ohne eigenen DatabaseManager
                        allDevices.AddRange(ReadDevicesFromDb(dbPath));
                        allSoftware.AddRange(ReadSoftwareFromDb(dbPath));
                        foreach (var k in ReadCustomersFromDb(dbPath))
                            if (!allKunden.Any(e => e.Name == k.Name))
                                allKunden.Add(k);
                    }
                    catch { }
                }

                var devices = allDevices;
                var software = allSoftware;
                var kunden = allKunden.Count > 0 ? allKunden : _db.GetCustomers();

                // ── Sheet 1: Geräte ───────────────────────────
                var wsGeraete = wb.Worksheets.Add("Geräte");
                WriteHeader(wsGeraete, new[] {
                    "ID","IP","Hostname","MAC-Adresse","Betriebssystem",
                    "Status","Ports","Zuletzt gesehen"
                });
                for (int r = 0; r < devices.Count; r++)
                {
                    var d = devices[r];
                    wsGeraete.Cell(r + 2, 1).Value = d.ID;
                    wsGeraete.Cell(r + 2, 2).Value = d.IP ?? "";
                    wsGeraete.Cell(r + 2, 3).Value = d.Hostname ?? "";
                    wsGeraete.Cell(r + 2, 4).Value = d.MacAddress ?? "";
                    wsGeraete.Cell(r + 2, 5).Value = d.OS ?? "";
                    wsGeraete.Cell(r + 2, 6).Value = d.Status ?? "";
                    wsGeraete.Cell(r + 2, 7).Value = d.Ports ?? "";
                    wsGeraete.Cell(r + 2, 8).Value = d.Zeitstempel ?? "";
                }
                FormatSheet(wsGeraete, devices.Count);

                // ── Sheet 2: Software ─────────────────────────
                var wsSoftware = wb.Worksheets.Add("Software");
                WriteHeader(wsSoftware, new[] {
                    "ID","PC/Gerät","Software","Version",
                    "Hersteller","Installiert am","Letzte Änderung"
                });
                for (int r = 0; r < software.Count; r++)
                {
                    var s = software[r];
                    wsSoftware.Cell(r + 2, 1).Value = s.ID;
                    wsSoftware.Cell(r + 2, 2).Value = s.PCName ?? "";
                    wsSoftware.Cell(r + 2, 3).Value = s.Name ?? "";
                    wsSoftware.Cell(r + 2, 4).Value = s.Version ?? "";
                    wsSoftware.Cell(r + 2, 5).Value = s.Publisher ?? "";
                    wsSoftware.Cell(r + 2, 6).Value = s.InstallDate ?? "";
                    wsSoftware.Cell(r + 2, 7).Value = s.LastUpdate ?? "";
                }
                FormatSheet(wsSoftware, software.Count);

                // ── Sheet 3: Kunden ───────────────────────────
                var wsKunden = wb.Worksheets.Add("Kunden");
                WriteHeader(wsKunden, new[] { "ID", "Name", "Adresse" });
                for (int r = 0; r < kunden.Count; r++)
                {
                    wsKunden.Cell(r + 2, 1).Value = kunden[r].ID;
                    wsKunden.Cell(r + 2, 2).Value = kunden[r].Name ?? "";
                    wsKunden.Cell(r + 2, 3).Value = kunden[r].Address ?? "";
                }
                FormatSheet(wsKunden, kunden.Count);

                // ── Sheet 4: Standorte ────────────────────────
                var wsStandorte = wb.Worksheets.Add("Standorte");
                WriteHeader(wsStandorte, new[] {
                    "ID","Kundenname","Standort","Ebene","Adresse"
                });
                int row = 2;
                foreach (var k in kunden)
                {
                    var locs = _db.GetLocationsByCustomer(k.ID);
                    foreach (var loc in locs)
                        WriteLocationRecursive(wsStandorte, loc, k.Name, ref row, _db);
                }
                FormatSheet(wsStandorte, row - 2);

                // ── Sheet 5: Geräte×Software (Pivot-Basis) ────
                var wsPivot = wb.Worksheets.Add("Pivot-Basis");
                WriteHeader(wsPivot, new[] {
                    "Hostname","IP","Betriebssystem","Software","Version","Hersteller"
                });
                int pr = 2;
                var swByDevice = software.GroupBy(s => s.PCName ?? "")
                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
                foreach (var dev in devices)
                {
                    swByDevice.TryGetValue(dev.Hostname ?? "", out var devSw);
                    if (devSw == null) swByDevice.TryGetValue(dev.IP ?? "", out devSw);
                    if (devSw == null || devSw.Count == 0)
                    {
                        wsPivot.Cell(pr, 1).Value = dev.Hostname ?? "";
                        wsPivot.Cell(pr, 2).Value = dev.IP ?? "";
                        wsPivot.Cell(pr, 3).Value = dev.OS ?? "";
                        pr++;
                    }
                    else
                    {
                        foreach (var sw in devSw)
                        {
                            wsPivot.Cell(pr, 1).Value = dev.Hostname ?? "";
                            wsPivot.Cell(pr, 2).Value = dev.IP ?? "";
                            wsPivot.Cell(pr, 3).Value = dev.OS ?? "";
                            wsPivot.Cell(pr, 4).Value = sw.Name ?? "";
                            wsPivot.Cell(pr, 5).Value = sw.Version ?? "";
                            wsPivot.Cell(pr, 6).Value = sw.Publisher ?? "";
                            pr++;
                        }
                    }
                }
                // Als Tabelle formatieren — Excel erkennt sie direkt als Pivot-Quelle
                if (pr > 2)
                {
                    var tbl = wsPivot.Range(1, 1, pr - 1, 6).CreateTable("GeraeteSoftware");
                    tbl.Theme = XLTableTheme.TableStyleMedium2;
                }
                FormatSheet(wsPivot, pr - 2);

                // ── Sheet 6: Anleitung ────────────────────────
                var wsInfo = wb.Worksheets.Add("Pivot-Anleitung");
                wsInfo.Cell("A1").Value = "Pivot-Tabelle erstellen — Schritt für Schritt";
                wsInfo.Cell("A1").Style.Font.Bold = true;
                wsInfo.Cell("A1").Style.Font.FontSize = 13;
                string[] steps = {
                    "", "1. Wechseln Sie zu 'Pivot-Basis'",
                    "2. Klicken Sie in die Tabelle 'GeraeteSoftware'",
                    "3. Menü: Einfügen → PivotTable → OK",
                    "4. Felder zuordnen:",
                    "   Zeilen:   Hostname",
                    "   Spalten:  Software",
                    "   Werte:    Anzahl von Software",
                    "   Filter:   Betriebssystem oder Hersteller",
                    "",
                    "Weitere Auswertungs-Ideen:",
                    "   • Welche Software ist auf wie vielen Geräten?  → Software in Zeilen, Anzahl Hostname als Wert",
                    "   • Welche Geräte haben noch Windows 10?         → Betriebssystem filtern",
                    "   • Software-Vergleich je Kunde:                 → Kundenname als Filter hinzufügen"
                };
                for (int i = 0; i < steps.Length; i++)
                {
                    wsInfo.Cell(i + 3, 1).Value = steps[i];
                    if (steps[i].StartsWith("   "))
                        wsInfo.Cell(i + 3, 1).Style.Font.FontColor = XLColor.FromArgb(80, 80, 80);
                    else if (steps[i].StartsWith("1.") || steps[i].StartsWith("2.") ||
                             steps[i].StartsWith("3.") || steps[i].StartsWith("4."))
                        wsInfo.Cell(i + 3, 1).Style.Font.Bold = true;
                }
                wsInfo.Column(1).Width = 75;

                // Pivot-Basis als aktives Sheet
                wsPivot.SetTabActive();

                wb.SaveAs(path);
                lblStatus.Text = $"✅ {devices.Count} Geräte, {software.Count} Software-Einträge exportiert";

                // LibreOffice Calc direkt öffnen
                OpenWithLibreOffice(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler:\n{ex.Message}", "Export fehlgeschlagen",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Fehler beim Export";
            }
        }

        private void WriteHeader(IXLWorksheet ws, string[] headers)
        {
            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(34, 84, 150);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            }
        }

        private void FormatSheet(IXLWorksheet ws, int rowCount)
        {
            ws.Columns().AdjustToContents(1, 50);
            ws.SheetView.FreezeRows(1);
            if (rowCount > 0)
            {
                var range = ws.Range(1, 1, rowCount + 1, ws.LastColumnUsed().ColumnNumber());
                for (int r = 2; r <= rowCount + 1; r++)
                    if (r % 2 == 0)
                        ws.Range(r, 1, r, range.ColumnCount()).Style
                          .Fill.BackgroundColor = XLColor.FromArgb(245, 248, 255);
            }
        }

        private static void WriteLocationRecursive(
            IXLWorksheet ws, Location loc, string kundenName, ref int row, DatabaseManager db)
        {
            ws.Cell(row, 1).Value = loc.ID;
            ws.Cell(row, 2).Value = kundenName;
            ws.Cell(row, 3).Value = loc.Name ?? "";
            ws.Cell(row, 4).Value = loc.Level;
            ws.Cell(row, 5).Value = loc.Address ?? "";
            row++;
            foreach (var child in db.GetChildLocations(loc.ID))
                WriteLocationRecursive(ws, child, kundenName, ref row, db);
        }

        // ── XLSX via ClosedXML — pivot-fähig ─────────────────
        // NuGet: Install-Package ClosedXML
        // Erstellt zwei Sheets: "Daten" (pivot-fähige Rohdaten) + "Auswertung" (formatierte Ansicht)
        // LibreOffice Calc öffnet danach automatisch
        private void ExportExcel(string path, DataTable dt)
        {
            var wb = new XLWorkbook();

            // ── Sheet 1: Rohdaten (pivot-optimiert) ───────────
            // Flache Tabelle — jede Zeile = ein Gerät+Software-Kombination
            // Excel kann daraus sofort eine Pivot-Tabelle erstellen
            var wsDaten = wb.Worksheets.Add("Daten");

            // Header
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var cell = wsDaten.Cell(1, c + 1);
                cell.Value = dt.Columns[c].ColumnName;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(34, 84, 150);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // Daten
            for (int r = 0; r < dt.Rows.Count; r++)
                for (int c = 0; c < dt.Columns.Count; c++)
                    wsDaten.Cell(r + 2, c + 1).Value = dt.Rows[r][c]?.ToString() ?? "";

            // Als Excel-Tabelle (ListObject) formatieren → Pivot kann direkt darauf zugreifen
            if (dt.Rows.Count > 0)
            {
                var tbl = wsDaten.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count)
                                 .CreateTable("Geraetedaten");
                tbl.Theme = XLTableTheme.TableStyleMedium2;
            }

            wsDaten.Columns().AdjustToContents(1, 40);
            wsDaten.SheetView.FreezeRows(1);

            // ── Sheet 2: Formatierte Auswertung ───────────────
            var ws = wb.Worksheets.Add("Auswertung");

            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = dt.Columns[c].ColumnName;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(220, 230, 245);
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            }

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var cell = ws.Cell(r + 2, c + 1);
                    string val = dt.Rows[r][c]?.ToString() ?? "";
                    cell.Value = val;
                    if (val == "✔")
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromArgb(220, 245, 220);
                        cell.Style.Font.FontColor = XLColor.FromArgb(30, 130, 60);
                    }
                    else if (val == "–")
                        cell.Style.Font.FontColor = XLColor.FromArgb(180, 180, 180);
                    else if (r % 2 == 1)
                        cell.Style.Fill.BackgroundColor = XLColor.FromArgb(248, 248, 252);
                }
            }

            ws.Columns().AdjustToContents(1, 50);
            ws.SheetView.FreezeRows(1);

            // ── Sheet 3: Anleitung Pivot ──────────────────────
            var wsInfo = wb.Worksheets.Add("Pivot-Anleitung");
            wsInfo.Cell("A1").Value = "So erstellen Sie eine Pivot-Tabelle:";
            wsInfo.Cell("A1").Style.Font.Bold = true;
            wsInfo.Cell("A1").Style.Font.FontSize = 14;
            string[] steps = {
                "1. Wechseln Sie zum Blatt 'Daten'",
                "2. Klicken Sie in die Tabelle 'Geraetedaten'",
                "3. Menü: Einfügen → PivotTable",
                "4. 'Vorhandenes Arbeitsblatt' wählen → OK",
                "5. Ziehen Sie Felder in Zeilen / Spalten / Werte:",
                "   • Zeilen:  Hostname oder Kunde",
                "   • Spalten: Software-Name",
                "   • Werte:   Anzahl von Software (zeigt wer was hat)",
                "   • Filter:  Gerätetyp oder Betriebssystem",
                "",
                "Tipp: Mit dem Filter 'Werte anzeigen als → % der Gesamtsumme'",
                "      sehen Sie den Verbreitungsgrad jeder Software."
            };
            for (int i = 0; i < steps.Length; i++)
            {
                wsInfo.Cell(i + 3, 1).Value = steps[i];
                if (steps[i].StartsWith("  "))
                    wsInfo.Cell(i + 3, 1).Style.Font.FontColor = XLColor.FromArgb(80, 80, 80);
            }
            wsInfo.Column(1).Width = 65;

            // Daten-Sheet als aktives Sheet setzen
            wsDaten.SetTabActive();

            wb.SaveAs(path);
        }

        // ── PDF via iTextSharp ────────────────────────────────
        // NuGet: Install-Package iTextSharp
        private void ExportPdf(string path, DataTable dt)
        {
            var doc = new itextText.Document(
                itextText.PageSize.A4.Rotate(), 20, 20, 30, 30);
            var writer = itextPdf.PdfWriter.GetInstance(
                doc, new FileStream(path, FileMode.Create));
            doc.Open();

            // Titel
            var titleFont = itextText.FontFactory.GetFont(
                itextText.FontFactory.HELVETICA_BOLD, 14f);
            doc.Add(new itextText.Paragraph(
                $"Auswertung — {DateTime.Now:dd.MM.yyyy HH:mm}", titleFont));
            doc.Add(new itextText.Paragraph(
                $"Zeilen: {dt.Rows.Count}  ·  Spalten: {dt.Columns.Count}",
                itextText.FontFactory.GetFont(itextText.FontFactory.HELVETICA, 9f)));
            doc.Add(itextText.Chunk.NEWLINE);

            // Tabelle
            var table = new itextPdf.PdfPTable(dt.Columns.Count)
            {
                WidthPercentage = 100,
                SpacingBefore = 10,
                HorizontalAlignment = itextText.Element.ALIGN_LEFT
            };

            // Header
            var hdrFont = itextText.FontFactory.GetFont(
                itextText.FontFactory.HELVETICA_BOLD, 8f,
                itextText.BaseColor.WHITE);
            var hdrColor = new itextText.BaseColor(50, 100, 160);

            foreach (DataColumn col in dt.Columns)
            {
                var cell = new itextPdf.PdfPCell(
                    new itextText.Phrase(col.ColumnName, hdrFont))
                {
                    BackgroundColor = hdrColor,
                    Padding = 4,
                    HorizontalAlignment = itextText.Element.ALIGN_CENTER
                };
                table.AddCell(cell);
            }

            // Daten
            var dataFont = itextText.FontFactory.GetFont(
                itextText.FontFactory.HELVETICA, 7f);
            var greenFont = itextText.FontFactory.GetFont(
                itextText.FontFactory.HELVETICA_BOLD, 7f,
                new itextText.BaseColor(30, 130, 60));
            var grayFont = itextText.FontFactory.GetFont(
                itextText.FontFactory.HELVETICA, 7f,
                new itextText.BaseColor(180, 180, 180));
            var altBg = new itextText.BaseColor(248, 248, 252);

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    string val = dt.Rows[r][c]?.ToString() ?? "";
                    var font = val == "✔" ? greenFont : (val == "–" ? grayFont : dataFont);
                    var cell = new itextPdf.PdfPCell(
                        new itextText.Phrase(val, font))
                    {
                        Padding = 3,
                        HorizontalAlignment = itextText.Element.ALIGN_CENTER
                    };
                    if (r % 2 == 1)
                        cell.BackgroundColor = altBg;
                    table.AddCell(cell);
                }
            }

            doc.Add(table);
            doc.Close();
        }

        // ── CSV (nativ, kein NuGet) ───────────────────────────
        private void ExportCsvFile(string path, DataTable dt)
        {
            var sb = new StringBuilder();
            // BOM für Excel-Kompatibilität
            sb.Append('\uFEFF');
            sb.AppendLine(string.Join(";", dt.Columns.Cast<DataColumn>()
                .Select(c => $"\"{c.ColumnName}\"")));
            foreach (DataRow row in dt.Rows)
                sb.AppendLine(string.Join(";",
                    row.ItemArray.Select(v => $"\"{v?.ToString()?.Replace("\"", "\"\"")}\"")));
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }
    }
}