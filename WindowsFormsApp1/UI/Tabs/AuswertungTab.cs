using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace NmapInventory
{
    public partial class MainForm
    {
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
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportDb.Click += (s, e) => ExportDatabase();

            var btnImportDb = new Button
            {
                Text = "📤 Datenbank importieren",
                Location = new Point(236, 8), Width = 220, Height = 34,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
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
                calcExportButton.Text = "Exportiere... (1/4) Geräte";
                Application.DoEvents();
                var dtGeraete = dbManager.GetViewData("Inventar_Geraete");

                // Sheet 2: Hardware (geparst in einzelne Spalten)
                calcExportButton.Text = "Exportiere... (2/4) Hardware";
                Application.DoEvents();
                var dtHardware = dbManager.GetHardwareExportData();

                // Sheet 3: Software
                calcExportButton.Text = "Exportiere... (3/4) Software";
                Application.DoEvents();
                var dtSoftware = dbManager.GetViewData("Inventar_Software");

                // Sheet 4: Ports
                calcExportButton.Text = "Exportiere... (4/4) Ports";
                Application.DoEvents();
                var dtPorts = dbManager.GetViewData("Inventar_Ports");

                if (dtGeraete.Rows.Count == 0 && dtSoftware.Rows.Count == 0 && dtPorts.Rows.Count == 0 && dtHardware.Rows.Count == 0)
                {
                    MessageBox.Show("Keine Daten vorhanden.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string filePath = Path.Combine(Path.GetTempPath(),
                    $"Inventarisierung_{DateTime.Now:yyyy-MM-dd_HHmmss}.xlsx");

                using (var wb = new XLWorkbook())
                {
                    AddSheet(wb, "Geräte", dtGeraete);
                    AddSheet(wb, "Hardware", dtHardware);
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
    }
}
