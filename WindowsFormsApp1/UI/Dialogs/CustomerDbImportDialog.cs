using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace NmapInventory
{
    /// <summary>
    /// Dialog zum Importieren einer Kunden-Datenbank in einen Ziel-Kunden.
    /// Modi: Zusammenführen (Merge), Erweitern (nur Neue), Überschreiben.
    /// </summary>
    public class CustomerDbImportDialog : Form
    {
        private readonly DatabaseManager _db;

        private TextBox _tbSourcePath;
        private Button _btnBrowse;
        private ComboBox _cbTargetCustomer;
        private RadioButton _rbMerge, _rbExtend, _rbOverwrite;
        private Button _btnImport, _btnCancel;

        public string SourcePath => _tbSourcePath.Text.Trim();
        public int TargetCustomerId => (_cbTargetCustomer.SelectedItem is ComboItem it) ? it.ID : -1;
        public DatabaseManager.CustomerImportMode Mode =>
            _rbOverwrite.Checked ? DatabaseManager.CustomerImportMode.Overwrite
            : _rbExtend.Checked   ? DatabaseManager.CustomerImportMode.Extend
            :                       DatabaseManager.CustomerImportMode.Merge;

        public DatabaseManager.CustomerImportResult ImportResult { get; private set; }

        public CustomerDbImportDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Kunden-Datenbank importieren";
            Width = 520; Height = 380;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 10);

            // ── Quelle ──
            Controls.Add(new Label { Text = "Quell-DB-Datei:", Location = new Point(14, 16), AutoSize = true });
            _tbSourcePath = new TextBox { Location = new Point(14, 38), Width = 380, ReadOnly = true };
            _btnBrowse = new Button { Text = "...", Location = new Point(400, 37), Width = 90, Height = 26 };
            _btnBrowse.Click += (s, e) => BrowseSource();
            Controls.Add(_tbSourcePath);
            Controls.Add(_btnBrowse);

            // ── Ziel-Kunde ──
            Controls.Add(new Label { Text = "Ziel-Kunde:", Location = new Point(14, 76), AutoSize = true });
            _cbTargetCustomer = new ComboBox { Location = new Point(14, 98), Width = 476, DropDownStyle = ComboBoxStyle.DropDownList };
            PopulateCustomers();
            Controls.Add(_cbTargetCustomer);

            // ── Modus ──
            var grp = new GroupBox { Text = "Import-Modus", Location = new Point(14, 134), Size = new Size(476, 140) };
            _rbMerge = new RadioButton
            {
                Text = "Zusammenführen — bestehende Einträge aktualisieren, neue hinzufügen",
                Location = new Point(12, 22), AutoSize = true, Checked = true
            };
            _rbExtend = new RadioButton
            {
                Text = "Erweitern — nur neue Einträge hinzufügen, bestehende unverändert lassen",
                Location = new Point(12, 54), AutoSize = true
            };
            _rbOverwrite = new RadioButton
            {
                Text = "Überschreiben — Ziel-Kunden-DB vollständig durch Quelle ersetzen",
                Location = new Point(12, 86), AutoSize = true
            };
            var warnLbl = new Label
            {
                Text = "⚠ Bei 'Überschreiben' gehen alle bisherigen Daten des Ziel-Kunden verloren.",
                Location = new Point(12, 112), AutoSize = true,
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Firebrick
            };
            grp.Controls.AddRange(new Control[] { _rbMerge, _rbExtend, _rbOverwrite, warnLbl });
            Controls.Add(grp);

            // ── Buttons ──
            _btnImport = new Button { Text = "Importieren", Location = new Point(276, 292), Width = 110, Height = 28 };
            _btnImport.Click += (s, e) => DoImport();
            _btnCancel = new Button { Text = "Abbrechen", Location = new Point(396, 292), Width = 94, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.Add(_btnImport);
            Controls.Add(_btnCancel);
            AcceptButton = _btnImport;
            CancelButton = _btnCancel;
        }

        private void PopulateCustomers()
        {
            _cbTargetCustomer.Items.Clear();
            foreach (var c in _db.GetCustomers())
                _cbTargetCustomer.Items.Add(new ComboItem { ID = c.ID, Display = $"{c.Name} (ID {c.ID})" });
            if (_cbTargetCustomer.Items.Count > 0) _cbTargetCustomer.SelectedIndex = 0;
        }

        private void BrowseSource()
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "Kunden-Datenbank auswählen",
                Filter = "SQLite Datenbank (*.db)|*.db|Alle Dateien|*.*"
            })
            {
                if (ofd.ShowDialog(this) == DialogResult.OK)
                    _tbSourcePath.Text = ofd.FileName;
            }
        }

        private void DoImport()
        {
            if (string.IsNullOrWhiteSpace(SourcePath) || !File.Exists(SourcePath))
            {
                MessageBox.Show("Bitte eine gültige Quell-DB-Datei auswählen.", "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (TargetCustomerId <= 0)
            {
                MessageBox.Show("Bitte einen Ziel-Kunden auswählen.", "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (Mode == DatabaseManager.CustomerImportMode.Overwrite)
            {
                var confirm = MessageBox.Show(
                    "Beim Überschreiben werden ALLE bisherigen Daten des Ziel-Kunden gelöscht und durch die Quell-DB ersetzt.\n\nFortfahren?",
                    "Überschreiben bestätigen", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm != DialogResult.Yes) return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                ImportResult = _db.ImportCustomerDatabase(SourcePath, TargetCustomerId, Mode);

                MessageBox.Show(
                    "Import erfolgreich abgeschlossen:\n\n" + ImportResult.Summary,
                    "Import abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Import:\n" + ex.Message, "Fehler",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}
