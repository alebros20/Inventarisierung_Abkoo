using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public class DatabaseSelectionDialog : Form
    {
        private readonly DatabaseManager _db;
        private CheckedListBox clb;
        public List<string> SelectedDatabasePaths { get; private set; } = new List<string>();

        public DatabaseSelectionDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Datenbanken auswählen";
            Size = new Size(520, 420);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 10);

            var header = new Label
            {
                Text = "Welche Datenbanken sollen in LibreOffice Calc geöffnet werden?",
                Location = new Point(12, 12),
                Size = new Size(480, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            clb = new CheckedListBox
            {
                Location = new Point(12, 58),
                Size = new Size(480, 210),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 10)
            };

            LoadDatabases();

            var btnAlle = new Button { Text = "✔ Alle auswählen", Location = new Point(12, 276), Width = 155, Height = 28 };
            btnAlle.Click += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true); };

            var btnKeine = new Button { Text = "✖ Keine", Location = new Point(175, 276), Width = 100, Height = 28 };
            btnKeine.Click += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false); };

            var lblAnzahl = new Label
            {
                Location = new Point(12, 312),
                Size = new Size(300, 20),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray,
                Text = $"0 von {clb.Items.Count} Datenbanken ausgewählt"
            };
            clb.ItemCheck += (s, e) =>
            {
                int cnt = clb.CheckedItems.Count + (e.NewValue == CheckState.Checked ? 1 : -1);
                lblAnzahl.Text = $"{cnt} von {clb.Items.Count} Datenbanken ausgewählt";
            };

            var btnOk = new Button
            {
                Text = "In Calc öffnen",
                Location = new Point(290, 340),
                Width = 140,
                Height = 30,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) =>
            {
                SelectedDatabasePaths = clb.CheckedItems.Cast<DbItem>()
                    .Select(item => item.Path).ToList();
                if (SelectedDatabasePaths.Count == 0)
                {
                    MessageBox.Show("Bitte mindestens eine Datenbank auswählen.",
                        "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                }
            };

            var btnAbbrechen = new Button
            {
                Text = "Abbrechen",
                Location = new Point(440, 340),
                Width = 68,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[] { header, clb, btnAlle, btnKeine, lblAnzahl, btnOk, btnAbbrechen });
            AcceptButton = btnOk;
            CancelButton = btnAbbrechen;
        }

        private void LoadDatabases()
        {
            clb.Items.Clear();
            try
            {
                string mainPath = _db.GetMainDatabasePath();
                var kunden      = _db.GetCustomers();
                var allFiles    = _db.GetAllDatabaseFiles();

                // Haupt-DB zuerst
                if (System.IO.File.Exists(mainPath))
                {
                    long kb = new System.IO.FileInfo(mainPath).Length / 1024;
                    clb.Items.Add(new DbItem { Path = mainPath, Display = $"Haupt-Datenbank  ({kb} KB)  —  nmap_inventory.db" }, true);
                }

                // Kunden-DBs
                foreach (var file in allFiles.OrderBy(f => f))
                {
                    if (file == mainPath) continue;
                    var custId = _db.TryGetCustomerIdFromPath(file);
                    string name = custId.HasValue
                        ? kunden.FirstOrDefault(k => k.ID == custId.Value)?.Name ?? $"Kunde {custId}"
                        : System.IO.Path.GetFileName(file);
                    long kb = new System.IO.FileInfo(file).Length / 1024;
                    clb.Items.Add(new DbItem { Path = file, Display = $"{name,-28}  ({kb} KB)  —  {System.IO.Path.GetFileName(file)}" }, true);
                }
            }
            catch { }

            if (clb.Items.Count == 0)
                clb.Items.Add(new DbItem { Path = "", Display = "Keine Datenbanken gefunden" }, false);
        }

        private class DbItem
        {
            public string Path { get; set; }
            public string Display { get; set; }
            public override string ToString() => Display;
        }
    }
}
