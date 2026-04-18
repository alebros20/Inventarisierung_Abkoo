using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    /// <summary>
    /// Dialog zum Exportieren einzelner Kunden-Datenbanken.
    /// Der Benutzer wählt einen oder mehrere Kunden per Checkbox und ein Zielverzeichnis.
    /// </summary>
    public class CustomerDbExportDialog : Form
    {
        private readonly DatabaseManager _db;
        private CheckedListBox _list;
        private Button _btnAll, _btnNone, _btnExport, _btnCancel;

        public int ExportedCount { get; private set; }

        public CustomerDbExportDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Kunden-Datenbanken exportieren";
            Width = 460; Height = 480;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 10);

            var lbl = new Label
            {
                Text = "Kunden auswählen, deren Datenbank exportiert werden soll:",
                Location = new Point(12, 12), AutoSize = true
            };

            _list = new CheckedListBox
            {
                Location = new Point(12, 40), Size = new Size(420, 320),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 10)
            };
            PopulateList();

            _btnAll = new Button { Text = "Alle", Location = new Point(12, 370), Width = 100, Height = 28 };
            _btnAll.Click += (s, e) => { for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, true); };

            _btnNone = new Button { Text = "Keine", Location = new Point(120, 370), Width = 100, Height = 28 };
            _btnNone.Click += (s, e) => { for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, false); };

            _btnExport = new Button { Text = "Exportieren", Location = new Point(222, 370), Width = 110, Height = 28, DialogResult = DialogResult.None };
            _btnExport.Click += (s, e) => DoExport();

            _btnCancel = new Button { Text = "Abbrechen", Location = new Point(340, 370), Width = 92, Height = 28, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { lbl, _list, _btnAll, _btnNone, _btnExport, _btnCancel });
            AcceptButton = _btnExport;
            CancelButton = _btnCancel;
        }

        private void PopulateList()
        {
            _list.Items.Clear();
            foreach (var c in _db.GetCustomers())
            {
                string path = _db.GetCustomerDatabasePath(c.ID);
                string label = File.Exists(path)
                    ? $"{c.Name}   (ID {c.ID})"
                    : $"{c.Name}   (ID {c.ID})  — keine DB-Datei";
                _list.Items.Add(new Entry { CustomerId = c.ID, Name = c.Name, Path = path, Exists = File.Exists(path), Label = label });
            }
            _list.DisplayMember = "Label";
        }

        private void DoExport()
        {
            var selected = _list.CheckedItems.Cast<Entry>().Where(e => e.Exists).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("Bitte mindestens einen Kunden mit vorhandener DB-Datei auswählen.",
                    "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var fbd = new FolderBrowserDialog { Description = "Zielverzeichnis für den Export wählen:" })
            {
                if (fbd.ShowDialog(this) != DialogResult.OK) return;
                string targetDir = fbd.SelectedPath;

                try
                {
                    int count = 0;
                    foreach (var entry in selected)
                    {
                        string safeName = string.Concat(entry.Name.Split(Path.GetInvalidFileNameChars()));
                        string destFile = Path.Combine(targetDir,
                            $"nmap_customer_{entry.CustomerId}_{safeName}_{DateTime.Now:yyyy-MM-dd}.db");
                        File.Copy(entry.Path, destFile, true);
                        count++;
                    }
                    ExportedCount = count;
                    MessageBox.Show($"{count} Kunden-Datenbank(en) exportiert nach:\n{targetDir}",
                        "Export erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.OK;
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Export:\n" + ex.Message, "Fehler",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private class Entry
        {
            public int CustomerId { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public bool Exists { get; set; }
            public string Label { get; set; }
            public override string ToString() => Label;
        }
    }
}
