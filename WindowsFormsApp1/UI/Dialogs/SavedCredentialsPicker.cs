using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public class SavedCredentialsPicker : Form
    {
        // Für Einzelauswahl (Kompatibilität)
        public CredentialEntry Selected => SelectedEntries.Count > 0 ? SelectedEntries[0] : null;
        // Für Mehrfachauswahl
        public List<CredentialEntry> SelectedEntries { get; private set; } = new List<CredentialEntry>();

        private readonly DatabaseManager _db;
        private TreeView tree;
        private Label lblCount;

        public SavedCredentialsPicker(DatabaseManager db)
        {
            _db = db;
            Text = "Gespeicherte Zugangsdaten auswählen";
            Size = new Size(440, 530);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 9);

            var header = new Label
            {
                Text = "Geräte auswählen (Mehrfachauswahl möglich):",
                Location = new Point(12, 12),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            tree = new TreeView
            {
                Location = new Point(12, 34),
                Size = new Size(406, 380),
                FullRowSelect = true,
                ShowLines = true,
                ShowPlusMinus = true,
                CheckBoxes = true,
                Font = new Font("Segoe UI", 9)
            };

            LoadTree();

            // Kunden-Knoten anklicken → alle Kinder an/abwählen
            tree.AfterCheck += (s, e) =>
            {
                if (e.Action == TreeViewAction.Unknown) return;
                // Kunden-Gruppe: alle Kinder synchronisieren
                if (e.Node.Tag == null)
                {
                    foreach (TreeNode child in e.Node.Nodes)
                        child.Checked = e.Node.Checked;
                }
                // Gerät: Eltern-Status aktualisieren
                else if (e.Node.Parent != null)
                {
                    bool allChecked = true;
                    foreach (TreeNode sibling in e.Node.Parent.Nodes)
                    {
                        if (!sibling.Checked) allChecked = false;
                    }
                    e.Node.Parent.Checked = allChecked;
                }
                UpdateCount();
            };

            // Doppelklick = Einzelauswahl sofort bestätigen
            tree.NodeMouseDoubleClick += (s, e) =>
            {
                if (e.Node?.Tag is CredentialEntry)
                {
                    e.Node.Checked = true;
                    ConfirmSelection();
                }
            };

            lblCount = new Label
            {
                Text = "0 Geräte ausgewählt",
                Location = new Point(12, 420),
                Size = new Size(280, 18),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };

            // Alle / Keine Buttons
            var btnAll = new Button
            {
                Text = "✔ Alle",
                Location = new Point(290, 416),
                Width = 60,
                Height = 24,
                Font = new Font("Segoe UI", 8)
            };
            btnAll.Click += (s, e) =>
            {
                foreach (TreeNode root in tree.Nodes)
                    foreach (TreeNode child in root.Nodes)
                        child.Checked = true;
                UpdateCount();
            };

            var btnNone = new Button
            {
                Text = "✖ Keine",
                Location = new Point(356, 416),
                Width = 62,
                Height = 24,
                Font = new Font("Segoe UI", 8)
            };
            btnNone.Click += (s, e) =>
            {
                foreach (TreeNode root in tree.Nodes)
                    foreach (TreeNode child in root.Nodes)
                        child.Checked = false;
                UpdateCount();
            };

            var btnOk = new Button
            {
                Text = "Abfragen",
                Location = new Point(12, 448),
                Width = 130,
                Height = 30,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) => ConfirmSelection();

            var btnCancel = new Button
            {
                Text = "Abbrechen",
                Location = new Point(152, 448),
                Width = 110,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[] { header, tree, lblCount, btnAll, btnNone, btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void UpdateCount()
        {
            int count = 0;
            foreach (TreeNode root in tree.Nodes)
                foreach (TreeNode child in root.Nodes)
                    if (child.Checked) count++;
            lblCount.Text = count == 0 ? "Kein Gerät ausgewählt"
                : count == 1 ? "1 Gerät ausgewählt"
                : $"{count} Geräte ausgewählt";
        }

        private void ConfirmSelection()
        {
            SelectedEntries = new List<CredentialEntry>();
            foreach (TreeNode root in tree.Nodes)
                foreach (TreeNode child in root.Nodes)
                    if (child.Checked && child.Tag is CredentialEntry entry)
                        SelectedEntries.Add(entry);

            if (SelectedEntries.Count == 0)
            {
                MessageBox.Show("Bitte mindestens ein Gerät auswählen.", "Hinweis",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void LoadTree()
        {
            tree.Nodes.Clear();
            var entries = CredentialStore.GetAllRaw();
            if (entries.Count == 0)
            {
                tree.Nodes.Add("(Keine gespeicherten Einträge)");
                return;
            }

            // Kunden aus DB laden für Gruppierung
            var kundenMap = new Dictionary<string, TreeNode>();

            if (_db != null)
            {
                try
                {
                    foreach (var k in _db.GetCustomers())
                    {
                        // IPs dieses Kunden sammeln
                        var kundeIPs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var loc in _db.GetLocationsByCustomer(k.ID))
                            foreach (var lip in _db.GetIPsWithWorkstationByLocation(loc.ID))
                                kundeIPs.Add(lip.IPAddress);

                        // Gespeicherte Einträge die zu diesem Kunden gehören
                        var kundeEntries = entries.Where(e => kundeIPs.Contains(e.IP)).ToList();
                        if (kundeEntries.Count == 0) continue;

                        var kundeNode = new TreeNode($"👤  {k.Name}  ({kundeEntries.Count})")
                        {
                            Tag = null,
                            NodeFont = new Font("Segoe UI", 9, FontStyle.Bold)
                        };

                        foreach (var entry in kundeEntries)
                        {
                            var node = new TreeNode($"🖥  {entry.Alias}  —  {entry.IP}")
                            {
                                Tag = entry
                            };
                            kundeNode.Nodes.Add(node);
                        }
                        kundeNode.Expand();
                        tree.Nodes.Add(kundeNode);

                        // Bereits zugeordnete Einträge merken
                        foreach (var ent in kundeEntries)
                            kundenMap[ent.IP] = kundeNode;
                    }
                }
                catch { }
            }

            // Nicht zugeordnete Einträge unter "Sonstige"
            var unassigned = entries.Where(e => !kundenMap.ContainsKey(e.IP)).ToList();
            if (unassigned.Count > 0)
            {
                var sonstigeNode = new TreeNode($"📁  Sonstige  ({unassigned.Count})")
                {
                    NodeFont = new Font("Segoe UI", 9, FontStyle.Bold)
                };
                foreach (var entry in unassigned)
                {
                    var node = new TreeNode($"🖥  {entry.Alias}  —  {entry.IP}")
                    {
                        Tag = entry
                    };
                    sonstigeNode.Nodes.Add(node);
                }
                sonstigeNode.Expand();
                tree.Nodes.Add(sonstigeNode);
            }
        }

    }
}
