using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public class CredentialTemplateDialog : Form
    {
        private readonly DatabaseManager _db;
        private DataGridView grid;

        public CredentialTemplateDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Passwort-Vorlagen verwalten";
            Width = 820; Height = 480;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9);
            MinimizeBox = false;

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                MultiSelect = false
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID", HeaderText = "ID", Width = 40, Visible = false });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Name", FillWeight = 35 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Username", HeaderText = "Benutzer", FillWeight = 30 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Password", HeaderText = "Passwort", FillWeight = 20 });

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            var btnAdd = new Button { Text = "Hinzufügen", Location = new Point(6, 8), Width = 110, Height = 28 };
            var btnEdit = new Button { Text = "Bearbeiten", Location = new Point(122, 8), Width = 110, Height = 28 };
            var btnDelete = new Button { Text = "Löschen", Location = new Point(238, 8), Width = 110, Height = 28 };
            var btnClose = new Button { Text = "Schließen", Location = new Point(694, 8), Width = 100, Height = 28, DialogResult = DialogResult.Cancel };

            btnAdd.Click += (s, e) => AddTemplate();
            btnEdit.Click += (s, e) => EditTemplate();
            btnDelete.Click += (s, e) => DeleteTemplate();
            grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) EditTemplate(); };

            btnPanel.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDelete, btnClose });
            Controls.Add(grid);
            Controls.Add(btnPanel);
            CancelButton = btnClose;

            EnsureMasterPassword();
            LoadTemplates();
        }

        private void EnsureMasterPassword()
        {
            if (!SessionKey.IsSet)
            {
                using (var dlg = new CredentialPasswordDialog())
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK)
                    {
                        Close();
                        return;
                    }
                }
            }
        }

        private void LoadTemplates()
        {
            grid.Rows.Clear();
            foreach (var t in _db.GetCredentialTemplates())
                grid.Rows.Add(t.ID, t.Name, t.Username, "••••••");
        }

        private void AddTemplate()
        {
            using (var dlg = new CredentialTemplateEditDialog())
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Template != null)
                {
                    _db.AddCredentialTemplate(dlg.Template);
                    LoadTemplates();
                }
            }
        }

        private void EditTemplate()
        {
            if (grid.CurrentRow == null) return;
            int id = Convert.ToInt32(grid.CurrentRow.Cells["ID"].Value);
            var template = _db.GetCredentialTemplates().FirstOrDefault(t => t.ID == id);
            if (template == null) return;

            using (var dlg = new CredentialTemplateEditDialog(template))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Template != null)
                {
                    _db.UpdateCredentialTemplate(dlg.Template);
                    LoadTemplates();
                }
            }
        }

        private void DeleteTemplate()
        {
            if (grid.CurrentRow == null) return;
            string name = grid.CurrentRow.Cells["Name"].Value?.ToString();
            if (MessageBox.Show($"Vorlage \"{name}\" wirklich löschen?\n\nZugewiesene Geräte werden zurückgesetzt.",
                "Löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            int id = Convert.ToInt32(grid.CurrentRow.Cells["ID"].Value);
            _db.DeleteCredentialTemplate(id);
            LoadTemplates();
        }
    }
}
