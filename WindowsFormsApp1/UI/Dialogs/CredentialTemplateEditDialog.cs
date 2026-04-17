using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class CredentialTemplateEditDialog : Form
    {
        public CredentialTemplate Template { get; private set; }

        private TextBox txtName, txtUsername, txtPassword;
        private bool passwordVisible = false;

        public CredentialTemplateEditDialog(CredentialTemplate existing = null)
        {
            Text = existing == null ? "Neue Vorlage" : "Vorlage bearbeiten";
            Width = 460; Height = 240;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 9);

            int y = 14;

            // Name
            Controls.Add(new Label { Text = "Name:", Location = new Point(14, y), AutoSize = true });
            txtName = new TextBox { Location = new Point(140, y - 2), Width = 280 };
            Controls.Add(txtName);
            y += 32;

            // Username
            Controls.Add(new Label { Text = "Benutzername:", Location = new Point(14, y), AutoSize = true });
            txtUsername = new TextBox { Location = new Point(140, y - 2), Width = 280 };
            Controls.Add(txtUsername);
            y += 32;

            // Password + eye toggle
            Controls.Add(new Label { Text = "Passwort:", Location = new Point(14, y), AutoSize = true });
            txtPassword = new TextBox
            {
                Location = new Point(140, y - 2), Width = 244,
                UseSystemPasswordChar = true
            };
            Controls.Add(txtPassword);

            var btnEye = new Button
            {
                Text = "\ud83d\udc41", Location = new Point(388, y - 3),
                Width = 32, Height = 24, FlatStyle = FlatStyle.Flat
            };
            btnEye.Click += (s, e) =>
            {
                passwordVisible = !passwordVisible;
                txtPassword.UseSystemPasswordChar = !passwordVisible;
            };
            Controls.Add(btnEye);
            y += 40;

            // Buttons
            var btnOk = new Button
            {
                Text = "Speichern", DialogResult = DialogResult.OK,
                Location = new Point(240, y), Width = 90, Height = 30
            };
            var btnCancel = new Button
            {
                Text = "Abbrechen", DialogResult = DialogResult.Cancel,
                Location = new Point(336, y), Width = 90, Height = 30
            };
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // Load existing data
            if (existing != null)
            {
                txtName.Text = existing.Name;
                txtUsername.Text = existing.Username;
                if (!string.IsNullOrEmpty(existing.EncryptedPass) && SessionKey.IsSet)
                    txtPassword.Text = CredentialStore.Decrypt(existing.EncryptedPass);
            }

            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtName.Text))
                {
                    MessageBox.Show("Name darf nicht leer sein.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
                if (string.IsNullOrWhiteSpace(txtPassword.Text))
                {
                    MessageBox.Show("Passwort darf nicht leer sein.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                Template = new CredentialTemplate
                {
                    ID            = existing?.ID ?? 0,
                    Name          = txtName.Text.Trim(),
                    Username      = txtUsername.Text.Trim(),
                    EncryptedPass = CredentialStore.Encrypt(txtPassword.Text)
                };
            };
        }
    }
}
