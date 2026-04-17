using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class CredentialPasswordDialog : Form
    {
        public string EnteredPassword { get; private set; }

        public CredentialPasswordDialog()
        {
            Text = "Zugangsdaten entsperren";
            Size = new Size(380, 190);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 10);

            var icon = new Label
            {
                Text = "🔒",
                Location = new Point(16, 18),
                AutoSize = true,
                Font = new Font("Segoe UI", 18)
            };
            var lblInfo = new Label
            {
                Text = "Passwort zum Entschlüsseln der gespeicherten Zugangsdaten:",
                Location = new Point(56, 14),
                Size = new Size(290, 40),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };

            var pwBox = new TextBox
            {
                Location = new Point(56, 60),
                Width = 280,
                UseSystemPasswordChar = true,
                Font = new Font("Segoe UI", 11)
            };

            var btnOk = new Button
            {
                Text = "Entsperren",
                Location = new Point(56, 96),
                Width = 130,
                Height = 28,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var btnCancel = new Button
            {
                Text = "Abbrechen",
                Location = new Point(196, 96),
                Width = 100,
                Height = 28,
                DialogResult = DialogResult.Cancel
            };

            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(pwBox.Text))
                {
                    MessageBox.Show("Bitte Passwort eingeben.", "Hinweis",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                    return;
                }

                // Passwort gegen Verifikationsdatei prüfen
                if (!CredentialStore.VerifyPassword(pwBox.Text))
                {
                    MessageBox.Show("Falsches Passwort.", "Fehler",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    pwBox.Clear();
                    pwBox.Focus();
                    DialogResult = DialogResult.None;
                    return;
                }

                EnteredPassword = pwBox.Text;
            };

            Controls.AddRange(new Control[] { icon, lblInfo, pwBox, btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
            Shown += (s, e) => pwBox.Focus();
        }
    }
}
