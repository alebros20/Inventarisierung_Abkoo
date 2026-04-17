using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class LoginDialog : Form
    {
        public LoginDialog()
        {
            Text = "NmapInventory — Anmeldung";
            Size = new Size(400, 240);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 10);

            // Logo / Titel
            var title = new Label
            {
                Text = "NmapInventory",
                Location = new Point(20, 18),
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                AutoSize = true,
                ForeColor = Color.FromArgb(40, 80, 140)
            };
            var sub = new Label
            {
                Text = "Bitte Passwort eingeben um fortzufahren.",
                Location = new Point(20, 48),
                AutoSize = true,
                ForeColor = Color.DarkSlateGray
            };

            var lblPw = new Label { Text = "Passwort:", Location = new Point(20, 88), AutoSize = true };
            var pwBox = new TextBox
            {
                Location = new Point(100, 85),
                Width = 260,
                UseSystemPasswordChar = true,
                Font = new Font("Segoe UI", 10)
            };

            var btnOk = new Button
            {
                Text = "Anmelden",
                Location = new Point(100, 122),
                Width = 130,
                Height = 28,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var btnCancel = new Button
            {
                Text = "Abbrechen",
                Location = new Point(240, 122),
                Width = 120,
                Height = 28,
                DialogResult = DialogResult.Cancel
            };

            var lblHint = new Label
            {
                Text = "Beim ersten Start ein neues Passwort vergeben.",
                Location = new Point(20, 165),
                Size = new Size(350, 16),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray
            };

            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(pwBox.Text))
                {
                    MessageBox.Show("Bitte ein Passwort eingeben.", "Hinweis",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                    return;
                }
                SessionKey.Set(pwBox.Text);
            };

            Controls.AddRange(new Control[] {
                title, sub, lblPw, pwBox, btnOk, btnCancel, lblHint
            });
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // Passwort-Feld fokussieren
            Shown += (s, e) => pwBox.Focus();
        }
    }
}
