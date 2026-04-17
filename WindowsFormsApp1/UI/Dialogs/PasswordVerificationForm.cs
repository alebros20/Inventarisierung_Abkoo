using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class PasswordVerificationForm : Form
    {
        private const string REQUIRED_PASSWORD = "Administrator";

        public PasswordVerificationForm(string action)
        {
            Text = "Sicherheitsabfrage";
            Width = 450; Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            var passwordTextBox = new TextBox { Location = new Point(20, 90), Width = 400, Font = new Font("Segoe UI", 11), UseSystemPasswordChar = true };
            var okBtn = new Button { Text = "Bestätigen", Location = new Point(150, 150), Width = 100 };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 150), Width = 100, DialogResult = DialogResult.Cancel };

            okBtn.Click += (s, e) =>
            {
                if (passwordTextBox.Text == REQUIRED_PASSWORD) { DialogResult = DialogResult.OK; Close(); }
                else
                {
                    MessageBox.Show("Falsches Passwort!", "Zugriff verweigert", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    passwordTextBox.Clear(); passwordTextBox.Focus();
                }
            };

            Controls.AddRange(new Control[] {
                new Label { Text = $"WARNUNG: Du bist dabei '{action}' auszuführen!\n\nGib das Sicherheitspasswort ein:", Location = new Point(20, 20), Width = 400, Height = 60, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.DarkRed },
                passwordTextBox,
                new Label { Text = $"Erforderliches Passwort: \"{REQUIRED_PASSWORD}\"", Location = new Point(20, 120), Width = 400, ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Italic) },
                okBtn, cancelBtn
            });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }
    }
}
