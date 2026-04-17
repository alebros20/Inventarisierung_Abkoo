using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class LinuxSshConnectionForm : Form
    {
        public string Host        { get; private set; }
        public int    Port        { get; private set; } = 22;
        public string Username    { get; private set; }
        public string Password    { get; private set; }
        public string KeyFilePath { get; private set; } = "";

        private TextBox tbHost, tbUser, tbPass, tbKey;
        private NumericUpDown nudPort;

        public LinuxSshConnectionForm()
        {
            Text = "Linux SSH-Verbindung";
            Width = 420; Height = 310;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 9);

            int y = 16;

            Controls.Add(Lbl("Host / IP:", 14, y));
            tbHost = new TextBox { Left = 130, Top = y, Width = 240 };
            Controls.Add(tbHost);
            y += 30;

            Controls.Add(Lbl("SSH-Port:", 14, y));
            nudPort = new NumericUpDown { Left = 130, Top = y, Width = 70, Minimum = 1, Maximum = 65535, Value = 22 };
            Controls.Add(nudPort);
            y += 30;

            Controls.Add(Lbl("Benutzername:", 14, y));
            tbUser = new TextBox { Left = 130, Top = y, Width = 240 };
            Controls.Add(tbUser);
            y += 30;

            Controls.Add(Lbl("Passwort:", 14, y));
            tbPass = new TextBox { Left = 130, Top = y, Width = 240, UseSystemPasswordChar = true };
            Controls.Add(tbPass);
            y += 30;

            var sep = new Label { Left = 14, Top = y, Width = 370, Height = 1, BackColor = Color.LightGray };
            Controls.Add(sep);
            y += 10;

            Controls.Add(Lbl("SSH-Key (opt.):", 14, y));
            tbKey = new TextBox { Left = 130, Top = y, Width = 190 };
            Controls.Add(tbKey);
            var btnBrowse = new Button { Text = "...", Left = 326, Top = y - 1, Width = 44, Height = 24 };
            btnBrowse.Click += (s, e) =>
            {
                using (var dlg = new OpenFileDialog { Filter = "Private Key|*.pem;*.ppk;*.key|Alle Dateien|*.*" })
                    if (dlg.ShowDialog() == DialogResult.OK) tbKey.Text = dlg.FileName;
            };
            Controls.Add(btnBrowse);
            y += 36;

            var btnOk = new Button { Text = "Verbinden", DialogResult = DialogResult.OK, Width = 90, Left = 210, Top = y };
            var btnCancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Width = 90, Left = 308, Top = y };
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tbHost.Text)) { MessageBox.Show("Bitte Host angeben."); DialogResult = DialogResult.None; return; }
                if (string.IsNullOrWhiteSpace(tbUser.Text)) { MessageBox.Show("Bitte Benutzernamen angeben."); DialogResult = DialogResult.None; return; }
                if (string.IsNullOrWhiteSpace(tbPass.Text) && string.IsNullOrWhiteSpace(tbKey.Text))
                { MessageBox.Show("Bitte Passwort oder SSH-Key angeben."); DialogResult = DialogResult.None; return; }

                Host        = tbHost.Text.Trim();
                Port        = (int)nudPort.Value;
                Username    = tbUser.Text.Trim();
                Password    = tbPass.Text;
                KeyFilePath = tbKey.Text.Trim();
            };
            Controls.AddRange(new Control[] { btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        public void SetIP(string ip) => tbHost.Text = ip;

        private static Label Lbl(string text, int x, int y)
            => new Label { Text = text, Left = x, Top = y + 3, AutoSize = true };
    }
}
