using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class SnmpSettingsDialog : Form
    {
        public SnmpSettings Settings { get; private set; }

        private ComboBox cmbVersion;
        private TextBox tbCommunity;
        private NumericUpDown nudPort, nudTimeout;
        // v3-Felder
        private TextBox tbUsername, tbAuthPw, tbPrivPw;
        private ComboBox cmbAuthProto, cmbPrivProto, cmbSecLevel;
        private Panel panelV3;

        public SnmpSettingsDialog(SnmpSettings current = null)
        {
            var s = current ?? new SnmpSettings();
            Settings = s;

            Text = "SNMP-Einstellungen";
            Width = 420; Height = 470;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 9);

            int y = 14;

            // ── Version ──────────────────────────────────────
            Controls.Add(MkLabel("SNMP-Version:", 14, y));
            cmbVersion = new ComboBox { Left = 130, Top = y, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbVersion.Items.AddRange(new object[] { "v1", "v2c", "v3" });
            cmbVersion.SelectedIndex = s.Version == 1 ? 0 : s.Version == 3 ? 2 : 1;
            cmbVersion.SelectedIndexChanged += (_, __) => UpdateV3Panel();
            Controls.Add(cmbVersion);
            y += 30;

            // ── Community (v1/v2c) ────────────────────────────
            Controls.Add(MkLabel("Community:", 14, y));
            tbCommunity = new TextBox { Left = 130, Top = y, Width = 150, Text = s.Community };
            Controls.Add(tbCommunity);
            y += 30;

            // ── Port ──────────────────────────────────────────
            Controls.Add(MkLabel("UDP-Port:", 14, y));
            nudPort = new NumericUpDown { Left = 130, Top = y, Width = 80, Minimum = 1, Maximum = 65535, Value = s.Port };
            Controls.Add(nudPort);
            y += 30;

            // ── Timeout ───────────────────────────────────────
            Controls.Add(MkLabel("Timeout (ms):", 14, y));
            nudTimeout = new NumericUpDown { Left = 130, Top = y, Width = 80, Minimum = 200, Maximum = 30000, Increment = 500, Value = s.TimeoutMs };
            Controls.Add(nudTimeout);
            y += 36;

            // ── Trennlinie ────────────────────────────────────
            var sep = new Label { Left = 14, Top = y, Width = 370, Height = 1, BackColor = Color.LightGray };
            Controls.Add(sep);
            y += 8;

            // ── SNMPv3-Panel ──────────────────────────────────
            panelV3 = new Panel { Left = 0, Top = y, Width = 410, Height = 200 };
            int py = 0;

            panelV3.Controls.Add(MkLabel("Benutzername:", 14, py));
            tbUsername = new TextBox { Left = 130, Top = py, Width = 200, Text = s.Username };
            panelV3.Controls.Add(tbUsername);
            py += 28;

            panelV3.Controls.Add(MkLabel("Auth-Protokoll:", 14, py));
            cmbAuthProto = new ComboBox { Left = 130, Top = py, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAuthProto.Items.AddRange(new object[] { "SHA256", "SHA384", "SHA512" });
            cmbAuthProto.SelectedItem = s.AuthProtocol == "SHA384" ? "SHA384" : s.AuthProtocol == "SHA512" ? "SHA512" : "SHA256";
            if (cmbAuthProto.SelectedIndex < 0) cmbAuthProto.SelectedIndex = 0;
            panelV3.Controls.Add(cmbAuthProto);
            py += 28;

            panelV3.Controls.Add(MkLabel("Auth-Passwort:", 14, py));
            tbAuthPw = new TextBox { Left = 130, Top = py, Width = 200, UseSystemPasswordChar = true, Text = s.AuthPassword };
            panelV3.Controls.Add(tbAuthPw);
            py += 28;

            panelV3.Controls.Add(MkLabel("Priv-Protokoll:", 14, py));
            cmbPrivProto = new ComboBox { Left = 130, Top = py, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbPrivProto.Items.AddRange(new object[] { "AES128", "AES192", "AES256" });
            cmbPrivProto.SelectedItem = s.PrivProtocol == "AES192" ? "AES192" : s.PrivProtocol == "AES256" ? "AES256" : "AES128";
            if (cmbPrivProto.SelectedIndex < 0) cmbPrivProto.SelectedIndex = 0;
            panelV3.Controls.Add(cmbPrivProto);
            py += 28;

            panelV3.Controls.Add(MkLabel("Priv-Passwort:", 14, py));
            tbPrivPw = new TextBox { Left = 130, Top = py, Width = 200, UseSystemPasswordChar = true, Text = s.PrivPassword };
            panelV3.Controls.Add(tbPrivPw);
            py += 28;

            panelV3.Controls.Add(MkLabel("Security Level:", 14, py));
            cmbSecLevel = new ComboBox { Left = 130, Top = py, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbSecLevel.Items.AddRange(new object[] { "NoAuth", "AuthNoPriv", "AuthPriv" });
            cmbSecLevel.SelectedItem = s.SecurityLevel ?? "AuthPriv";
            if (cmbSecLevel.SelectedIndex < 0) cmbSecLevel.SelectedIndex = 2;
            panelV3.Controls.Add(cmbSecLevel);

            Controls.Add(panelV3);

            // ── Buttons ───────────────────────────────────────
            var btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
            var btnCancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Width = 90 };
            btnOk.Click += (_, __) => ApplySettings();
            btnOk.Left = 220; btnOk.Top = Height - 74;
            btnCancel.Left = 310; btnCancel.Top = Height - 74;
            Controls.AddRange(new Control[] { btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            UpdateV3Panel();
        }

        private void UpdateV3Panel()
        {
            bool isV3 = cmbVersion.SelectedIndex == 2;
            panelV3.Visible = isV3;
            tbCommunity.Enabled = !isV3;
        }

        private void ApplySettings()
        {
            Settings.Version = cmbVersion.SelectedIndex == 0 ? 1 : cmbVersion.SelectedIndex == 2 ? 3 : 2;
            Settings.Community = tbCommunity.Text.Trim();
            Settings.Port = (int)nudPort.Value;
            Settings.TimeoutMs = (int)nudTimeout.Value;
            Settings.Username = tbUsername.Text.Trim();
            Settings.AuthPassword = tbAuthPw.Text;
            Settings.PrivPassword = tbPrivPw.Text;
            Settings.AuthProtocol = cmbAuthProto.SelectedItem?.ToString() ?? "SHA256";
            Settings.PrivProtocol = cmbPrivProto.SelectedItem?.ToString() ?? "AES128";
            Settings.SecurityLevel = cmbSecLevel.SelectedItem?.ToString() ?? "AuthPriv";
        }

        private static Label MkLabel(string text, int x, int y)
            => new Label { Text = text, Left = x, Top = y + 3, AutoSize = true };
    }
}
