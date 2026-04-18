using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    public class AdbConnectionDialog : Form
    {
        public string IpPort    { get; private set; }
        public bool   DoPairing { get; private set; }
        public string PairPort  { get; private set; }
        public string PairCode  { get; private set; }

        private TextBox tbIp, tbPort, tbPairPort, tbPairCode;
        private CheckBox cbPair;
        private Panel pairPanel;

        public AdbConnectionDialog(string defaultIp = "")
        {
            Text            = "Android ADB – Verbindung";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition   = FormStartPosition.CenterParent;
            MaximizeBox     = false;
            MinimizeBox     = false;
            ClientSize      = new Size(420, 310);

            var lblInfo = new Label
            {
                Text      = "Android 11+: Entwickleroptionen → Wireless Debugging\n" +
                            "Den dort angezeigten Port eintragen (z.B. :43215).\n" +
                            "Port 5555 gilt NUR für ältere Geräte (Android ≤10).",
                Left = 10, Top = 10, Width = 395, Height = 52,
                ForeColor = Color.DarkBlue
            };

            Controls.Add(lblInfo);
            Controls.Add(Lbl("IP-Adresse:", 10, 72));
            tbIp   = new TextBox { Left = 110, Top = 70, Width = 165, Text = defaultIp };
            Controls.Add(tbIp);
            Controls.Add(Lbl("Port:", 290, 72));
            tbPort = new TextBox { Left = 325, Top = 70, Width = 75, Text = "5555" };
            Controls.Add(tbPort);

            cbPair = new CheckBox
            {
                Text  = "Erst koppeln – Android 11+ (einmalig pro Gerät)",
                Left  = 10, Top = 108, Width = 390, Height = 22
            };
            cbPair.CheckedChanged += (s, e) => { pairPanel.Visible = cbPair.Checked; Height = cbPair.Checked ? 390 : 310; };
            Controls.Add(cbPair);

            pairPanel = new Panel { Left = 0, Top = 133, Width = 420, Height = 100, Visible = false };
            var lblPI = new Label
            {
                Text      = "Wireless Debugging → 'Mit Code koppeln'.\nDort werden Koppel-Port und 6-stelliger Code angezeigt:",
                Left = 10, Top = 2, Width = 395, Height = 38, ForeColor = Color.DarkGreen
            };
            pairPanel.Controls.Add(lblPI);
            pairPanel.Controls.Add(Lbl("Koppel-Port:", 10, 44));
            tbPairPort = new TextBox { Left = 100, Top = 42, Width = 70 };
            pairPanel.Controls.Add(tbPairPort);
            pairPanel.Controls.Add(Lbl("Code:", 190, 44));
            tbPairCode = new TextBox { Left = 230, Top = 42, Width = 90 };
            pairPanel.Controls.Add(tbPairCode);
            Controls.Add(pairPanel);

            var btnOk = new Button { Text = "Verbinden", Left = 225, Top = 268, Width = 90, Height = 28, DialogResult = DialogResult.OK };
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tbIp.Text) || string.IsNullOrWhiteSpace(tbPort.Text))
                { MessageBox.Show("IP und Port angeben.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (cbPair.Checked && (string.IsNullOrWhiteSpace(tbPairPort.Text) || string.IsNullOrWhiteSpace(tbPairCode.Text)))
                { MessageBox.Show("Koppel-Port und -Code angeben.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                IpPort = $"{tbIp.Text.Trim()}:{tbPort.Text.Trim()}";
                DoPairing = cbPair.Checked;
                PairPort  = tbPairPort.Text.Trim();
                PairCode  = tbPairCode.Text.Trim();
                DialogResult = DialogResult.OK;
                Close();
            };
            var btnCancel = new Button { Text = "Abbrechen", Left = 323, Top = 268, Width = 85, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private static Label Lbl(string text, int x, int y)
            => new Label { Text = text, Left = x, Top = y + 3, AutoSize = true };
    }
}
