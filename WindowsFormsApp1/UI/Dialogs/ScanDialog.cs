using System.Drawing;
using System.Windows.Forms;

namespace NmapInventory
{
    /// <summary>
    /// Einheitlicher Scan-Dialog: Benutzer wählt den Scan-Modus.
    /// Ersetzt die 3 separaten Scan-Buttons im TopPanel.
    /// </summary>
    public class ScanDialog : Form
    {
        public enum ScanMode { None, NetworkDiscovery, RemoteScan, SnmpScan }

        public ScanMode SelectedMode { get; private set; } = ScanMode.None;

        public ScanDialog()
        {
            Text = "Scan starten";
            Size = new Size(420, 280);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 10);

            var header = new Label
            {
                Text = "Welchen Scan möchtest du starten?",
                Location = new Point(20, 16),
                AutoSize = true,
                Font = new Font("Segoe UI", 11, FontStyle.Bold)
            };

            // ── Netzwerk-Discovery ───────────────────────────
            var btnNetwork = new Button
            {
                Text = "🔍  Netzwerk scannen",
                Location = new Point(20, 56),
                Size = new Size(365, 42),
                Font = new Font("Segoe UI", 11),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };
            var lblNetwork = new Label
            {
                Text = "ARP-Discovery — findet alle Geräte im Subnetz",
                Location = new Point(35, 100),
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DarkSlateGray
            };
            btnNetwork.Click += (s, e) => { SelectedMode = ScanMode.NetworkDiscovery; DialogResult = DialogResult.OK; Close(); };

            // ── Remote-Scan ──────────────────────────────────
            var btnRemote = new Button
            {
                Text = "🖥  Remote Scan (WMI)",
                Location = new Point(20, 120),
                Size = new Size(365, 42),
                Font = new Font("Segoe UI", 11),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };
            var lblRemote = new Label
            {
                Text = "Hardware + Software per Remote-Verbindung abfragen",
                Location = new Point(35, 164),
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DarkSlateGray
            };
            btnRemote.Click += (s, e) => { SelectedMode = ScanMode.RemoteScan; DialogResult = DialogResult.OK; Close(); };

            // ── SNMP-Scan ────────────────────────────────────
            var btnSnmp = new Button
            {
                Text = "📡  SNMP-Scan",
                Location = new Point(20, 184),
                Size = new Size(365, 42),
                Font = new Font("Segoe UI", 11),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };
            var lblSnmp = new Label
            {
                Text = "SNMP-Abfrage aller gescannten Geräte",
                Location = new Point(35, 228),
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DarkSlateGray
            };
            btnSnmp.Click += (s, e) => { SelectedMode = ScanMode.SnmpScan; DialogResult = DialogResult.OK; Close(); };

            Controls.AddRange(new Control[] { header, btnNetwork, lblNetwork, btnRemote, lblRemote, btnSnmp, lblSnmp });
        }
    }
}
