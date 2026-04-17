using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public class CredentialScanDialog : Form
    {
        private readonly DatabaseManager _db;
        private readonly List<DatabaseDevice> _devices;
        private ProgressBar progressBar;
        private RichTextBox logBox;
        private CheckBox chkRetest;
        private Button btnStart, btnCancel;
        private Label lblStatus;
        private CredentialScanner _scanner;
        private int _successCount;

        public CredentialScanDialog(DatabaseManager db, List<DatabaseDevice> devices)
        {
            _db = db;
            _devices = devices;
            Text = "Zugangsdaten testen";
            Width = 700; Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            Font = new Font("Segoe UI", 9);

            lblStatus = new Label
            {
                Text = $"{_devices.Count} Geräte bereit zum Testen",
                Dock = DockStyle.Top, Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 24, Maximum = _devices.Count };

            chkRetest = new CheckBox
            {
                Text = "Bereits getestete Geräte erneut prüfen",
                Dock = DockStyle.Top, Height = 26,
                Padding = new Padding(8, 4, 0, 0)
            };

            logBox = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White
            };

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            btnStart = new Button
            {
                Text = "🔑 Test starten", Location = new Point(6, 8),
                Width = 140, Height = 28,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnCancel = new Button
            {
                Text = "Schließen", Location = new Point(574, 8),
                Width = 100, Height = 28, DialogResult = DialogResult.Cancel
            };

            btnStart.Click += async (s, e) => await StartScan();
            btnPanel.Controls.AddRange(new Control[] { btnStart, btnCancel });

            Controls.Add(logBox);
            Controls.Add(progressBar);
            Controls.Add(chkRetest);
            Controls.Add(lblStatus);
            Controls.Add(btnPanel);
            CancelButton = btnCancel;
        }

        private async System.Threading.Tasks.Task StartScan()
        {
            if (!SessionKey.IsSet)
            {
                using (var dlg = new CredentialPasswordDialog())
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
            }

            var templates = _db.GetCredentialTemplates();
            if (templates.Count == 0)
            {
                MessageBox.Show("Keine Passwort-Vorlagen vorhanden.\n\nBitte zuerst Vorlagen anlegen.",
                    "Keine Vorlagen", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnStart.Enabled = false;
            chkRetest.Enabled = false;
            btnCancel.Text = "Abbrechen";
            _successCount = 0;
            logBox.Clear();
            progressBar.Value = 0;

            _scanner = new CredentialScanner(5);

            _scanner.DeviceScanned += result =>
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => OnDeviceScanned(result)));
                    return;
                }
                OnDeviceScanned(result);
            };

            _scanner.ProgressChanged += (current, total) =>
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        progressBar.Value = current;
                        lblStatus.Text = $"Fortschritt: {current}/{total} Geräte";
                    }));
                    return;
                }
                progressBar.Value = current;
                lblStatus.Text = $"Fortschritt: {current}/{total} Geräte";
            };

            btnCancel.Click += (s, e) => _scanner.Cancel();

            var results = await _scanner.ScanDevicesAsync(_devices, templates, chkRetest.Checked);

            // Save successful matches to DB + UniqueID + Interfaces
            int mergeCount = 0;
            foreach (var r in results.Where(r => r.Success && r.MatchedTemplate != null))
            {
                _db.SetDeviceCredential(r.Device.ID, r.MatchedTemplate.ID);

                if (r.Identity != null && !string.IsNullOrEmpty(r.Identity.UniqueID))
                {
                    // Prüfen ob bereits ein anderes Gerät diese UniqueID hat → Merge
                    int existingId = _db.FindDeviceByUniqueID(r.Identity.UniqueID);
                    if (existingId > 0 && existingId != r.Device.ID)
                    {
                        _db.MergeDevices(existingId, r.Device.ID);
                        AppendLog($"  ↳ Duplikat erkannt: Gerät #{r.Device.ID} → #{existingId} zusammengeführt\n",
                            Color.FromArgb(255, 200, 100));
                        mergeCount++;
                        // Interfaces auf das Zielgerät speichern
                        foreach (var iface in r.Identity.Interfaces)
                            _db.AddDeviceInterface(existingId, iface.Mac, iface.IP, iface.Type);
                    }
                    else
                    {
                        _db.SetDeviceUniqueID(r.Device.ID, r.Identity.UniqueID, r.Identity.Source);
                        foreach (var iface in r.Identity.Interfaces)
                            _db.AddDeviceInterface(r.Device.ID, iface.Mac, iface.IP, iface.Type);
                    }
                }
            }

            string mergeInfo = mergeCount > 0 ? $", {mergeCount} Duplikat(e) zusammengeführt" : "";
            lblStatus.Text = $"Fertig: {_successCount}/{_devices.Count} Geräte authentifiziert{mergeInfo}";
            btnStart.Enabled = true;
            chkRetest.Enabled = true;
            btnCancel.Text = "Schließen";
        }

        private void OnDeviceScanned(CredentialScanner.ScanResult result)
        {
            string host = !string.IsNullOrEmpty(result.Device.Hostname) ? result.Device.Hostname : result.Device.IP;
            if (result.Success)
            {
                _successCount++;
                AppendLog($"✔ {result.Device.IP} ({host}) — {result.Message}\n", Color.LightGreen);
            }
            else
            {
                AppendLog($"✘ {result.Device.IP} ({host}) — {result.Message}\n", Color.FromArgb(255, 150, 150));
            }
        }

        private void AppendLog(string text, Color color)
        {
            logBox.SelectionStart = logBox.TextLength;
            logBox.SelectionLength = 0;
            logBox.SelectionColor = color;
            logBox.AppendText(text);
            logBox.ScrollToCaret();
        }
    }
}
