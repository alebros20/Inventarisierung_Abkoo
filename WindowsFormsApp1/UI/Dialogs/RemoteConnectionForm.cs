using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    // Wrapper für ComboBox-Anzeige
    internal class CustomerItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name ?? "";
    }

    internal class IPItem
    {
        public string Display { get; set; } // Gerätename oder IP
        public string IP { get; set; } // immer die echte IP
        public override string ToString() => Display ?? IP ?? "";
    }

    public class RemoteConnectionForm : Form
    {
        public string ComputerIP { get; private set; }
        public string Username { get; private set; }
        public string Password { get; private set; }
        // Mehrfachauswahl aus Picker — enthält alle ausgewählten Einträge
        public List<(string IP, string Username, string Password)> SelectedTargets { get; private set; }
            = new List<(string, string, string)>();

        private TextBox ipTb, userTb, passTb;
        private Button btnDeleteEntry;

        private readonly DatabaseManager _db;

        public RemoteConnectionForm() : this(null) { }

        public RemoteConnectionForm(DatabaseManager db)
        {
            _db = db;
            Text = "Remote Verbindung";
            Width = 520; Height = 400;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Segoe UI", 9);

            // ── Gespeicherte Auswahl ──────────────────────────
            var lblSaved = new Label { Text = "Gespeichert:", Location = new Point(20, 14), AutoSize = true };
            var btnPickSaved = new Button
            {
                Text = "📋 Gespeicherte Zugangsdaten wählen...",
                Location = new Point(110, 10),
                Width = 270,
                Height = 24,
                Font = new Font("Segoe UI", 9)
            };
            btnPickSaved.Click += (s, e) =>
            {
                // Session entsperren falls nötig
                if (!SessionKey.IsSet)
                {
                    if (CredentialStore.GetAllRaw().Count == 0)
                    {
                        MessageBox.Show("Noch keine Zugangsdaten gespeichert.", "Hinweis",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    using (var pwDlg = new CredentialPasswordDialog())
                    {
                        if (pwDlg.ShowDialog() != DialogResult.OK) return;
                        SessionKey.Set(pwDlg.EnteredPassword);
                    }
                }

                using (var picker = new SavedCredentialsPicker(_db))
                {
                    if (picker.ShowDialog(this) == DialogResult.OK && picker.SelectedEntries.Count > 0)
                    {
                        if (picker.SelectedEntries.Count == 1)
                        {
                            // Einzelauswahl — Felder befüllen wie bisher
                            var sel = picker.SelectedEntries[0];
                            ipTb.Text = sel.IP;
                            userTb.Text = sel.Username;
                            passTb.Text = CredentialStore.Decrypt(sel.EncryptedPassword);
                            btnDeleteEntry.Enabled = true;
                        }
                        else
                        {
                            // Mehrfachauswahl — SelectedTargets befüllen + erstes Gerät anzeigen
                            SelectedTargets = picker.SelectedEntries
                                .Select(entry => (entry.IP, entry.Username, CredentialStore.Decrypt(entry.EncryptedPassword)))
                                .ToList();
                            var first = picker.SelectedEntries[0];
                            ipTb.Text = first.IP;
                            userTb.Text = first.Username;
                            passTb.Text = CredentialStore.Decrypt(first.EncryptedPassword);
                            btnDeleteEntry.Enabled = true;
                            // Hinweis anzeigen
                            MessageBox.Show(
                                $"{picker.SelectedEntries.Count} Geräte ausgewählt.\n" +
                                "Alle werden beim Klick auf 'Verbinden' nacheinander abgefragt.",
                                "Mehrfachauswahl", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            };


            // ── Kunden-IP Import ──────────────────────────────
            // Lädt IPs direkt aus der Kunden-Datenbank
            var lblKunde = new Label { Text = "Aus Kunde:", Location = new Point(20, 47), AutoSize = true };
            var cmbKunde = new ComboBox
            {
                Location = new Point(110, 44),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            var cmbIP = new ComboBox
            {
                Location = new Point(318, 44),
                Width = 165,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbKunde.Items.Add(new CustomerItem { ID = -1, Name = "-- Kunden wählen --" });
            cmbKunde.SelectedIndex = 0;

            // Kunden laden falls DB vorhanden
            if (_db != null)
            {
                try
                {
                    foreach (var k in _db.GetCustomers())
                        cmbKunde.Items.Add(new CustomerItem { ID = k.ID, Name = k.Name });
                }
                catch { }
            }

            // Bei Kunden-Auswahl: IPs aus allen Standorten laden
            cmbKunde.SelectedIndexChanged += (s, e) =>
            {
                cmbIP.Items.Clear();
                cmbIP.Items.Add(new IPItem { Display = "-- Gerät wählen --", IP = "" });
                cmbIP.SelectedIndex = 0;
                if (_db == null || !(cmbKunde.SelectedItem is CustomerItem ki) || ki.ID < 0) return;
                try
                {
                    var locs = _db.GetLocationsByCustomer(ki.ID);
                    foreach (var loc in locs)
                        foreach (var lip in _db.GetIPsWithWorkstationByLocation(loc.ID))
                            cmbIP.Items.Add(new IPItem
                            {
                                Display = string.IsNullOrEmpty(lip.WorkstationName) ? lip.IPAddress : lip.WorkstationName,
                                IP = lip.IPAddress
                            });
                }
                catch { }
            };

            // Bei IP-Auswahl: IP-Feld befüllen, Gerätename als Alias vorschlagen
            cmbIP.SelectedIndexChanged += (s, e) =>
            {
                if (!(cmbIP.SelectedItem is IPItem item) || string.IsNullOrEmpty(item.IP)) return;
                ipTb.Text = item.IP;

                // Alias vorschlagen
                var alias = Controls.Find("aliasTb", false).FirstOrDefault() as TextBox;
                if (alias != null && (alias.ForeColor == Color.Gray || string.IsNullOrEmpty(alias.Text)))
                {
                    alias.Text = item.Display;
                    alias.ForeColor = SystemColors.WindowText;
                }

                // Gespeicherte Zugangsdaten für diese IP laden
                // Falls Session noch nicht entsperrt: erst prüfen ob Eintrag existiert,
                // dann Passwort abfragen
                var allSaved = CredentialStore.GetAllRaw();
                var match = allSaved.FirstOrDefault(c => c.IP == item.IP);

                if (match != null)
                {
                    // Eintrag vorhanden — Session entsperren falls nötig
                    if (!SessionKey.IsSet)
                    {
                        using (var pwDlg = new CredentialPasswordDialog())
                        {
                            if (pwDlg.ShowDialog() == DialogResult.OK)
                                SessionKey.Set(pwDlg.EnteredPassword);
                        }
                    }

                    if (SessionKey.IsSet)
                    {
                        userTb.Text = match.Username;
                        passTb.Text = CredentialStore.Decrypt(match.EncryptedPassword);
                        btnDeleteEntry.Enabled = true;
                    }
                }
                else
                {
                    // Kein gespeicherter Eintrag — Felder zurücksetzen
                    userTb.Text = "admin";
                    passTb.Text = "";
                }
            };

            var sep2 = new Panel
            {
                Location = new Point(0, 74),
                Width = 520,
                Height = 1,
                BackColor = Color.FromArgb(200, 200, 200)
            };

            // ── Felder ────────────────────────────────────────
            var sep = new Panel
            {
                Location = new Point(0, 40),
                Width = 520,
                Height = 1,
                BackColor = Color.FromArgb(200, 200, 200)
            };

            ipTb = new TextBox
            {
                Name = "ipTextBox",
                Location = new Point(110, 84),
                Width = 310,
                Font = new Font("Segoe UI", 10)
            };
            userTb = new TextBox
            {
                Location = new Point(110, 116),
                Width = 310,
                Text = "admin",
                Font = new Font("Segoe UI", 10)
            };
            passTb = new TextBox
            {
                Location = new Point(110, 148),
                Width = 310,
                UseSystemPasswordChar = true,
                Font = new Font("Segoe UI", 10)
            };

            // ── Bezeichnung ───────────────────────────────────
            var lblAlias = new Label { Text = "Bezeichnung:", Location = new Point(20, 182), AutoSize = true };
            var aliasTb = new TextBox
            {
                Name = "aliasTb",
                Location = new Point(110, 179),
                Width = 310,
                Text = "z.B. Büro-PC",
                ForeColor = Color.Gray
            };
            aliasTb.GotFocus += (s, e) => { if (aliasTb.ForeColor == Color.Gray) { aliasTb.Text = ""; aliasTb.ForeColor = SystemColors.WindowText; } };
            aliasTb.LostFocus += (s, e) => { if (string.IsNullOrEmpty(aliasTb.Text)) { aliasTb.Text = "z.B. Büro-PC"; aliasTb.ForeColor = Color.Gray; } };

            var sep3 = new Panel
            {
                Location = new Point(0, 205),
                Width = 520,
                Height = 1,
                BackColor = Color.FromArgb(200, 200, 200)
            };

            var hinweis = new Label
            {
                Text = "AES-256 verschlüsselt — Passwort nur im RAM.",
                Location = new Point(20, 212),
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DarkSlateGray
            };

            // ── Button-Leiste ─────────────────────────────────
            // [Verbinden]  [Abbrechen]  (Zeile 1)
            // [💾 Speichern]  [🗑 Löschen]  (Zeile 2)
            var okBtn = new Button
            {
                Text = "Verbinden",
                Location = new Point(20, 232),
                Width = 120,
                Height = 28,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(148, 232),
                Width = 110,
                Height = 28,
                Font = new Font("Segoe UI", 10),
                DialogResult = DialogResult.Cancel
            };

            var btnSave = new Button
            {
                Text = "💾 Speichern",
                Location = new Point(20, 268),
                Width = 120,
                Height = 28,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnDeleteEntry = new Button
            {
                Text = "🗑 Löschen",
                Location = new Point(148, 268),
                Width = 110,
                Height = 28,
                Font = new Font("Segoe UI", 9),
                Enabled = false
            };

            // Speichern-Button
            btnSave.Click += (s, e) =>
            {
                string ip = ipTb.Text.Trim();
                string user = userTb.Text.Trim();
                string pass = passTb.Text;

                if (string.IsNullOrEmpty(ip) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                {
                    MessageBox.Show("Bitte IP, Benutzername und Passwort eingeben.",
                        "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Passwort abfragen falls noch nicht gesetzt
                if (!SessionKey.IsSet)
                {
                    using (var pwDlg = new CredentialPasswordDialog())
                    {
                        if (pwDlg.ShowDialog() != DialogResult.OK) return;
                        SessionKey.Set(pwDlg.EnteredPassword);
                    }
                }

                string aliasText = aliasTb.Text.Trim();
                if (string.IsNullOrEmpty(aliasText) || aliasText == "z.B. Büro-PC")
                    aliasText = $"{user}@{ip}";

                CredentialStore.Save(new CredentialEntry
                {
                    Alias = aliasText,
                    IP = ip,
                    Username = user,
                    EncryptedPassword = CredentialStore.Encrypt(pass)
                });

                MessageBox.Show($"Eintrag '{aliasText}' gespeichert.", "Gespeichert",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            // Einzelnen Eintrag löschen — anhand der aktuell eingetragenen IP
            btnDeleteEntry.Click += (s, e) =>
            {
                string ip = ipTb.Text.Trim();
                var match = CredentialStore.GetAllRaw().FirstOrDefault(c => c.IP == ip);
                if (match == null) { MessageBox.Show("Kein gespeicherter Eintrag für diese IP.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                if (MessageBox.Show($"Eintrag '{match.Alias}' löschen?", "Löschen",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    CredentialStore.Delete(match.Alias);
                    ipTb.Text = "";
                    userTb.Text = "admin";
                    passTb.Text = "";
                    btnDeleteEntry.Enabled = false;
                }
            };


            okBtn.Click += (s, e) =>
            {
                ComputerIP = ipTb.Text.Trim();
                Username = userTb.Text.Trim();
                Password = passTb.Text;

                if (string.IsNullOrEmpty(ComputerIP))
                {
                    MessageBox.Show("Bitte IP eingeben!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
                if (string.IsNullOrEmpty(Username))
                {
                    MessageBox.Show("Bitte Benutzernamen eingeben!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };

            Controls.AddRange(new Control[] {
                lblSaved, btnPickSaved, sep,
                lblKunde, cmbKunde, cmbIP, sep2,
                new Label { Text = "Computer-IP:",  Location = new Point(20, 87),  AutoSize = true },
                ipTb,
                new Label { Text = "Benutzername:", Location = new Point(20, 119), AutoSize = true },
                userTb,
                new Label { Text = "Passwort:",     Location = new Point(20, 151), AutoSize = true },
                passTb,
                lblAlias, aliasTb, sep3, hinweis,
                okBtn, cancelBtn, btnSave, btnDeleteEntry
            });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;

            // SessionKey beim Schließen löschen — muss jedes Mal neu eingegeben werden
            FormClosed += (s, e) => SessionKey.Clear();
        }

        public void SetIP(string ip)
        {
            if (ipTb != null) ipTb.Text = ip;

            // Gespeicherten Eintrag für diese IP automatisch laden
            var match = CredentialStore.GetAllRaw().FirstOrDefault(c => c.IP == ip);
            if (match != null && SessionKey.IsSet)
            {
                userTb.Text = match.Username;
                passTb.Text = CredentialStore.Decrypt(match.EncryptedPassword);
                btnDeleteEntry.Enabled = true;
            }
        }
    }
}
