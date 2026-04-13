using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace NmapInventory
{
    // =========================================================
    // === DATA EXPORTER ===
    // =========================================================
    public class DataExporter
    {
        public void Export(string filePath, List<DeviceInfo> devices, List<SoftwareInfo> software, string hardware)
        {
            if (filePath.EndsWith(".json"))
            {
                var data = new { Zeitstempel = DateTime.Now, Geräte = devices, Software = software, Hardware = hardware };
                System.IO.File.WriteAllText(filePath, Newtonsoft.Json.JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented));
            }
            else
            {
                var doc = new System.Xml.Linq.XDocument(
                    new System.Xml.Linq.XDeclaration("1.0", "utf-8", "yes"),
                    new System.Xml.Linq.XElement("Inventar",
                        new System.Xml.Linq.XElement("Zeitstempel", DateTime.Now),
                        new System.Xml.Linq.XElement("Geräte", devices.Select(d => new System.Xml.Linq.XElement("Gerät",
                            new System.Xml.Linq.XElement("IP", d.IP),
                            new System.Xml.Linq.XElement("Hostname", d.Hostname ?? "")))),
                        new System.Xml.Linq.XElement("Hardware", new System.Xml.Linq.XCData(hardware))));
                doc.Save(filePath);
            }
        }
    }

    // =========================================================
    // === REMOTE CONNECTION FORM ===
    // =========================================================
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
            Font = new System.Drawing.Font("Segoe UI", 9);

            // ── Gespeicherte Auswahl ──────────────────────────
            var lblSaved = new Label { Text = "Gespeichert:", Location = new Point(20, 14), AutoSize = true };
            var btnPickSaved = new Button
            {
                Text = "📋 Gespeicherte Zugangsdaten wählen...",
                Location = new Point(110, 10),
                Width = 270,
                Height = 24,
                Font = new System.Drawing.Font("Segoe UI", 9)
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
                BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            };

            // ── Felder ────────────────────────────────────────
            var sep = new Panel
            {
                Location = new Point(0, 40),
                Width = 520,
                Height = 1,
                BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            };

            ipTb = new TextBox
            {
                Name = "ipTextBox",
                Location = new Point(110, 84),
                Width = 310,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            userTb = new TextBox
            {
                Location = new Point(110, 116),
                Width = 310,
                Text = "admin",
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            passTb = new TextBox
            {
                Location = new Point(110, 148),
                Width = 310,
                UseSystemPasswordChar = true,
                Font = new System.Drawing.Font("Segoe UI", 10)
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
                BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            };

            var hinweis = new Label
            {
                Text = "AES-256 verschlüsselt — Passwort nur im RAM.",
                Location = new Point(20, 212),
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 8),
                ForeColor = System.Drawing.Color.DarkSlateGray
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
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(148, 232),
                Width = 110,
                Height = 28,
                Font = new System.Drawing.Font("Segoe UI", 10),
                DialogResult = DialogResult.Cancel
            };

            var btnSave = new Button
            {
                Text = "💾 Speichern",
                Location = new Point(20, 268),
                Width = 120,
                Height = 28,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };
            btnDeleteEntry = new Button
            {
                Text = "🗑 Löschen",
                Location = new Point(148, 268),
                Width = 110,
                Height = 28,
                Font = new System.Drawing.Font("Segoe UI", 9),
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

                string alias = aliasTb.Text.Trim();
                if (string.IsNullOrEmpty(alias) || alias == "z.B. Büro-PC")
                    alias = $"{user}@{ip}";

                CredentialStore.Save(new CredentialEntry
                {
                    Alias = alias,
                    IP = ip,
                    Username = user,
                    EncryptedPassword = CredentialStore.Encrypt(pass)
                });

                MessageBox.Show($"Eintrag '{alias}' gespeichert.", "Gespeichert",
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

    // =========================================================
    // === CREDENTIAL STORE (Windows DPAPI) ===
    // Speichert Zugangsdaten verschlüsselt mit Windows-Benutzerkonto.
    // Nur auf dem gleichen PC + Benutzer entschlüsselbar.
    // =========================================================

    // =========================================================
    // === GESPEICHERTE ZUGANGSDATEN — AUSWAHL-POPUP ===
    // Zeigt alle gespeicherten Einträge nach Kunden gruppiert.
    // =========================================================
    public class SavedCredentialsPicker : Form
    {
        // Für Einzelauswahl (Kompatibilität)
        public CredentialEntry Selected => SelectedEntries.Count > 0 ? SelectedEntries[0] : null;
        // Für Mehrfachauswahl
        public List<CredentialEntry> SelectedEntries { get; private set; } = new List<CredentialEntry>();

        private readonly DatabaseManager _db;
        private TreeView tree;
        private Label lblCount;

        public SavedCredentialsPicker(DatabaseManager db)
        {
            _db = db;
            Text = "Gespeicherte Zugangsdaten auswählen";
            Size = new Size(440, 530);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 9);

            var header = new Label
            {
                Text = "Geräte auswählen (Mehrfachauswahl möglich):",
                Location = new Point(12, 12),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            tree = new TreeView
            {
                Location = new Point(12, 34),
                Size = new Size(406, 380),
                FullRowSelect = true,
                ShowLines = true,
                ShowPlusMinus = true,
                CheckBoxes = true,
                Font = new Font("Segoe UI", 9)
            };

            LoadTree();

            // Kunden-Knoten anklicken → alle Kinder an/abwählen
            tree.AfterCheck += (s, e) =>
            {
                if (e.Action == TreeViewAction.Unknown) return;
                // Kunden-Gruppe: alle Kinder synchronisieren
                if (e.Node.Tag == null)
                {
                    foreach (TreeNode child in e.Node.Nodes)
                        child.Checked = e.Node.Checked;
                }
                // Gerät: Eltern-Status aktualisieren
                else if (e.Node.Parent != null)
                {
                    bool allChecked = true;
                    foreach (TreeNode sibling in e.Node.Parent.Nodes)
                    {
                        if (!sibling.Checked) allChecked = false;
                    }
                    e.Node.Parent.Checked = allChecked;
                }
                UpdateCount();
            };

            // Doppelklick = Einzelauswahl sofort bestätigen
            tree.NodeMouseDoubleClick += (s, e) =>
            {
                if (e.Node?.Tag is CredentialEntry)
                {
                    e.Node.Checked = true;
                    ConfirmSelection();
                }
            };

            lblCount = new Label
            {
                Text = "0 Geräte ausgewählt",
                Location = new Point(12, 420),
                Size = new Size(280, 18),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray
            };

            // Alle / Keine Buttons
            var btnAll = new Button
            {
                Text = "✔ Alle",
                Location = new Point(290, 416),
                Width = 60,
                Height = 24,
                Font = new Font("Segoe UI", 8)
            };
            btnAll.Click += (s, e) =>
            {
                foreach (TreeNode root in tree.Nodes)
                    foreach (TreeNode child in root.Nodes)
                        child.Checked = true;
                UpdateCount();
            };

            var btnNone = new Button
            {
                Text = "✖ Keine",
                Location = new Point(356, 416),
                Width = 62,
                Height = 24,
                Font = new Font("Segoe UI", 8)
            };
            btnNone.Click += (s, e) =>
            {
                foreach (TreeNode root in tree.Nodes)
                    foreach (TreeNode child in root.Nodes)
                        child.Checked = false;
                UpdateCount();
            };

            var btnOk = new Button
            {
                Text = "Abfragen",
                Location = new Point(12, 448),
                Width = 130,
                Height = 30,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) => ConfirmSelection();

            var btnCancel = new Button
            {
                Text = "Abbrechen",
                Location = new Point(152, 448),
                Width = 110,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[] { header, tree, lblCount, btnAll, btnNone, btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void UpdateCount()
        {
            int count = 0;
            foreach (TreeNode root in tree.Nodes)
                foreach (TreeNode child in root.Nodes)
                    if (child.Checked) count++;
            lblCount.Text = count == 0 ? "Kein Gerät ausgewählt"
                : count == 1 ? "1 Gerät ausgewählt"
                : $"{count} Geräte ausgewählt";
        }

        private void ConfirmSelection()
        {
            SelectedEntries = new List<CredentialEntry>();
            foreach (TreeNode root in tree.Nodes)
                foreach (TreeNode child in root.Nodes)
                    if (child.Checked && child.Tag is CredentialEntry entry)
                        SelectedEntries.Add(entry);

            if (SelectedEntries.Count == 0)
            {
                MessageBox.Show("Bitte mindestens ein Gerät auswählen.", "Hinweis",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void LoadTree()
        {
            tree.Nodes.Clear();
            var entries = CredentialStore.GetAllRaw();
            if (entries.Count == 0)
            {
                tree.Nodes.Add("(Keine gespeicherten Einträge)");
                return;
            }

            // Kunden aus DB laden für Gruppierung
            var kundenMap = new Dictionary<string, TreeNode>();

            if (_db != null)
            {
                try
                {
                    foreach (var k in _db.GetCustomers())
                    {
                        // IPs dieses Kunden sammeln
                        var kundeIPs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var loc in _db.GetLocationsByCustomer(k.ID))
                            foreach (var lip in _db.GetIPsWithWorkstationByLocation(loc.ID))
                                kundeIPs.Add(lip.IPAddress);

                        // Gespeicherte Einträge die zu diesem Kunden gehören
                        var kundeEntries = entries.Where(e => kundeIPs.Contains(e.IP)).ToList();
                        if (kundeEntries.Count == 0) continue;

                        var kundeNode = new TreeNode($"👤  {k.Name}  ({kundeEntries.Count})")
                        {
                            Tag = null,
                            NodeFont = new Font("Segoe UI", 9, FontStyle.Bold)
                        };

                        foreach (var entry in kundeEntries)
                        {
                            var node = new TreeNode($"🖥  {entry.Alias}  —  {entry.IP}")
                            {
                                Tag = entry
                            };
                            kundeNode.Nodes.Add(node);
                        }
                        kundeNode.Expand();
                        tree.Nodes.Add(kundeNode);

                        // Bereits zugeordnete Einträge merken
                        foreach (var e in kundeEntries)
                            kundenMap[e.IP] = kundeNode;
                    }
                }
                catch { }
            }

            // Nicht zugeordnete Einträge unter "Sonstige"
            var unassigned = entries.Where(e => !kundenMap.ContainsKey(e.IP)).ToList();
            if (unassigned.Count > 0)
            {
                var sonstigeNode = new TreeNode($"📁  Sonstige  ({unassigned.Count})")
                {
                    NodeFont = new Font("Segoe UI", 9, FontStyle.Bold)
                };
                foreach (var entry in unassigned)
                {
                    var node = new TreeNode($"🖥  {entry.Alias}  —  {entry.IP}")
                    {
                        Tag = entry
                    };
                    sonstigeNode.Nodes.Add(node);
                }
                sonstigeNode.Expand();
                tree.Nodes.Add(sonstigeNode);
            }
        }

    }

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

    public class CredentialEntry
    {
        public string Alias { get; set; }
        public string IP { get; set; }
        public string Username { get; set; }
        public string EncryptedPassword { get; set; }

        public override string ToString() => $"{Alias}  ({Username}@{IP})";
    }

    public static class CredentialStore
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NmapInventory", "credentials.dat");

        // Verifikationsdatei — enthält verschlüsselten Prüftext
        private static readonly string VerifyPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NmapInventory", "credentials.verify");

        private const string VerifyMagic = "NMAP_INVENTORY_OK"; // bekannter Klartext

        // ── PBKDF2 + AES-256 ─────────────────────────────────
        private const int KeySize = 32;
        private const int IVSize = 16;
        private const int SaltSize = 16;
        private const int Iterations = 10000;

        private static byte[] DeriveKey(string password, byte[] salt)
        {
            using (var kdf = new Rfc2898DeriveBytes(password, salt, Iterations, HashAlgorithmName.SHA256))
                return kdf.GetBytes(KeySize);
        }

        // Prüft ob das Passwort korrekt ist.
        // Beim allerersten Aufruf: Verifikationsdatei anlegen und true zurückgeben.
        // Danach: Nur true wenn HMAC-Signatur stimmt.
        // Erster Aufruf — Verifikationsdatei mit Standardpasswort 1234 anlegen
        public static bool VerifyPassword(string password)
        {
            if (!File.Exists(VerifyPath))
            {
                CreateVerifyFile("1234");
                // Prüfen ob eingegebenes Passwort dem Standard entspricht  Veschlüsselung
                byte[] data = Convert.FromBase64String(File.ReadAllText(VerifyPath).Trim());
                byte[] salt = new byte[SaltSize]; Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                byte[] stored = new byte[32]; Buffer.BlockCopy(data, SaltSize, stored, 0, 32);
                byte[] computed = ComputeHmac(password, salt);
                int diff = 0;
                for (int i = 0; i < computed.Length; i++) diff |= computed[i] ^ stored[i];
                return diff == 0;
            }
            try
            {
                // Datei lesen: Salt(16) + HMAC(32)
                byte[] data = Convert.FromBase64String(File.ReadAllText(VerifyPath).Trim());
                if (data.Length < SaltSize + 32) return false;

                byte[] salt = new byte[SaltSize];
                byte[] stored = new byte[32];
                Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                Buffer.BlockCopy(data, SaltSize, stored, 0, 32);

                byte[] computed = ComputeHmac(password, salt);
                // Zeitkonstanter Vergleich gegen Timing-Angriffe
                if (computed.Length != stored.Length) return false;
                int diff = 0;
                for (int i = 0; i < computed.Length; i++)
                    diff |= computed[i] ^ stored[i];
                return diff == 0;
            }
            catch { return false; }
        }

        // Erzeugt HMAC-SHA256 aus Passwort + Salt (PBKDF2-abgeleiteter Schlüssel)
        private static byte[] ComputeHmac(string password, byte[] salt)
        {
            byte[] key = DeriveKey(password, salt);
            using (var hmac = new HMACSHA256(key))
                return hmac.ComputeHash(Encoding.UTF8.GetBytes(VerifyMagic));
        }

        // Legt die Verifikationsdatei an
        private static void CreateVerifyFile(string password)
        {
            try
            {
                byte[] salt = new byte[SaltSize];
                using (var rng = RandomNumberGenerator.Create())
                    rng.GetBytes(salt);

                byte[] hmac = ComputeHmac(password, salt);
                byte[] result = new byte[SaltSize + hmac.Length];
                Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
                Buffer.BlockCopy(hmac, 0, result, SaltSize, hmac.Length);

                Directory.CreateDirectory(Path.GetDirectoryName(VerifyPath));
                File.WriteAllText(VerifyPath, Convert.ToBase64String(result));
            }
            catch { }
        }

        // Speichert das Passwort als Verifikationsdatei (beim ersten Speichern)
        private static void SaveVerification(string password)
        {
            if (File.Exists(VerifyPath)) return;
            CreateVerifyFile(password);
        }

        public static string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            string password = SessionKey.Current;
            if (string.IsNullOrEmpty(password)) return "";
            try
            {
                byte[] salt = new byte[SaltSize];
                using (var rng = RandomNumberGenerator.Create())
                    rng.GetBytes(salt);

                byte[] key = DeriveKey(password, salt);
                using (var aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.GenerateIV();
                    using (var enc = aes.CreateEncryptor())
                    {
                        byte[] plain = Encoding.UTF8.GetBytes(plainText);
                        byte[] cipher = enc.TransformFinalBlock(plain, 0, plain.Length);
                        // Format: Salt(16) + IV(16) + CipherText
                        byte[] result = new byte[SaltSize + IVSize + cipher.Length];
                        Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
                        Buffer.BlockCopy(aes.IV, 0, result, SaltSize, IVSize);
                        Buffer.BlockCopy(cipher, 0, result, SaltSize + IVSize, cipher.Length);
                        return Convert.ToBase64String(result);
                    }
                }
            }
            catch { return ""; }
        }

        public static string Decrypt(string encryptedBase64)
        {
            if (string.IsNullOrEmpty(encryptedBase64)) return "";
            string password = SessionKey.Current;
            if (string.IsNullOrEmpty(password)) return "";
            try
            {
                byte[] data = Convert.FromBase64String(encryptedBase64);
                byte[] salt = new byte[SaltSize];
                byte[] iv = new byte[IVSize];
                byte[] cipher = new byte[data.Length - SaltSize - IVSize];
                Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                Buffer.BlockCopy(data, SaltSize, iv, 0, IVSize);
                Buffer.BlockCopy(data, SaltSize + IVSize, cipher, 0, cipher.Length);

                byte[] key = DeriveKey(password, salt);
                using (var aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    using (var dec = aes.CreateDecryptor())
                    {
                        byte[] plain = dec.TransformFinalBlock(cipher, 0, cipher.Length);
                        return Encoding.UTF8.GetString(plain);
                    }
                }
            }
            catch { return ""; }
        }

        // Liest alle Einträge roh (ohne Entschlüsselung) — für Existenzprüfung
        public static List<CredentialEntry> GetAllRaw()
        {
            var list = new List<CredentialEntry>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                foreach (var line in File.ReadAllLines(FilePath))
                {
                    var parts = line.Split('\t');
                    if (parts.Length == 4)
                        list.Add(new CredentialEntry
                        {
                            Alias = parts[0],
                            IP = parts[1],
                            Username = parts[2],
                            EncryptedPassword = parts[3]
                        });
                }
            }
            catch { }
            return list;
        }

        public static List<CredentialEntry> GetAll() => GetAllRaw();

        public static void Save(CredentialEntry entry)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));

            // Verifikationsdatei beim ersten Speichern anlegen
            SaveVerification(SessionKey.Current);
            var all = GetAll();

            // Bestehenden Eintrag mit gleichem Alias überschreiben
            var existing = all.FirstOrDefault(e =>
                e.Alias == entry.Alias || (e.IP == entry.IP && e.Username == entry.Username));
            if (existing != null) all.Remove(existing);
            all.Add(entry);

            File.WriteAllLines(FilePath, all.Select(e =>
                $"{e.Alias}	{e.IP}	{e.Username}	{e.EncryptedPassword}"));
        }

        public static void Delete(string alias)
        {
            if (!File.Exists(FilePath)) return;
            var all = GetAll().Where(e => e.Alias != alias).ToList();
            File.WriteAllLines(FilePath, all.Select(e =>
                $"{e.Alias}	{e.IP}	{e.Username}	{e.EncryptedPassword}"));
        }
    }

    // =========================================================
    // === PASSWORD VERIFICATION FORM ===
    // =========================================================
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
            var okBtn = new Button { Text = "Bestätigen", Location = new Point(150, 150), Width = 100, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 150), Width = 100, DialogResult = DialogResult.Cancel, BackColor = Color.LightCoral };

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

    // =========================================================
    // === INPUT DIALOG ===
    // =========================================================
    public class InputDialog : Form
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }

        public InputDialog(string title, string label1, string label2)
        {
            Text = title;
            Width = 450; Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            var txt1 = new TextBox { Location = new Point(20, 45), Width = 400 };
            var txt2 = new TextBox { Location = new Point(20, 105), Width = 400, Multiline = true, Height = 50 };
            var okBtn = new Button { Text = "OK", Location = new Point(220, 165), Width = 90, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(320, 165), Width = 100, DialogResult = DialogResult.Cancel };

            okBtn.Click += (s, e) => { Value1 = txt1.Text; Value2 = txt2.Text; };
            this.Load += (s, e) => { txt1.Text = Value1 ?? ""; txt2.Text = Value2 ?? ""; };

            Controls.AddRange(new Control[] {
                new Label { Text = label1, Location = new Point(20, 20), AutoSize = true },
                txt1,
                new Label { Text = label2, Location = new Point(20, 80), AutoSize = true },
                txt2,
                okBtn, cancelBtn
            });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }
    }

    // =========================================================
    // === CUSTOMER SELECTION FORM ===
    // =========================================================
    public class CustomerSelectionForm : Form
    {
        public int SelectedCustomerID { get; private set; }

        public CustomerSelectionForm(List<Customer> customers)
        {
            Text = "Kunde auswählen";
            Width = 400; Height = 200;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var label = new Label { Text = "Wähle einen Kunden aus:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 11) };
            var combo = new ComboBox { Location = new Point(20, 50), Width = 350, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            foreach (var customer in customers)
                combo.Items.Add(new ComboCustomerItem { ID = customer.ID, Name = customer.Name });
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;

            var okBtn = new Button { Text = "OK", Location = new Point(150, 120), Width = 100, DialogResult = DialogResult.OK, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 120), Width = 100, DialogResult = DialogResult.Cancel, Font = new Font("Segoe UI", 10) };

            okBtn.Click += (s, e) =>
            {
                if (combo.SelectedItem is ComboCustomerItem item)
                    SelectedCustomerID = item.ID;
            };

            Controls.AddRange(new Control[] { label, combo, okBtn, cancelBtn });
            AcceptButton = okBtn; CancelButton = cancelBtn;
        }

        // Typsicheres ComboBox-Item (kein dynamic mehr)
        private class ComboCustomerItem
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public override string ToString() => Name;
        }
    }

    // =========================================================
    // === IP IMPORT DIALOG ===
    //
    // Zeigt alle gescannten Geräte aus der Devices-Tabelle an.
    // Bereits zugewiesene IPs werden ausgeblendet.
    // Benutzer wählt per Checkbox welche IPs importiert werden.
    //
    // Rückgabe über SelectedIPs als List<Tuple<string, string>>
    // (IP, Arbeitsplatzname) — kompatibel mit Dialogs.cs Aufruf.
    // =========================================================
    public class IPImportDialog : Form
    {
        // Tuple statt ValueTuple — kompatibel mit dem Aufruf in MainForm
        public List<Tuple<string, string>> SelectedIPs { get; private set; }
            = new List<Tuple<string, string>>();

        private CheckedListBox checkedListBox;
        private List<DeviceListItem> allDevices;
        private TextBox searchBox;
        private Label countLabel;

        public IPImportDialog(DatabaseManager dbManager, int targetLocationId, string locationName)
        {
            Text = $"IPs importieren → {locationName}";
            Width = 600; Height = 610;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 10);

            // Bereits zugewiesene IPs ermitteln (inkl. Unterabteilungen) und ausblenden
            var alreadyAssigned = new HashSet<string>(
                dbManager.GetAllIPsRecursive(targetLocationId)
                         .Select(ip => ip.IPAddress));

            // Alle gescannten Geräte aus Devices-Tabelle laden
            allDevices = dbManager.GetAllIPsFromDevices()
                .Where(d => !alreadyAssigned.Contains(d.IP))
                .OrderBy(d => d.IP)
                .Select(d => new DeviceListItem { IP = d.IP, Hostname = d.Hostname ?? "" })
                .ToList();

            // --- Header ---
            var headerLabel = new Label
            {
                Text = $"Ziel: {locationName}\nVerfügbare Geräte aus DB (bereits zugewiesene ausgeblendet):",
                Location = new Point(10, 10),
                Size = new Size(565, 40),
                Font = new Font("Segoe UI", 10)
            };

            // --- Suchbox ---
            var searchLabel = new Label { Text = "🔍 Suchen:", Location = new Point(10, 58), AutoSize = true };
            searchBox = new TextBox { Location = new Point(95, 55), Width = 475, Font = new Font("Segoe UI", 10) };
            searchBox.TextChanged += (s, e) => FilterList(searchBox.Text);

            // --- CheckedListBox ---
            checkedListBox = new CheckedListBox
            {
                Location = new Point(10, 88),
                Size = new Size(565, 390),
                CheckOnClick = true,
                Font = new Font("Consolas", 10)
            };
            PopulateList(allDevices);

            // --- Alle / Keine ---
            var selectAllBtn = new Button { Text = "Alle auswählen", Location = new Point(10, 488), Width = 130, BackColor = SystemColors.Control, Font = new Font("Segoe UI", 10) };
            selectAllBtn.Click += (s, e) => { for (int i = 0; i < checkedListBox.Items.Count; i++) checkedListBox.SetItemChecked(i, true); };

            var selectNoneBtn = new Button { Text = "Keine", Location = new Point(150, 488), Width = 80, Font = new Font("Segoe UI", 10) };
            selectNoneBtn.Click += (s, e) => { for (int i = 0; i < checkedListBox.Items.Count; i++) checkedListBox.SetItemChecked(i, false); };

            countLabel = new Label { Location = new Point(245, 491), AutoSize = true, ForeColor = Color.DarkBlue };
            UpdateCountLabel(allDevices.Count);

            // --- OK / Abbrechen ---
            var okBtn = new Button
            {
                Text = "✅ Importieren",
                Location = new Point(355, 538),
                Width = 120,
                Height = 30,
                BackColor = SystemColors.Control,
                DialogResult = DialogResult.OK,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            okBtn.Click += (s, e) =>
            {
                SelectedIPs.Clear();
                foreach (var item in checkedListBox.CheckedItems)
                {
                    if (item is DeviceListItem entry)
                        SelectedIPs.Add(Tuple.Create(entry.IP, entry.Hostname));
                }
            };

            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(485, 538),
                Width = 95,
                Height = 30,
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 10)
            };

            Controls.AddRange(new Control[] {
                headerLabel, searchLabel, searchBox,
                checkedListBox,
                selectAllBtn, selectNoneBtn, countLabel,
                okBtn, cancelBtn
            });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void PopulateList(List<DeviceListItem> devices)
        {
            checkedListBox.Items.Clear();
            foreach (var d in devices)
                checkedListBox.Items.Add(d);
            UpdateCountLabel(devices.Count);
        }

        private void FilterList(string filter)
        {
            var filtered = string.IsNullOrWhiteSpace(filter)
                ? allDevices
                : allDevices.Where(d =>
                    d.IP.Contains(filter) ||
                    d.Hostname.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToList();
            PopulateList(filtered);
        }

        private void UpdateCountLabel(int count)
        {
            if (countLabel != null)
                countLabel.Text = $"{count} Geräte verfügbar";
        }

        private class DeviceListItem
        {
            public string IP { get; set; }
            public string Hostname { get; set; }
            public override string ToString()
                => string.IsNullOrEmpty(Hostname) ? IP : $"{IP,-18}  {Hostname}";
        }
    }

    // =========================================================
    // === DATENBANK-AUSWAHL-DIALOG ===
    // =========================================================
    public class DatabaseSelectionDialog : Form
    {
        private readonly DatabaseManager _db;
        private CheckedListBox clb;
        public List<string> SelectedDatabasePaths { get; private set; } = new List<string>();

        public DatabaseSelectionDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Datenbanken auswählen";
            Size = new Size(520, 420);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Font = new Font("Segoe UI", 10);

            var header = new Label
            {
                Text = "Welche Datenbanken sollen in LibreOffice Calc geöffnet werden?",
                Location = new Point(12, 12),
                Size = new Size(480, 40),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            clb = new CheckedListBox
            {
                Location = new Point(12, 58),
                Size = new Size(480, 210),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 10)
            };

            LoadDatabases();

            var btnAlle = new Button { Text = "✔ Alle auswählen", Location = new Point(12, 276), Width = 155, Height = 28 };
            btnAlle.Click += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true); };

            var btnKeine = new Button { Text = "✖ Keine", Location = new Point(175, 276), Width = 100, Height = 28 };
            btnKeine.Click += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false); };

            var lblAnzahl = new Label
            {
                Location = new Point(12, 312),
                Size = new Size(300, 20),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DarkSlateGray,
                Text = $"0 von {clb.Items.Count} Datenbanken ausgewählt"
            };
            clb.ItemCheck += (s, e) =>
            {
                int cnt = clb.CheckedItems.Count + (e.NewValue == CheckState.Checked ? 1 : -1);
                lblAnzahl.Text = $"{cnt} von {clb.Items.Count} Datenbanken ausgewählt";
            };

            var btnOk = new Button
            {
                Text = "In Calc öffnen",
                Location = new Point(290, 340),
                Width = 140,
                Height = 30,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnOk.Click += (s, e) =>
            {
                SelectedDatabasePaths = clb.CheckedItems.Cast<DbItem>()
                    .Select(item => item.Path).ToList();
                if (SelectedDatabasePaths.Count == 0)
                {
                    MessageBox.Show("Bitte mindestens eine Datenbank auswählen.",
                        "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                }
            };

            var btnAbbrechen = new Button
            {
                Text = "Abbrechen",
                Location = new Point(440, 340),
                Width = 68,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };

            Controls.AddRange(new Control[] { header, clb, btnAlle, btnKeine, lblAnzahl, btnOk, btnAbbrechen });
            AcceptButton = btnOk;
            CancelButton = btnAbbrechen;
        }

        private void LoadDatabases()
        {
            clb.Items.Clear();
            try
            {
                string mainPath = _db.GetMainDatabasePath();
                var kunden      = _db.GetCustomers();
                var allFiles    = _db.GetAllDatabaseFiles();

                // Haupt-DB zuerst
                if (System.IO.File.Exists(mainPath))
                {
                    long kb = new System.IO.FileInfo(mainPath).Length / 1024;
                    clb.Items.Add(new DbItem { Path = mainPath, Display = $"Haupt-Datenbank  ({kb} KB)  —  nmap_inventory.db" }, true);
                }

                // Kunden-DBs
                foreach (var file in allFiles.OrderBy(f => f))
                {
                    if (file == mainPath) continue;
                    var custId = _db.TryGetCustomerIdFromPath(file);
                    string name = custId.HasValue
                        ? kunden.FirstOrDefault(k => k.ID == custId.Value)?.Name ?? $"Kunde {custId}"
                        : System.IO.Path.GetFileName(file);
                    long kb = new System.IO.FileInfo(file).Length / 1024;
                    clb.Items.Add(new DbItem { Path = file, Display = $"{name,-28}  ({kb} KB)  —  {System.IO.Path.GetFileName(file)}" }, true);
                }
            }
            catch { }

            if (clb.Items.Count == 0)
                clb.Items.Add(new DbItem { Path = "", Display = "Keine Datenbanken gefunden" }, false);
        }

        private class DbItem
        {
            public string Path { get; set; }
            public string Display { get; set; }
            public override string ToString() => Display;
        }
    }

    // =========================================================
    // === SESSION-SCHLÜSSEL ===
    // Hält das Benutzerpasswort für die laufende Session im RAM.
    // Wird nie auf Disk geschrieben.
    // =========================================================
    public static class SessionKey
    {
        private static string _key = null;

        public static string Current => _key;
        public static bool IsSet => !string.IsNullOrEmpty(_key);

        public static void Set(string password) => _key = password;
        public static void Clear() => _key = null;
    }

    // =========================================================
    // === LOGIN-DIALOG ===
    // Wird beim Programmstart angezeigt.
    // Passwort bleibt nur im RAM (SessionKey).
    // =========================================================
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

    // =========================================================
    // === CREDENTIAL-PASSWORT-DIALOG ===
    // Erscheint NUR wenn auf gespeicherte Zugangsdaten
    // zugegriffen wird — nicht beim Programmstart.
    // =========================================================
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

    // =========================================================
    // === SNMP-EINSTELLUNGS-DIALOG ===
    // =========================================================
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

    // =========================================================
    // === LINUX / SSH VERBINDUNGS-DIALOG ===
    // =========================================================
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

    // =========================================================
    // === ADB CONNECTION DIALOG ===
    // =========================================================
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

            var btnOk = new Button { Text = "Verbinden", Left = 225, Top = 268, Width = 90, DialogResult = DialogResult.OK };
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
            var btnCancel = new Button { Text = "Abbrechen", Left = 323, Top = 268, Width = 85, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private static Label Lbl(string text, int x, int y)
            => new Label { Text = text, Left = x, Top = y + 3, AutoSize = true };
    }

    // =========================================================
    // === CREDENTIAL TEMPLATE EDIT DIALOG ===
    // =========================================================
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
            Font = new System.Drawing.Font("Segoe UI", 9);

            int y = 14;

            // Name
            Controls.Add(new Label { Text = "Name:", Location = new System.Drawing.Point(14, y), AutoSize = true });
            txtName = new TextBox { Location = new System.Drawing.Point(140, y - 2), Width = 280 };
            Controls.Add(txtName);
            y += 32;

            // Username
            Controls.Add(new Label { Text = "Benutzername:", Location = new System.Drawing.Point(14, y), AutoSize = true });
            txtUsername = new TextBox { Location = new System.Drawing.Point(140, y - 2), Width = 280 };
            Controls.Add(txtUsername);
            y += 32;

            // Password + eye toggle
            Controls.Add(new Label { Text = "Passwort:", Location = new System.Drawing.Point(14, y), AutoSize = true });
            txtPassword = new TextBox
            {
                Location = new System.Drawing.Point(140, y - 2), Width = 244,
                UseSystemPasswordChar = true
            };
            Controls.Add(txtPassword);

            var btnEye = new Button
            {
                Text = "\ud83d\udc41", Location = new System.Drawing.Point(388, y - 3),
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
                Location = new System.Drawing.Point(240, y), Width = 90, Height = 30
            };
            var btnCancel = new Button
            {
                Text = "Abbrechen", DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(336, y), Width = 90, Height = 30
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

    // =========================================================
    // === CREDENTIAL TEMPLATE LIST DIALOG ===
    // =========================================================
    public class CredentialTemplateDialog : Form
    {
        private readonly DatabaseManager _db;
        private DataGridView grid;

        public CredentialTemplateDialog(DatabaseManager db)
        {
            _db = db;
            Text = "Passwort-Vorlagen verwalten";
            Width = 820; Height = 480;
            StartPosition = FormStartPosition.CenterParent;
            Font = new System.Drawing.Font("Segoe UI", 9);
            MinimizeBox = false;

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                MultiSelect = false
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 245, 245);
            grid.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);

            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID", HeaderText = "ID", Width = 40, Visible = false });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Name", FillWeight = 35 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Username", HeaderText = "Benutzer", FillWeight = 30 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Password", HeaderText = "Passwort", FillWeight = 20 });

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            var btnAdd = new Button { Text = "Hinzufügen", Location = new System.Drawing.Point(6, 8), Width = 110, Height = 28 };
            var btnEdit = new Button { Text = "Bearbeiten", Location = new System.Drawing.Point(122, 8), Width = 110, Height = 28 };
            var btnDelete = new Button { Text = "Löschen", Location = new System.Drawing.Point(238, 8), Width = 110, Height = 28, BackColor = System.Drawing.Color.FromArgb(255, 220, 220) };
            var btnClose = new Button { Text = "Schließen", Location = new System.Drawing.Point(694, 8), Width = 100, Height = 28, DialogResult = DialogResult.Cancel };

            btnAdd.Click += (s, e) => AddTemplate();
            btnEdit.Click += (s, e) => EditTemplate();
            btnDelete.Click += (s, e) => DeleteTemplate();
            grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) EditTemplate(); };

            btnPanel.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDelete, btnClose });
            Controls.Add(grid);
            Controls.Add(btnPanel);
            CancelButton = btnClose;

            EnsureMasterPassword();
            LoadTemplates();
        }

        private void EnsureMasterPassword()
        {
            if (!SessionKey.IsSet)
            {
                using (var dlg = new CredentialPasswordDialog())
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK)
                    {
                        Close();
                        return;
                    }
                }
            }
        }

        private void LoadTemplates()
        {
            grid.Rows.Clear();
            foreach (var t in _db.GetCredentialTemplates())
                grid.Rows.Add(t.ID, t.Name, t.Username, "••••••");
        }

        private void AddTemplate()
        {
            using (var dlg = new CredentialTemplateEditDialog())
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Template != null)
                {
                    _db.AddCredentialTemplate(dlg.Template);
                    LoadTemplates();
                }
            }
        }

        private void EditTemplate()
        {
            if (grid.CurrentRow == null) return;
            int id = Convert.ToInt32(grid.CurrentRow.Cells["ID"].Value);
            var template = _db.GetCredentialTemplates().FirstOrDefault(t => t.ID == id);
            if (template == null) return;

            using (var dlg = new CredentialTemplateEditDialog(template))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && dlg.Template != null)
                {
                    _db.UpdateCredentialTemplate(dlg.Template);
                    LoadTemplates();
                }
            }
        }

        private void DeleteTemplate()
        {
            if (grid.CurrentRow == null) return;
            string name = grid.CurrentRow.Cells["Name"].Value?.ToString();
            if (MessageBox.Show($"Vorlage \"{name}\" wirklich löschen?\n\nZugewiesene Geräte werden zurückgesetzt.",
                "Löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            int id = Convert.ToInt32(grid.CurrentRow.Cells["ID"].Value);
            _db.DeleteCredentialTemplate(id);
            LoadTemplates();
        }
    }

    // =========================================================
    // === CREDENTIAL SCAN PROGRESS DIALOG ===
    // =========================================================
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
            Font = new System.Drawing.Font("Segoe UI", 9);

            lblStatus = new Label
            {
                Text = $"{_devices.Count} Geräte bereit zum Testen",
                Dock = DockStyle.Top, Height = 28,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
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
                Font = new System.Drawing.Font("Consolas", 9),
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                ForeColor = System.Drawing.Color.White
            };

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            btnStart = new Button
            {
                Text = "🔑 Test starten", Location = new System.Drawing.Point(6, 8),
                Width = 140, Height = 28,
                BackColor = System.Drawing.Color.FromArgb(220, 240, 220),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
            };
            btnCancel = new Button
            {
                Text = "Schließen", Location = new System.Drawing.Point(574, 8),
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
                            System.Drawing.Color.FromArgb(255, 200, 100));
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
                AppendLog($"✔ {result.Device.IP} ({host}) — {result.Message}\n", System.Drawing.Color.LightGreen);
            }
            else
            {
                AppendLog($"✘ {result.Device.IP} ({host}) — {result.Message}\n", System.Drawing.Color.FromArgb(255, 150, 150));
            }
        }

        private void AppendLog(string text, System.Drawing.Color color)
        {
            logBox.SelectionStart = logBox.TextLength;
            logBox.SelectionLength = 0;
            logBox.SelectionColor = color;
            logBox.AppendText(text);
            logBox.ScrollToCaret();
        }
    }

}