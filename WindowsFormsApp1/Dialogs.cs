using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        private TextBox ipTb, userTb, passTb;
        private CheckBox chkSave;
        private ComboBox cmbSaved;

        public RemoteConnectionForm()
        {
            Text = "Remote Verbindung";
            Width = 520; Height = 370;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new System.Drawing.Font("Segoe UI", 9);

            // Gespeicherte Zugangsdaten laden
            var saved = CredentialStore.GetAll();

            // ── Gespeicherte Auswahl ──────────────────────────
            var lblSaved = new Label { Text = "Gespeichert:", Location = new Point(20, 14), AutoSize = true };
            cmbSaved = new ComboBox
            {
                Location = new Point(110, 11),
                Width = 270,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSaved.Items.Add("-- Neu eingeben --");
            foreach (var c in saved)
                cmbSaved.Items.Add(c);
            cmbSaved.SelectedIndex = 0;

            var btnDelete = new Button
            {
                Text = "🗑",
                Location = new Point(388, 10),
                Width = 30,
                Height = 24,
                Font = new System.Drawing.Font("Segoe UI", 9)
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
                Location = new Point(110, 52),
                Width = 310,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            userTb = new TextBox
            {
                Location = new Point(110, 84),
                Width = 310,
                Text = Environment.UserName,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            passTb = new TextBox
            {
                Location = new Point(110, 116),
                Width = 310,
                UseSystemPasswordChar = true,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };

            // ── Speichern-Checkbox ────────────────────────────
            chkSave = new CheckBox
            {
                Text = "Zugangsdaten verschlüsselt speichern",
                Location = new Point(110, 150),
                AutoSize = true,
                Checked = true
            };

            var lblAlias = new Label { Text = "Bezeichnung:", Location = new Point(110, 178), AutoSize = true };
            var aliasTb = new TextBox
            {
                Name = "aliasTb",
                Location = new Point(200, 175),
                Width = 220,
                Text = "z.B. Büro-PC",
                ForeColor = Color.Gray
            };
            aliasTb.GotFocus += (s, e) => { if (aliasTb.ForeColor == Color.Gray) { aliasTb.Text = ""; aliasTb.ForeColor = SystemColors.WindowText; } };
            aliasTb.LostFocus += (s, e) => { if (string.IsNullOrEmpty(aliasTb.Text)) { aliasTb.Text = "z.B. Büro-PC"; aliasTb.ForeColor = Color.Gray; } };
            chkSave.CheckedChanged += (s, e) => { lblAlias.Visible = chkSave.Checked; aliasTb.Visible = chkSave.Checked; };

            var hinweis = new Label
            {
                Text = "Verschlüsselung via Windows DPAPI — nur auf diesem PC entschlüsselbar.",
                Location = new Point(20, 210),
                Size = new Size(470, 16),
                Font = new System.Drawing.Font("Segoe UI", 8),
                ForeColor = System.Drawing.Color.DarkSlateGray
            };

            // ── Buttons ───────────────────────────────────────
            var okBtn = new Button
            {
                Text = "Verbinden",
                Location = new Point(110, 235),
                Width = 120,
                Height = 28,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            var cancelBtn = new Button
            {
                Text = "Abbrechen",
                Location = new Point(240, 235),
                Width = 100,
                Height = 28,
                Font = new System.Drawing.Font("Segoe UI", 10),
                DialogResult = DialogResult.Cancel
            };

            // Gespeicherten Eintrag laden wenn ausgewählt
            cmbSaved.SelectedIndexChanged += (s, e) =>
            {
                if (cmbSaved.SelectedIndex <= 0) return;
                var sel = cmbSaved.SelectedItem as CredentialEntry;
                if (sel == null) return;
                ipTb.Text = sel.IP;
                userTb.Text = sel.Username;
                passTb.Text = CredentialStore.Decrypt(sel.EncryptedPassword);
                chkSave.Checked = false; // nicht automatisch nochmal speichern
            };

            // Gespeicherten Eintrag löschen
            btnDelete.Click += (s, e) =>
            {
                if (cmbSaved.SelectedIndex <= 0) return;
                var sel = cmbSaved.SelectedItem as CredentialEntry;
                if (sel == null) return;
                if (MessageBox.Show($"Eintrag '{sel.Alias}' löschen?", "Löschen",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    CredentialStore.Delete(sel.Alias);
                    cmbSaved.Items.Remove(sel);
                    cmbSaved.SelectedIndex = 0;
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
                    return;
                }

                // Zugangsdaten speichern wenn gewünscht
                if (chkSave.Checked && !string.IsNullOrEmpty(Password))
                {
                    string alias = aliasTb.Text.Trim();
                    if (string.IsNullOrEmpty(alias))
                        alias = $"{Username}@{ComputerIP}";
                    CredentialStore.Save(new CredentialEntry
                    {
                        Alias = alias,
                        IP = ComputerIP,
                        Username = Username,
                        EncryptedPassword = CredentialStore.Encrypt(Password)
                    });
                }
            };

            Controls.AddRange(new Control[] {
                lblSaved, cmbSaved, btnDelete, sep,
                new Label { Text = "Computer-IP:",  Location = new Point(20, 55),  AutoSize = true },
                ipTb,
                new Label { Text = "Benutzername:", Location = new Point(20, 87),  AutoSize = true },
                userTb,
                new Label { Text = "Passwort:",     Location = new Point(20, 119), AutoSize = true },
                passTb,
                chkSave, lblAlias, aliasTb, hinweis,
                okBtn, cancelBtn
            });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        public void SetIP(string ip)
        {
            if (ipTb != null) ipTb.Text = ip;

            // Passenden gespeicherten Eintrag automatisch vorwählen
            for (int i = 1; i < cmbSaved.Items.Count; i++)
            {
                if (cmbSaved.Items[i] is CredentialEntry e && e.IP == ip)
                {
                    cmbSaved.SelectedIndex = i;
                    return;
                }
            }
        }
    }

    // =========================================================
    // === CREDENTIAL STORE (Windows DPAPI) ===
    // Speichert Zugangsdaten verschlüsselt mit Windows-Benutzerkonto.
    // Nur auf dem gleichen PC + Benutzer entschlüsselbar.
    // =========================================================
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

        public static string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            var bytes = Encoding.UTF8.GetBytes(plainText);
            var encrypted = System.Security.Cryptography.ProtectedData.Protect(bytes, null, System.Security.Cryptography.DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(encrypted);
        }

        public static string Decrypt(string encryptedBase64)
        {
            if (string.IsNullOrEmpty(encryptedBase64)) return "";
            try
            {
                var bytes = Convert.FromBase64String(encryptedBase64);
                var decrypted = System.Security.Cryptography.ProtectedData.Unprotect(bytes, null, System.Security.Cryptography.DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch { return ""; }
        }

        public static List<CredentialEntry> GetAll()
        {
            var list = new List<CredentialEntry>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                foreach (var line in File.ReadAllLines(FilePath))
                {
                    var parts = line.Split('	');
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

        public static void Save(CredentialEntry entry)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
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
            string mainPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "nmap_inventory.db");
            if (System.IO.File.Exists(mainPath))
                clb.Items.Add(new DbItem { Path = mainPath, Display = "Haupt-Datenbank  (nmap_inventory.db)" }, true);

            try
            {
                var kunden = _db.GetCustomers();
                var allFiles = _db.GetAllDatabaseFiles();
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
}