using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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

        public RemoteConnectionForm()
        {
            Text = "Remote Verbindung";
            Width = 500; Height = 320;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var ipTb = new TextBox { Location = new Point(150, 30), Width = 300 };
            var userTb = new TextBox { Location = new Point(150, 70), Width = 300, Text = Environment.UserName };
            var passTb = new TextBox { Location = new Point(150, 110), Width = 300, UseSystemPasswordChar = true };

            var okBtn = new Button { Text = "Verbinden", Location = new Point(150, 210), Width = 100, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 210), Width = 100, DialogResult = DialogResult.Cancel };

            okBtn.Click += (s, e) =>
            {
                ComputerIP = ipTb.Text.Trim();
                Username = userTb.Text.Trim();
                Password = passTb.Text;
                if (string.IsNullOrEmpty(ComputerIP))
                {
                    MessageBox.Show("Bitte gib eine Computer-IP ein!", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            };

            Controls.AddRange(new Control[] {
                new Label { Text = "Computer-IP:", Location = new Point(20, 33), AutoSize = true },
                ipTb,
                new Label { Text = "Benutzername:", Location = new Point(20, 73), AutoSize = true },
                userTb,
                new Label { Text = "Passwort:", Location = new Point(20, 113), AutoSize = true },
                passTb,
                new Label { Text = "Hinweis: Für Remote-Zugriff wird ein Passwort benötigt.\nBei leerem Passwort wird nur WMI Registry verwendet.", Location = new Point(20, 150), Width = 450, Height = 50, ForeColor = Color.DarkBlue },
                okBtn, cancelBtn
            });
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
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

            var okBtn = new Button { Text = "OK", Location = new Point(150, 120), Width = 100, DialogResult = DialogResult.OK, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
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
            var selectAllBtn = new Button { Text = "Alle auswählen", Location = new Point(10, 488), Width = 130, BackColor = Color.LightGreen, Font = new Font("Segoe UI", 10) };
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
                BackColor = Color.LightGreen,
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
}