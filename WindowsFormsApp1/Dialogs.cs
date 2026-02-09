using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Newtonsoft.Json;
using JsonFormatting = Newtonsoft.Json.Formatting;

namespace NmapInventory
{
    public class DataExporter
    {
        public void Export(string filePath, List<DeviceInfo> devices, List<SoftwareInfo> software, string hardware)
        {
            if (filePath.EndsWith(".json"))
            {
                var data = new { Zeitstempel = DateTime.Now, Geräte = devices, Software = software, Hardware = hardware };
                File.WriteAllText(filePath, JsonConvert.SerializeObject(data, JsonFormatting.Indented));
            }
            else
            {
                var doc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("Inventar",
                        new XElement("Zeitstempel", DateTime.Now),
                        new XElement("Geräte", devices.Select(d => new XElement("Gerät", 
                            new XElement("IP", d.IP),
                            new XElement("Hostname", d.Hostname ?? "")))),
                        new XElement("Hardware", new XCData(hardware))));
                doc.Save(filePath);
            }
        }
    }

    public class RemoteConnectionForm : Form
    {
        public string ComputerIP { get; private set; }
        public string Username { get; private set; }
        public string Password { get; private set; }

        public RemoteConnectionForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            Text = "Remote Verbindung";
            Width = 500;
            Height = 320;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var ipTb = new TextBox { Location = new Point(150, 30), Width = 300 };
            var userTb = new TextBox { Location = new Point(150, 70), Width = 300, Text = Environment.UserName };
            var passTb = new TextBox { Location = new Point(150, 110), Width = 300, UseSystemPasswordChar = true };
            
            var okBtn = new Button { Text = "Verbinden", Location = new Point(150, 210), Width = 100, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 210), Width = 100, DialogResult = DialogResult.Cancel };

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
        }
    }

    public class PasswordVerificationForm : Form
    {
        private const string REQUIRED_PASSWORD = "Administrator";

        public PasswordVerificationForm(string action)
        {
            InitializeUI(action);
        }

        private void InitializeUI(string action)
        {
            Text = "Sicherheitsabfrage";
            Width = 450;
            Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var passwordTextBox = new TextBox { Location = new Point(20, 90), Width = 400, Font = new Font("Segoe UI", 11), UseSystemPasswordChar = true };

            var okBtn = new Button { Text = "Bestätigen", Location = new Point(150, 150), Width = 100, BackColor = Color.LightGreen };
            okBtn.Click += (s, e) => VerifyPassword(passwordTextBox, action);
            
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(260, 150), Width = 100, DialogResult = DialogResult.Cancel, BackColor = Color.LightCoral };

            Controls.AddRange(new Control[] { 
                new Label { Text = $"WARNUNG: Du bist dabei '{action}' auszuführen!\n\nGib das Sicherheitspasswort ein:", Location = new Point(20, 20), Width = 400, Height = 60, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.DarkRed },
                passwordTextBox,
                new Label { Text = $"Erforderliches Passwort: \"{REQUIRED_PASSWORD}\"", Location = new Point(20, 120), Width = 400, ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Italic) },
                okBtn, cancelBtn
            });

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void VerifyPassword(TextBox passwordTextBox, string action)
        {
            if (passwordTextBox.Text == REQUIRED_PASSWORD)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show($"Falsches Passwort!\n\nDas korrekte Passwort lautet: \"{REQUIRED_PASSWORD}\"", "Zugriff verweigert", MessageBoxButtons.OK, MessageBoxIcon.Error);
                passwordTextBox.Clear();
                passwordTextBox.Focus();
            }
        }
    }
    
    public class InputDialog : Form
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        
        public InputDialog(string title, string label1, string label2)
        {
            InitializeUI(title, label1, label2);
        }

        private void InitializeUI(string title, string label1, string label2)
        {
            Text = title;
            Width = 450;
            Height = 220;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            
            var txt1 = new TextBox { Location = new Point(20, 45), Width = 400 };
            var txt2 = new TextBox { Location = new Point(20, 105), Width = 400, Multiline = true, Height = 50 };
            
            var okBtn = new Button { Text = "OK", Location = new Point(220, 165), Width = 90, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(320, 165), Width = 100, DialogResult = DialogResult.Cancel };
            
            okBtn.Click += (s, e) => { Value1 = txt1.Text; Value2 = txt2.Text; };
            
            this.Load += (s, e) => 
            {
                txt1.Text = Value1 ?? "";
                txt2.Text = Value2 ?? "";
            };
            
            Controls.AddRange(new Control[] { 
                new Label { Text = label1, Location = new Point(20, 20), AutoSize = true },
                txt1,
                new Label { Text = label2, Location = new Point(20, 80), AutoSize = true },
                txt2,
                okBtn, cancelBtn
            });
            
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }
    }
    
    public class IPImportDialog : Form
    {
        public List<(string IP, string Workstation)> SelectedIPs { get; private set; } = new List<(string, string)>();
        
        public IPImportDialog(DatabaseManager dbManager)
        {
            InitializeUI(dbManager);
        }

        private void InitializeUI(DatabaseManager dbManager)
        {
            Text = "IPs aus Datenbank importieren";
            Width = 600;
            Height = 500;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            
            var ipGrid = new DataGridView 
            { 
                Location = new Point(20, 50), 
                Width = 550, 
                Height = 350,
                AllowUserToAddRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true
            };
            
            ipGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "✓", Width = 30 });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP-Adresse", Width = 150, ReadOnly = true });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname (DB)", Width = 150, ReadOnly = true });
            ipGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Als Arbeitsplatz", Width = 200 });
            
            var ips = dbManager.GetAllIPsFromDevices();
            foreach (var ip in ips)
                ipGrid.Rows.Add(false, ip.IP, ip.Hostname, ip.Hostname);
            
            var selectAllBtn = new Button { Text = "Alle auswählen", Location = new Point(20, 410), Width = 120 };
            selectAllBtn.Click += (s, e) => 
            {
                foreach (DataGridViewRow row in ipGrid.Rows)
                    row.Cells[0].Value = true;
            };
            
            var deselectAllBtn = new Button { Text = "Alle abwählen", Location = new Point(150, 410), Width = 120 };
            deselectAllBtn.Click += (s, e) => 
            {
                foreach (DataGridViewRow row in ipGrid.Rows)
                    row.Cells[0].Value = false;
            };
            
            var okBtn = new Button { Text = "Importieren", Location = new Point(350, 410), Width = 100, BackColor = Color.LightGreen };
            okBtn.Click += (s, e) => OnOkClick(ipGrid);
            
            var cancelBtn = new Button { Text = "Abbrechen", Location = new Point(460, 410), Width = 100, DialogResult = DialogResult.Cancel };
            
            Controls.AddRange(new Control[] { 
                new Label { Text = "Wähle eine oder mehrere IP-Adressen aus der Geräte-Datenbank:", Location = new Point(20, 20), Width = 550, Font = new Font("Segoe UI", 10, FontStyle.Bold) },
                ipGrid, selectAllBtn, deselectAllBtn, okBtn, cancelBtn
            });
            
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void OnOkClick(DataGridView ipGrid)
        {
            SelectedIPs.Clear();
            foreach (DataGridViewRow row in ipGrid.Rows)
            {
                if (row.Cells[0].Value != null && (bool)row.Cells[0].Value)
                {
                    string ip = row.Cells[1].Value?.ToString();
                    string workstation = row.Cells[3].Value?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(ip))
                        SelectedIPs.Add((ip, workstation));
                }
            }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
