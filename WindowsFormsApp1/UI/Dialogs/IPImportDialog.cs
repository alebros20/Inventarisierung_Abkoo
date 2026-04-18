using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
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
            var selectAllBtn = new Button { Text = "Alle auswählen", Location = new Point(10, 488), Width = 130, Height = 28, Font = new Font("Segoe UI", 10) };
            selectAllBtn.Click += (s, e) => { for (int i = 0; i < checkedListBox.Items.Count; i++) checkedListBox.SetItemChecked(i, true); };

            var selectNoneBtn = new Button { Text = "Keine", Location = new Point(150, 488), Width = 80, Height = 28, Font = new Font("Segoe UI", 10) };
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
