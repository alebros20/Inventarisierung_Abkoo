using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public partial class MainForm
    {
        private TabPage CreateDBDeviceTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var filter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            filter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            filter.SelectedIndex = 0;
            filter.SelectedIndexChanged += (s, e) => LoadDatabaseDevices(filter.SelectedItem.ToString());
            var saveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            saveBtn.Click += (s, e) => SaveDatabaseDevices();
            var deleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            deleteBtn.Click += (s, e) => DeleteDatabaseDeviceRow();
            panel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, filter, saveBtn, deleteBtn });

            dbDeviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP", Width = 120 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC", Width = 150 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 80 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 350 });
            dbDeviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Gerätetyp", Width = 140, ReadOnly = true });
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(dbDeviceTable);
            container.Controls.Add(panel);
            return new TabPage("DB - Geräte") { Controls = { container } };
        }

        private TabPage CreateDBSoftwareTab()
        {
            Panel panel = new Panel { Dock = DockStyle.Top, Height = 50 };
            var filter = new ComboBox { Location = new Point(80, 12), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10) };
            filter.Items.AddRange(new[] { "Alle", "Heute", "Diese Woche", "Dieser Monat", "Dieses Jahr" });
            filter.SelectedIndex = 0;
            filter.SelectedIndexChanged += (s, e) => LoadDatabaseSoftware(filter.SelectedItem.ToString());
            var saveBtn = new Button { Text = "Speichern", Location = new Point(250, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            saveBtn.Click += (s, e) => SaveDatabaseSoftware();
            var deleteBtn = new Button { Text = "Zeile löschen", Location = new Point(360, 12), Width = 100, Font = new Font("Segoe UI", 10) };
            deleteBtn.Click += (s, e) => DeleteDatabaseSoftwareRow();
            panel.Controls.AddRange(new Control[] { new Label { Text = "Zeitraum:", Location = new Point(10, 15), AutoSize = true, Font = new Font("Segoe UI", 10) }, filter, saveBtn, deleteBtn });

            dbSoftwareTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = false, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10) };
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "ID", Width = 50, ReadOnly = true, Visible = false });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "GeräteID", Width = 80, ReadOnly = true });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Zeitstempel", Width = 150, ReadOnly = true });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "PC Name/IP", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Name", Width = 200 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Version", Width = 100 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller", Width = 150 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Installiert am", Width = 120 });
            dbSoftwareTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Letztes Update", Width = 140 });
            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(dbSoftwareTable);
            container.Controls.Add(panel);
            return new TabPage("DB - Software") { Controls = { container } };
        }

        private void RefreshAllViews()
        {
            refreshButton.Enabled = false;
            refreshButton.Text = "⏳ Lädt...";
            try
            {
                // Inventar_* sind jetzt VIEWs — kein Refresh nötig
                LoadDatabaseDevices();
                LoadDatabaseSoftware();
                LoadCustomerTree();
                statusLabel.Text = "✔ Daten aktualisiert — " + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Fehler beim Aktualisieren: " + ex.Message;
            }
            finally
            {
                refreshButton.Enabled = true;
                refreshButton.Text = "🔄 Aktualisieren";
            }
        }

        private void LoadDatabaseDevices(string filter = "Alle")
        {
            dbDeviceTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var dev in dbManager.LoadDevices(filter))
                if (displayed.Add(dev.IP))
                    dbDeviceTable.Rows.Add(dev.ID, dev.Zeitstempel, dev.IP, dev.Hostname, dev.MacAddress ?? "", dev.Status, dev.Ports,
                        $"{DeviceTypeHelper.GetIcon(dev.DeviceType)} {DeviceTypeHelper.GetLabel(dev.DeviceType)}");
        }

        private void LoadDatabaseSoftware(string filter = "Alle")
        {
            dbSoftwareTable.Rows.Clear();
            var displayed = new HashSet<string>();
            foreach (var sw in dbManager.LoadSoftware(filter))
            {
                string key = sw.PCName + "|" + sw.Name + "|" + sw.Version;
                if (displayed.Add(key))
                    dbSoftwareTable.Rows.Add(sw.ID, sw.DeviceID > 0 ? (object)sw.DeviceID : "", sw.Zeitstempel, sw.PCName, sw.Name, sw.Version, sw.Publisher, sw.InstallDate, sw.LastUpdate);
            }
        }

        private void SaveDatabaseDevices()
        {
            try { dbManager.UpdateDevices(GetDevicesFromGrid()); MessageBox.Show("Gespeichert!", "Erfolg"); LoadDatabaseDevices(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void SaveDatabaseSoftware()
        {
            try { dbManager.UpdateSoftwareEntries(GetSoftwareFromGrid()); MessageBox.Show("Gespeichert!", "Erfolg"); LoadDatabaseSoftware(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void DeleteDatabaseDeviceRow()
        {
            if (dbDeviceTable.SelectedRows.Count == 0) return;
            try { dbManager.DeleteDevice(Convert.ToInt32(dbDeviceTable.SelectedRows[0].Cells[0].Value)); LoadDatabaseDevices(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void DeleteDatabaseSoftwareRow()
        {
            if (dbSoftwareTable.SelectedRows.Count == 0) return;
            try { dbManager.DeleteSoftwareEntry(Convert.ToInt32(dbSoftwareTable.SelectedRows[0].Cells[0].Value)); LoadDatabaseSoftware(); }
            catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private void ExportData()
        {
            using (var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json|XML (*.xml)|*.xml" })
                if (sfd.ShowDialog() == DialogResult.OK)
                    try { new DataExporter().Export(sfd.FileName, currentDevices, new List<SoftwareInfo>(), ""); MessageBox.Show("Exportiert!", "Erfolg"); }
                    catch (Exception ex) { MessageBox.Show($"Fehler: {ex.Message}"); }
        }

        private List<DatabaseDevice> GetDevicesFromGrid()
        {
            var list = new List<DatabaseDevice>();
            foreach (DataGridViewRow row in dbDeviceTable.Rows)
                if (row.Cells[0].Value != null)
                    list.Add(new DatabaseDevice { ID = Convert.ToInt32(row.Cells[0].Value), Zeitstempel = row.Cells[1].Value?.ToString(), IP = row.Cells[2].Value?.ToString(), Hostname = row.Cells[3].Value?.ToString(), MacAddress = row.Cells[4].Value?.ToString(), Status = row.Cells[5].Value?.ToString(), Ports = row.Cells[6].Value?.ToString() });
            return list;
        }

        private List<DatabaseSoftware> GetSoftwareFromGrid()
        {
            var list = new List<DatabaseSoftware>();
            foreach (DataGridViewRow row in dbSoftwareTable.Rows)
                if (row.Cells[0].Value != null)
                    list.Add(new DatabaseSoftware { ID = Convert.ToInt32(row.Cells[0].Value), DeviceID = row.Cells[1].Value is int did ? did : 0, Zeitstempel = row.Cells[2].Value?.ToString(), PCName = row.Cells[3].Value?.ToString(), Name = row.Cells[4].Value?.ToString(), Version = row.Cells[5].Value?.ToString(), Publisher = row.Cells[6].Value?.ToString(), InstallDate = row.Cells[7].Value?.ToString(), LastUpdate = row.Cells[8].Value?.ToString() });
            return list;
        }
    }
}
