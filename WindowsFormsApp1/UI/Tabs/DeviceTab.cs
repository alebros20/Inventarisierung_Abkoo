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
        private TabPage CreateDeviceTab()
        {
            Panel devicePanel = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = SystemColors.Control, Padding = new Padding(10) };
            devicePanel.Controls.AddRange(new Control[] {
                new Label { Text = "Legende:", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) },
                new Label { Text = "●", Location = new Point(90, 10), Font = new Font("Arial", 16), ForeColor = Color.Green, AutoSize = true },
                new Label { Text = "Aktiv (online)", Location = new Point(110, 12), AutoSize = true, Font = new Font("Segoe UI", 10) },
                new Label { Text = "●", Location = new Point(230, 10), Font = new Font("Arial", 16), ForeColor = Color.Gold, AutoSize = true },
                new Label { Text = "Neu seit letztem Scan", Location = new Point(250, 12), AutoSize = true, Font = new Font("Segoe UI", 10) },
                new Label { Text = "●", Location = new Point(430, 10), Font = new Font("Arial", 16), ForeColor = Color.Red, AutoSize = true },
                new Label { Text = "Offline (fehlend)", Location = new Point(450, 12), AutoSize = true, Font = new Font("Segoe UI", 10) }
            });

            deviceTable = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ScrollBars = ScrollBars.Both, Font = new Font("Segoe UI", 10), RowTemplate = { Height = 35 } };
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status",        Width = 50  });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "IP",            Width = 120 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hostname",      Width = 150 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Gerätetyp",    Width = 130 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Hersteller",   Width = 130 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC-Adresse",  Width = 140 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Kommentar",    Width = 200 });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Online Status", Width = 80  });
            deviceTable.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Offene Ports", Width = 400 });
            // Nur Kommentar-Spalte (Index 6) editierbar — alle anderen per CellBeginEdit sperren
            deviceTable.CellBeginEdit += (s, e) =>
            {
                if (e.ColumnIndex != 6) e.Cancel = true;
            };
            deviceTable.CellPainting += (s, e) =>
            {
                if (e.ColumnIndex == 0 && e.RowIndex >= 0)
                {
                    e.PaintBackground(e.ClipBounds, false);
                    string status = e.Value?.ToString() ?? "";
                    Color circleColor = status == "NEU" ? Color.Gold : status == "OFFLINE" ? Color.Red : Color.Green;
                    using (var brush = new SolidBrush(circleColor))
                        e.Graphics.FillEllipse(brush, e.CellBounds.Left + 15, e.CellBounds.Top + 10, 20, 20);
                    e.Handled = true;
                }
            };

            // Kommentar-Spalte: Änderung direkt in DB speichern (Spalte 6)
            deviceTable.CellEndEdit += (s, e) =>
            {
                if (e.ColumnIndex != 6 || e.RowIndex < 0) return;
                string ip      = deviceTable.Rows[e.RowIndex].Cells[1].Value?.ToString();
                string comment = deviceTable.Rows[e.RowIndex].Cells[6].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(ip))
                    dbManager.SaveComment(ip, comment);
            };

            // Rechtsklick auf Zeile → Gerätetyp manuell setzen
            var ctxDeviceTable = new ContextMenuStrip();
            deviceTable.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Right) return;
                var hit = deviceTable.HitTest(e.X, e.Y);
                if (hit.RowIndex < 0) return;
                deviceTable.Rows[hit.RowIndex].Selected = true;

                ctxDeviceTable.Items.Clear();
                string rowIp = deviceTable.Rows[hit.RowIndex].Cells[1].Value?.ToString();
                if (string.IsNullOrEmpty(rowIp)) return;

                ctxDeviceTable.Items.Add(new ToolStripLabel($"Gerätetyp setzen für {rowIp}") { Font = new Font("Segoe UI", 9, FontStyle.Bold), Enabled = false });
                ctxDeviceTable.Items.Add(new ToolStripSeparator());

                void AddType(string label, DeviceType dt)
                {
                    ctxDeviceTable.Items.Add(label, null, (_, __) =>
                    {
                        dbManager.SetDeviceType(rowIp, dt);
                        LoadCustomerTree();
                        // Zeile direkt aktualisieren
                        string icon = $"{DeviceTypeHelper.GetIcon(dt)} {DeviceTypeHelper.GetLabel(dt)}";
                        deviceTable.Rows[hit.RowIndex].Cells[3].Value = icon;
                    });
                }

                AddType("🖥  Windows PC",      DeviceType.WindowsPC);
                AddType("🗄  Windows Server",   DeviceType.WindowsServer);
                AddType("💻 Laptop",            DeviceType.Laptop);
                AddType("🐧 Linux",             DeviceType.Linux);
                AddType("🍎 macOS / iOS",       DeviceType.MacOS);
                AddType("📱 Smartphone",        DeviceType.Smartphone);
                AddType("📟 Tablet",            DeviceType.Tablet);
                AddType("🔀 Netzwerkgerät",     DeviceType.NetzwerkGeraet);
                AddType("💾 NAS",               DeviceType.NAS);
                AddType("📺 Smart TV",          DeviceType.SmartTV);
                AddType("💡 IoT-Gerät",         DeviceType.IoT);
                AddType("🖨  Drucker",           DeviceType.Drucker);
                AddType("❓ Unbekannt",          DeviceType.Unbekannt);

                ctxDeviceTable.Show(deviceTable, e.Location);
            };

            var container = new Panel { Dock = DockStyle.Fill };
            container.Controls.Add(deviceTable);
            container.Controls.Add(devicePanel);
            return new TabPage("Geräte") { Controls = { container } };
        }

        private TabPage CreateNmapTab()
        {
            rawOutputTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10), ReadOnly = true };
            return new TabPage("Nmap Ausgabe") { Controls = { rawOutputTextBox } };
        }
    }
}
