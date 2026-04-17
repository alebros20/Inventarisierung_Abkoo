using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace NmapInventory
{
    public partial class MainForm
    {
        // =========================================================
        // === SCAN ===
        // =========================================================

        /// <summary>
        /// 🔍 NETZWERK SCANNEN — ARP-Discovery, schnell (~5 Sek)
        /// Findet ALLE Geräte im Subnetz inkl. MAC-Adressen.
        /// Befehl: nmap -sn -PR --send-eth
        /// </summary>
        private void StartScan()
        {
            if (selectedLocationID <= 0) { MessageBox.Show("Bitte wähle einen Standort aus!", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            scanButton.Enabled = false;
            remoteHardwareButton.Enabled = false;
            BeginScanUI("Netzwerk-Discovery läuft (ARP-Scan)...");

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var oldDevices = currentDevices.ToList();
                    lastScanDevices = oldDevices;
                    var result = nmapScanner.DiscoveryScan(networkTextBox.Text);
                    Invoke(new MethodInvoker(() =>
                    {
                        rawOutputTextBox.Text = result.RawOutput;
                        currentDevices = result.Devices;
                        dbManager.SaveDevices(currentDevices);
                        SyncDevicesToLocation(selectedLocationID, currentDevices);
                        DisplayDevicesWithStatus(currentDevices, oldDevices);
                        LoadDatabaseDevices();

                        if (!string.IsNullOrEmpty(currentDisplayedIP))
                            ShowDeviceDetails(currentDisplayedIP);

                        int found = currentDevices.Count;
                        int withMac = currentDevices.Count(d => !string.IsNullOrEmpty(d.MacAddress));
                        EndScanUI($"✅ Discovery abgeschlossen — {found} Geräte, {withMac} mit MAC");
                        scanButton.Enabled = true;
                        remoteHardwareButton.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    Invoke(new MethodInvoker(() =>
                    {
                        MessageBox.Show($"Fehler: {ex.Message}");
                        EndScanUI("Fehler beim Scan");
                        scanButton.Enabled = true;
                        remoteHardwareButton.Enabled = true;
                    }));
                }
            });
        }

        // Ladebalken + Sanduhr-Animation starten
        private void BeginScanUI(string message)
        {
            scanProgressBar.Visible = true;
            scanProgressBar.Value = 0;
            scanAnimTimer.Start();
            statusLabel.Text = "⏳  " + message;
            Cursor = Cursors.AppStarting;
        }

        // Ladebalken + Sanduhr-Animation stoppen
        private void EndScanUI(string message)
        {
            scanProgressBar.Visible = false;
            scanAnimTimer.Stop();
            statusLabel.Text = message;
            Cursor = Cursors.Default;
        }

        private void SyncDevicesToLocation(int locationID, List<DeviceInfo> devices)
        {
            var location = dbManager.GetLocationByID(locationID);
            if (location == null) { statusLabel.Text = "Fehler: Standort nicht gefunden!"; return; }

            var assignedIPs = new HashSet<string>(
                dbManager.GetIPsWithWorkstationByLocation(locationID).Select(x => x.IPAddress));

            var assignedMACs = new HashSet<string>(
                dbManager.GetIPsWithWorkstationByLocation(locationID)
                    .Select(x => x.IPAddress)
                    .SelectMany(ip => dbManager.LoadDevices("Alle")
                        .Where(d => d.IP == ip && !string.IsNullOrEmpty(d.MacAddress))
                        .Select(d => d.MacAddress)));

            int added = 0, skipped = 0;

            foreach (var device in devices)
            {
                if (assignedIPs.Contains(device.IP)) { skipped++; continue; }
                if (!string.IsNullOrEmpty(device.MacAddress) && assignedMACs.Contains(device.MacAddress)) { skipped++; continue; }
                dbManager.AddIPToLocation(locationID, device.IP, device.Hostname ?? "");
                try { dbManager.SyncDeviceToCustomerDb(device, location.CustomerID); } catch { }
                added++;
            }

            statusLabel.Text = added > 0
                ? $"{added} neue Geräte zu '{location.Name}' hinzugefügt, {skipped} bereits vorhanden"
                : $"Alle {skipped} Geräte bereits in '{location.Name}' vorhanden";
        }

        /// Extrahiert "Hersteller : XYZ" aus HW-Info-Text (WMI/Linux/ADB).
        private static string ExtractVendorFromHwInfo(string hwInfo)
        {
            if (string.IsNullOrEmpty(hwInfo)) return null;
            foreach (var line in hwInfo.Split('\n'))
            {
                var t = line.Trim();
                if (t.StartsWith("Hersteller") && t.Contains(":"))
                {
                    var val = t.Substring(t.IndexOf(':') + 1).Trim();
                    if (!string.IsNullOrEmpty(val) &&
                        val != "To Be Filled By O.E.M." &&
                        val != "Default string" &&
                        val.Length > 1)
                        return val;
                }
            }
            return null;
        }

        private static string CleanHostname(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^.+\s+\([0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}\)$")) return "";
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^[0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5}$")) return "";
            if (System.Text.RegularExpressions.Regex.IsMatch(raw,
                @"^\d{1,3}(\.\d{1,3}){3}$")) return "";
            return raw;
        }

        private void DisplayDevicesWithStatus(List<DeviceInfo> currentDevices, List<DeviceInfo> previousDevices)
        {
            deviceTable.Rows.Clear();
            var displayed   = new HashSet<string>();
            var locationIPs = dbManager.GetIPsWithWorkstationByLocation(selectedLocationID);

            var dbDevices   = dbManager.LoadDevices("Alle")
                                       .Where(d => locationIPs.Any(l => l.IPAddress == d.IP))
                                       .GroupBy(d => d.IP)
                                       .Select(g => g.First())
                                       .ToList();

            var lastScanIPs = new HashSet<string>(dbDevices.Select(d => d.IP));
            var dbHostnames = dbDevices.Where(d => !string.IsNullOrEmpty(d.Hostname)).ToDictionary(d => d.IP, d => d.Hostname);
            var dbTypes     = dbDevices.Where(d => d.DeviceType != DeviceType.Unbekannt).ToDictionary(d => d.IP, d => d.DeviceType);
            var dbVendors   = dbDevices.Where(d => !string.IsNullOrEmpty(d.Vendor)).ToDictionary(d => d.IP, d => d.Vendor);
            var dbComments  = dbDevices.ToDictionary(d => d.IP, d => d.Comment ?? "");
            var currentIPs  = new HashSet<string>(currentDevices.Select(d => d.IP));

            foreach (var dev in currentDevices)
                if (displayed.Add(dev.IP))
                {
                    string raw      = dbHostnames.TryGetValue(dev.IP, out string dbName) && !string.IsNullOrEmpty(dbName) ? dbName : dev.Hostname;
                    string hostname = CleanHostname(raw);
                    var dt          = dbTypes.TryGetValue(dev.IP, out DeviceType dbDt) ? dbDt : dev.DeviceType;
                    string vendor   = dbVendors.TryGetValue(dev.IP, out string dbVendor) && !string.IsNullOrEmpty(dbVendor) ? dbVendor : (dev.Vendor ?? "");
                    string comment  = dbComments.TryGetValue(dev.IP, out string c) ? c : "";
                    string mac      = !string.IsNullOrEmpty(dev.MacAddress) ? dev.MacAddress : dbDevices.FirstOrDefault(d => d.IP == dev.IP)?.MacAddress ?? "";
                    deviceTable.Rows.Add(
                        lastScanIPs.Contains(dev.IP) ? "AKTIV" : "NEU",
                        dev.IP, string.IsNullOrEmpty(hostname) ? "-" : hostname,
                        $"{DeviceTypeHelper.GetIcon(dt)} {DeviceTypeHelper.GetLabel(dt)}",
                        vendor, mac, comment, dev.Status, dev.Ports);
                }

            foreach (var last in dbDevices)
                if (!currentIPs.Contains(last.IP) && displayed.Add(last.IP))
                {
                    string hostname = CleanHostname(last.Hostname ?? "");
                    deviceTable.Rows.Add(
                        "OFFLINE", last.IP, string.IsNullOrEmpty(hostname) ? "-" : hostname,
                        $"{DeviceTypeHelper.GetIcon(last.DeviceType)} {DeviceTypeHelper.GetLabel(last.DeviceType)}",
                        last.Vendor ?? "", last.MacAddress ?? "", last.Comment ?? "", "Down", "-");
                }
        }

        // =========================================================
        // === HARDWARE / SOFTWARE / SNMP SCANS ===
        // =========================================================

        private void StartRemoteHardwareQuery()
        {
            using (var form = new RemoteConnectionForm(dbManager))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;

                var targets = new List<(string IP, string Username, string Password)>();
                if (form.SelectedTargets.Count > 0)
                    targets.AddRange(form.SelectedTargets);
                else if (!string.IsNullOrEmpty(form.ComputerIP))
                    targets.Add((form.ComputerIP, form.Username, form.Password));

                if (targets.Count == 0) return;

                remoteHardwareButton.Enabled = false;
                int total = targets.Count, current = 0;
                BeginScanUI(total == 1 ? $"Verbinde zu {targets[0].IP}..." : $"Starte Remote-Scan für {total} Geräte...");

                System.Threading.Tasks.Task.Run(() =>
                {
                    var errors = new List<string>();
                    int success = 0;
                    foreach (var target in targets)
                    {
                        current++;
                        Invoke(new MethodInvoker(() => statusLabel.Text = $"⏳  Gerät {current}/{total}: {target.IP}..."));
                        try
                        {
                            string hw = new HardwareManager().GetRemoteHardwareInfo(target.IP, target.Username, target.Password);
                            var sw    = new SoftwareManager().GetRemoteSoftware(target.IP, target.Username, target.Password);
                            string pcName = target.IP;
                            foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                            dbManager.CheckForUpdates(sw, pcName);
                            dbManager.SaveHardwareInfo(target.IP, $"=== Remote abgefragt: {DateTime.Now:dd.MM.yyyy HH:mm} ===\n{hw}");
                            dbManager.SaveSoftware(sw);
                            success++;
                        }
                        catch (Exception ex) { errors.Add($"{target.IP}: {ex.Message}"); }
                    }

                    Invoke(new MethodInvoker(() =>
                    {
                        remoteHardwareButton.Enabled = true;
                        if (errors.Count == 0)
                            EndScanUI(total == 1 ? "Remote-Abfrage abgeschlossen" : $"✅ {success}/{total} Geräte erfolgreich abgefragt");
                        else
                        {
                            EndScanUI($"⚠️ {success}/{total} erfolgreich — {errors.Count} Fehler");
                            MessageBox.Show("Fehler bei folgenden Geräten:\n\n" + string.Join("\n", errors),
                                "Teilweise Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }));
                });
            }
        }

        private void StartHwSwScanForDevice(DeviceType? forceType = null)
        {
            if (string.IsNullOrEmpty(currentDisplayedIP))
            {
                MessageBox.Show("Bitte zuerst ein Gerät im Baum auswählen!", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string ip = currentDisplayedIP;
            bool isLocal = IsLocalIP(ip);
            var scanBtn    = FindControl<Button>("scanHwSwBtn");
            var scanStatus = FindControl<Label>("scanHwSwStatus");
            if (scanBtn != null) scanBtn.Enabled = false;
            if (scanStatus != null) scanStatus.Text = "Abfrage läuft...";

            if (isLocal)
            {
                System.Threading.Tasks.Task.Run(() =>
                {
                    string hw = new HardwareManager().GetHardwareInfo();
                    var sw    = new SoftwareManager().GetInstalledSoftware();
                    string pcName = ip;
                    foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                    dbManager.CheckForUpdates(sw, pcName);
                    dbManager.SaveSoftware(sw);
                    var device = dbManager.LoadDevices("Alle").FirstOrDefault(d => d.IP == ip);
                    if (device != null)
                        dbManager.SaveHardwareInfo(ip, $"=== Lokal abgefragt: {DateTime.Now:dd.MM.yyyy HH:mm} ===\n{hw}");
                    dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                    Invoke(new MethodInvoker(() =>
                    {
                        ShowDeviceDetails(ip);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = $"✅ Abgefragt: {DateTime.Now:HH:mm}";
                        statusLabel.Text = $"Hardware/Software für {ip} gespeichert";
                    }));
                });
            }
            else
            {
                var knownDevice = dbManager.LoadDevices("Alle").FirstOrDefault(d => d.IP == ip);
                var deviceType  = forceType ?? knownDevice?.DeviceType ?? DeviceType.Unbekannt;
                bool forcediOS  = forceType == DeviceType.MacOS;
                bool isLinux    = deviceType == DeviceType.Linux;
                bool isAndroid  = deviceType == DeviceType.Smartphone && !forcediOS && !IsAppleDevice(knownDevice);
                bool isApple    = forcediOS || (IsAppleDevice(knownDevice) && deviceType == DeviceType.Smartphone);

                // ── Linux via SSH ──────────────────────────────
                if (isLinux)
                {
                    using (var form = new LinuxSshConnectionForm())
                    {
                        form.SetIP(ip);
                        if (form.ShowDialog(this) != DialogResult.OK)
                        { if (scanBtn != null) scanBtn.Enabled = true; if (scanStatus != null) scanStatus.Text = ""; return; }

                        string sshHost = form.Host, sshUser = form.Username, sshPass = form.Password, sshKey = form.KeyFilePath;
                        int    sshPort = form.Port;

                        System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                string hw = linuxManager.GetHardwareInfo(sshHost, sshUser, sshPass, sshKey, sshPort);
                                var sw    = linuxManager.GetInstalledSoftware(sshHost, sshUser, sshPass, sshKey, sshPort);
                                foreach (var s in sw) s.PCName = ip;
                                dbManager.SaveHardwareInfo(ip, hw);
                                dbManager.SaveSoftware(sw);
                                dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                                Invoke(new MethodInvoker(() =>
                                {
                                    ShowDeviceDetails(ip);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = $"✅ Linux SSH: {DateTime.Now:HH:mm}";
                                    statusLabel.Text = $"Linux-Abfrage für {ip} gespeichert";
                                }));
                            }
                            catch (Exception ex)
                            {
                                Invoke(new MethodInvoker(() =>
                                {
                                    MessageBox.Show($"SSH-Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                                }));
                            }
                        });
                    }
                }
                // ── Android via ADB ────────────────────────────
                else if (isAndroid || deviceType == DeviceType.Smartphone)
                {
                    if (!adbManager.IsAdbAvailable)
                    {
                        MessageBox.Show(
                            "ADB (Android Debug Bridge) wurde nicht gefunden.\n\n" +
                            "Bitte Android Platform Tools installieren:\n" +
                            "https://developer.android.com/tools/releases/platform-tools\n\n" +
                            "Dann adb.exe in denselben Ordner wie Inventarisierung.exe kopieren.",
                            "ADB nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = "";
                        return;
                    }

                    var adbDlg = new AdbConnectionDialog(ip);
                    if (adbDlg.ShowDialog(this) != System.Windows.Forms.DialogResult.OK)
                    { if (scanBtn != null) scanBtn.Enabled = true; if (scanStatus != null) scanStatus.Text = ""; return; }

                    string ipPort   = adbDlg.IpPort;
                    bool doPair     = adbDlg.DoPairing;
                    string pairIp   = $"{ip}:{adbDlg.PairPort}";
                    string pairCode = adbDlg.PairCode;
                    if (scanStatus != null) scanStatus.Text = $"Verbinde mit {ipPort}...";

                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            if (doPair)
                            {
                                string pairResult = adbManager.PairDevice(pairIp, pairCode);
                                Invoke(new MethodInvoker(() => MessageBox.Show(pairResult, "Kopplung", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                            }
                            string hw = adbManager.GetDeviceInfo(ipPort);
                            var sw    = adbManager.GetInstalledApps(ipPort);
                            foreach (var s in sw) s.PCName = ip;
                            dbManager.SaveHardwareInfo(ip, hw);
                            dbManager.SaveSoftware(sw);
                            dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                            string adbName = AdbManager.ExtractAdbDeviceName(hw);
                            if (!string.IsNullOrEmpty(adbName)) dbManager.SetHostnameIfEmpty(ip, adbName);
                            Invoke(new MethodInvoker(() =>
                            {
                                ShowDeviceDetails(ip);
                                LoadCustomerTree();
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = $"✅ Android ADB: {DateTime.Now:HH:mm}";
                                statusLabel.Text = $"Android-Abfrage für {ip} gespeichert";
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new MethodInvoker(() =>
                            {
                                MessageBox.Show($"ADB-Fehler:\n\n{ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                            }));
                        }
                    });
                }
                // ── iOS via USB ────────────────────────────────
                else if (isApple && deviceType == DeviceType.Smartphone)
                {
                    if (!iosManager.IsAvailable)
                    {
                        MessageBox.Show(
                            "libimobiledevice wurde nicht gefunden.\n\n" +
                            "Bitte ideviceinfo.exe herunterladen und neben Inventarisierung.exe ablegen:\n" +
                            "https://github.com/libimobiledevice-win32/imobiledevice-net/releases\n\n" +
                            "Außerdem: iPhone per USB anschließen und 'Vertrauen' antippen.",
                            "libimobiledevice nicht gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        if (scanBtn != null) scanBtn.Enabled = true;
                        if (scanStatus != null) scanStatus.Text = "";
                        return;
                    }
                    if (scanStatus != null) scanStatus.Text = "iOS-Abfrage via USB...";
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            string hw = iosManager.GetDeviceInfo();
                            var sw    = new List<SoftwareInfo> { new SoftwareInfo
                                { Name = "iOS: Softwareliste nicht verfügbar (kein Jailbreak)", Source = "iOS", PCName = ip } };
                            dbManager.SaveHardwareInfo(ip, hw);
                            dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                            Invoke(new MethodInvoker(() =>
                            {
                                ShowDeviceDetails(ip);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = $"✅ iOS: {DateTime.Now:HH:mm}";
                                statusLabel.Text = "iOS-Abfrage gespeichert";
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new MethodInvoker(() =>
                            {
                                MessageBox.Show($"iOS-Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                if (scanBtn != null) scanBtn.Enabled = true;
                                if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                            }));
                        }
                    });
                }
                // ── Windows WMI (Standard) ─────────────────────
                else
                {
                    using (var form = new RemoteConnectionForm(dbManager))
                    {
                        form.SetIP(ip);
                        if (form.ShowDialog(this) != DialogResult.OK)
                        { if (scanBtn != null) scanBtn.Enabled = true; if (scanStatus != null) scanStatus.Text = ""; return; }

                        string username = form.Username, password = form.Password;
                        System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                string hw = new HardwareManager().GetRemoteHardwareInfo(ip, username, password);
                                var sw    = new SoftwareManager().GetRemoteSoftware(ip, username, password);
                                string pcName = ip;
                                foreach (var s in sw) { s.PCName = pcName; s.Timestamp = DateTime.Now; }
                                dbManager.CheckForUpdates(sw, pcName);
                                dbManager.SaveSoftware(sw);
                                dbManager.SaveHardwareInfo(ip, $"=== Remote abgefragt: {DateTime.Now:dd.MM.yyyy HH:mm} ===\n{hw}");
                                dbManager.SaveVendorFromScan(ip, ExtractVendorFromHwInfo(hw));
                                Invoke(new MethodInvoker(() =>
                                {
                                    ShowDeviceDetails(ip);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = $"✅ Remote abgefragt: {DateTime.Now:HH:mm}";
                                    statusLabel.Text = $"Hardware/Software für {ip} gespeichert";
                                }));
                            }
                            catch (Exception ex)
                            {
                                Invoke(new MethodInvoker(() =>
                                {
                                    MessageBox.Show($"Fehler: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    if (scanBtn != null) scanBtn.Enabled = true;
                                    if (scanStatus != null) scanStatus.Text = "❌ Fehler";
                                }));
                            }
                        });
                    }
                }
            }
        }

        private static bool IsAppleDevice(DatabaseDevice d)
        {
            if (d == null) return false;
            string vendor = (d.MacAddress ?? "").ToLower();
            string host   = (d.Hostname ?? "").ToLower();
            return host.Contains("iphone") || host.Contains("ipad") || vendor.Contains("apple");
        }

        private static bool IsAppleVendor(DatabaseDevice d)
            => d != null && (d.Hostname ?? "").ToLower().Contains("apple");

        private bool IsLocalIP(string ip)
        {
            try
            {
                var localIPs = System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName()).Select(a => a.ToString());
                return ip == "127.0.0.1" || ip == "::1" || localIPs.Contains(ip);
            }
            catch { return false; }
        }

        // ─────────────────────────────────────────────────────────
        // 📡 SNMP-SCAN
        // ─────────────────────────────────────────────────────────
        private void StartSnmpScan()
        {
            if (currentDevices.Count == 0)
            {
                MessageBox.Show("Bitte zuerst einen Netzwerk-Scan durchführen.", "Kein Scan vorhanden",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            scanButton.Enabled = false;
            string snmpVerLabel = snmpSettings.Version == 3 ? "v3" : snmpSettings.Version == 1 ? "v1" : "v2c";
            BeginScanUI($"SNMP-Scan läuft ({snmpVerLabel}, Community: {snmpSettings.Community})...");
            var devicesToQuery = currentDevices.ToList();
            System.Threading.Tasks.Task.Run(() =>
            {
                int success = 0;
                snmpManager.QueryDevices(devicesToQuery, snmpSettings, msg =>
                    Invoke(new MethodInvoker(() => statusLabel.Text = "⏳  " + msg)));
                foreach (var d in devicesToQuery)
                    if (d.SnmpData?.Success == true) success++;
                Invoke(new MethodInvoker(() =>
                {
                    EndScanUI($"✅ SNMP-Scan abgeschlossen — {success}/{devicesToQuery.Count} Geräte antworteten");
                    scanButton.Enabled = true;
                    if (!string.IsNullOrEmpty(currentDisplayedIP))
                        ShowDeviceDetails(currentDisplayedIP);
                }));
            });
        }

        private void OpenSnmpSettings()
        {
            using (var dlg = new SnmpSettingsDialog(snmpSettings))
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    snmpSettings = dlg.Settings;
        }
    }
}
