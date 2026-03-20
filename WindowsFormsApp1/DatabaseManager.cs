using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;

namespace NmapInventory
{
    public class DatabaseManager
    {
        private const string DB_PATH = "nmap_inventory.db";

        public void InitializeDatabase()
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Devices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        IP TEXT UNIQUE NOT NULL,
                        Hostname TEXT,
                        MacAddress TEXT UNIQUE,
                        FirstSeen DATETIME DEFAULT CURRENT_TIMESTAMP,
                        LastSeen DATETIME DEFAULT CURRENT_TIMESTAMP
                    );

                    CREATE TABLE IF NOT EXISTS DeviceScanHistory (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Status TEXT,
                        Ports TEXT,
                        ScanTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS DeviceMacHistory (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        MacAddress TEXT NOT NULL,
                        IPAddress TEXT,
                        Timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS DeviceSoftware (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Name TEXT NOT NULL,
                        Version TEXT,
                        Publisher TEXT,
                        InstallLocation TEXT,
                        InstallDate TEXT,
                        Source TEXT,
                        QueryTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS SoftwareHistory (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Name TEXT NOT NULL,
                        OldVersion TEXT,
                        NewVersion TEXT,
                        ChangeTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS Customers (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        Name TEXT NOT NULL,
                        Address TEXT,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP
                    );

                    -- Selbstreferenzierende Tabelle fuer beliebig tiefe Hierarchie:
                    -- Level 0 = Standort (Root)
                    -- Level 1 = Abteilung
                    -- Level 2+ = Unterabteilung(en) (beliebig tief)
                   -- ParentID = NULL bedeutet: Diese Location ist ein Root-Standort
-- Teil 1
                    CREATE TABLE IF NOT EXISTS Locations (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        CustomerID INTEGER NOT NULL,
                        ParentID INTEGER DEFAULT NULL,
                        Name TEXT NOT NULL,
                        Address TEXT,
                        Level INTEGER DEFAULT 0,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(CustomerID) REFERENCES Customers(ID) ON DELETE CASCADE,
                        FOREIGN KEY(ParentID) REFERENCES Locations(ID) ON DELETE CASCADE
                    );

                    -- Geraete-Zuweisung zu einem Location-Knoten (beliebige Ebene)
                    -- Verbindet Locations mit Devices aus der Devices-Tabelle
                    CREATE TABLE IF NOT EXISTS LocationDevices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        LocationID INTEGER NOT NULL,
                        DeviceID INTEGER NOT NULL,
                        AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(LocationID) REFERENCES Locations(ID) ON DELETE CASCADE,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE,
                        UNIQUE(LocationID, DeviceID)
                    );

                    -- Vollstaendige Nmap-Scan-Details pro Gerät
                    CREATE TABLE IF NOT EXISTS DeviceNmapDetails (
                        ID          INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID    INTEGER NOT NULL,
                        OS          TEXT,
                        OSDetails   TEXT,
                        Vendor      TEXT,
                        Ports       TEXT,
                        PortsJson   TEXT,
                        RawOutput   TEXT,
                        ScanTime    DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    -- Hardware-Info Text pro Gerät (von WMI-Abfrage)
                    CREATE TABLE IF NOT EXISTS DeviceHardwareInfo (
                        ID          INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID    INTEGER NOT NULL,
                        HardwareText TEXT,
                        QueryTime   DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    -- Einzelne Ports pro Scan-Detail
                    CREATE TABLE IF NOT EXISTS DevicePorts (
                        ID          INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID    INTEGER NOT NULL,
                        Port        INTEGER NOT NULL,
                        Protocol    TEXT,
                        State       TEXT,
                        Service     TEXT,
                        Version     TEXT,
                        Banner      TEXT,
                        ScanTime    DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    -- Index fuer schnellen Port-Lookup
                    CREATE INDEX IF NOT EXISTS idx_device_ports ON DevicePorts(DeviceID);
                    CREATE TABLE IF NOT EXISTS LocationIPs (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        LocationID INTEGER NOT NULL,
                        IPAddress TEXT NOT NULL,
                        WorkstationName TEXT,
                        AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(LocationID) REFERENCES Locations(ID) ON DELETE CASCADE
                    );";

                using (var cmd = new SQLiteCommand(createTableQuery, conn))
                    cmd.ExecuteNonQuery();

                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN MacAddress TEXT UNIQUE");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN CustomHostname INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Customers ADD COLUMN CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN Level INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN ParentID INTEGER");
                TryAlterTable(conn, "ALTER TABLE LocationIPs ADD COLUMN AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_mac_address ON Devices(MacAddress)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_device_mac_history ON DeviceMacHistory(MacAddress)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_parent ON Locations(ParentID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_customer ON Locations(CustomerID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_devices ON LocationDevices(LocationID)");
            }
        }

        private void TryAlterTable(SQLiteConnection conn, string sql)
        {
            try { using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery(); }
            catch { }
        }

        // =========================================================
        // === DEVICES ===
        // =========================================================

        public void SaveDevices(List<DeviceInfo> devices)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var dev in devices)
                {
                    int deviceID = GetOrCreateDevice(conn, dev.IP, dev.Hostname, dev.MacAddress);
                    SaveDeviceScanHistory(conn, deviceID, dev.Status, dev.Ports);
                    UpdateDeviceLastSeen(conn, deviceID);

                    // MAC-History NUR einfügen wenn MAC neu ist (nicht bei jedem Scan)
                    if (!string.IsNullOrEmpty(dev.MacAddress))
                    {
                        string lastMac = null;
                        using (var cmd = new SQLiteCommand(
                            "SELECT MacAddress FROM DeviceMacHistory WHERE DeviceID=@ID ORDER BY Timestamp DESC LIMIT 1", conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", deviceID);
                            var result = cmd.ExecuteScalar();
                            lastMac = result?.ToString();
                        }

                        // Nur einfügen wenn MAC sich geändert hat oder noch kein Eintrag existiert
                        if (lastMac != dev.MacAddress)
                        {
                            using (var cmd = new SQLiteCommand(
                                "INSERT INTO DeviceMacHistory (DeviceID, MacAddress, IPAddress) VALUES (@DeviceID, @MAC, @IP)", conn))
                            {
                                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                                cmd.Parameters.AddWithValue("@MAC", MacParam(dev.MacAddress));
                                cmd.Parameters.AddWithValue("@IP", dev.IP);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }

        private int GetOrCreateDevice(SQLiteConnection conn, string ip, string hostname, string macAddress = null)
        {
            // Gerät suchen — zuerst per IP, falls nicht gefunden per MAC (Gerät hat IP gewechselt)
            int existingID = -1;

            using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP = @IP", conn))
            {
                cmd.Parameters.AddWithValue("@IP", ip);
                var result = cmd.ExecuteScalar();
                if (result != null) existingID = Convert.ToInt32(result);
            }

            // Per MAC suchen falls IP nicht gefunden und MAC vorhanden
            if (existingID < 0 && !string.IsNullOrWhiteSpace(macAddress))
            {
                using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE MacAddress = @MAC", conn))
                {
                    cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                    var result = cmd.ExecuteScalar();
                    if (result != null) existingID = Convert.ToInt32(result);
                }
            }

            if (existingID > 0)
            {
                // Hostname nur aktualisieren wenn er NICHT manuell gesetzt wurde (CustomHostname = 0)
                using (var cmd = new SQLiteCommand(
                    @"UPDATE Devices SET 
                        Hostname   = CASE WHEN CustomHostname = 0 AND @Hostname != '' THEN @Hostname ELSE Hostname END,
                        MacAddress = CASE WHEN @MAC IS NOT NULL AND @MAC != '' THEN @MAC ELSE MacAddress END,
                        IP         = @IP
                      WHERE ID = @ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", existingID);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.Parameters.AddWithValue("@Hostname", hostname ?? "");
                    cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                    cmd.ExecuteNonQuery();
                }
                return existingID;
            }

            // Neu anlegen
            using (var cmd = new SQLiteCommand(
                "INSERT INTO Devices (IP, Hostname, MacAddress) VALUES (@IP, @Hostname, @MAC)", conn))
            {
                cmd.Parameters.AddWithValue("@IP", ip);
                cmd.Parameters.AddWithValue("@Hostname", hostname ?? "");
                cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP = @IP", conn))
            {
                cmd.Parameters.AddWithValue("@IP", ip);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }


        private void SaveDeviceScanHistory(SQLiteConnection conn, int deviceID, string status, string ports)
        {
            using (var cmd = new SQLiteCommand("INSERT INTO DeviceScanHistory (DeviceID, Status, Ports) VALUES (@DeviceID, @Status, @Ports)", conn))
            {
                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                cmd.Parameters.AddWithValue("@Status", status ?? "");
                cmd.Parameters.AddWithValue("@Ports", ports ?? "");
                cmd.ExecuteNonQuery();
            }
        }

        private void UpdateDeviceLastSeen(SQLiteConnection conn, int deviceID)
        {
            using (var cmd = new SQLiteCommand("UPDATE Devices SET LastSeen = CURRENT_TIMESTAMP WHERE ID = @ID", conn))
            {
                cmd.Parameters.AddWithValue("@ID", deviceID);
                cmd.ExecuteNonQuery();
            }
        }

        // =========================================================
        // === NMAP SCAN DETAILS ===
        // =========================================================

        /// <summary>
        /// Speichert vollständige Nmap-Scan-Details (OS, Ports, Banner, Raw) für ein Gerät.
        /// Alte Ports werden ersetzt — immer aktueller Stand.
        /// </summary>
        public void SaveNmapDetails(DeviceInfo device)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                int deviceID = GetOrCreateDevice(conn, device.IP, device.Hostname, device.MacAddress);

                // Nmap-Detail-Eintrag anlegen
                using (var cmd = new SQLiteCommand(@"
                    INSERT INTO DeviceNmapDetails (DeviceID, OS, OSDetails, Vendor, Ports, PortsJson, RawOutput)
                    VALUES (@DeviceID, @OS, @OSDetails, @Vendor, @Ports, @PortsJson, @RawOutput)", conn))
                {
                    cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                    cmd.Parameters.AddWithValue("@OS", device.OS ?? "");
                    cmd.Parameters.AddWithValue("@OSDetails", device.OSDetails ?? "");
                    cmd.Parameters.AddWithValue("@Vendor", device.Vendor ?? "");
                    cmd.Parameters.AddWithValue("@Ports", device.Ports ?? "");
                    cmd.Parameters.AddWithValue("@PortsJson", BuildPortsJson(device.OpenPorts));
                    cmd.Parameters.AddWithValue("@RawOutput", "");
                    cmd.ExecuteNonQuery();
                }

                // Alte Ports löschen und neu schreiben
                using (var cmd = new SQLiteCommand("DELETE FROM DevicePorts WHERE DeviceID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", deviceID);
                    cmd.ExecuteNonQuery();
                }

                foreach (var port in device.OpenPorts)
                {
                    using (var cmd = new SQLiteCommand(@"
                        INSERT INTO DevicePorts (DeviceID, Port, Protocol, State, Service, Version, Banner)
                        VALUES (@DeviceID, @Port, @Protocol, @State, @Service, @Version, @Banner)", conn))
                    {
                        cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                        cmd.Parameters.AddWithValue("@Port", port.Port);
                        cmd.Parameters.AddWithValue("@Protocol", port.Protocol ?? "tcp");
                        cmd.Parameters.AddWithValue("@State", port.State ?? "open");
                        cmd.Parameters.AddWithValue("@Service", port.Service ?? "");
                        cmd.Parameters.AddWithValue("@Version", port.Version ?? "");
                        cmd.Parameters.AddWithValue("@Banner", port.Banner ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }

                // OS auch in Devices-Tabelle aktualisieren
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN OS TEXT");
                using (var cmd = new SQLiteCommand("UPDATE Devices SET OS=@OS WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@OS", device.OS ?? "");
                    cmd.Parameters.AddWithValue("@ID", deviceID);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Lädt die letzten Nmap-Details für ein Gerät per IP.
        /// </summary>
        public NmapScanDetail GetLatestNmapDetail(string ip)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT n.ID, n.DeviceID, d.IP, d.Hostname, d.MacAddress, n.OS, n.OSDetails,
                           n.Vendor, n.Ports, n.PortsJson, n.ScanTime
                    FROM DeviceNmapDetails n
                    JOIN Devices d ON n.DeviceID = d.ID
                    WHERE d.IP = @IP
                    ORDER BY n.ScanTime DESC LIMIT 1", conn))
                {
                    cmd.Parameters.AddWithValue("@IP", ip);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                            return new NmapScanDetail
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                DeviceID = Convert.ToInt32(reader["DeviceID"]),
                                IP = reader["IP"].ToString(),
                                Hostname = reader["Hostname"]?.ToString(),
                                MacAddress = reader["MacAddress"]?.ToString(),
                                OS = reader["OS"]?.ToString(),
                                OSDetails = reader["OSDetails"]?.ToString(),
                                Vendor = reader["Vendor"]?.ToString(),
                                Ports = reader["Ports"]?.ToString(),
                                PortsJson = reader["PortsJson"]?.ToString(),
                                ScanTime = Convert.ToDateTime(reader["ScanTime"])
                            };
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Lädt alle gespeicherten Ports für ein Gerät (aktuellster Stand).
        /// </summary>
        public List<NmapPort> GetPortsByDevice(string ip)
        {
            var ports = new List<NmapPort>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT dp.Port, dp.Protocol, dp.State, dp.Service, dp.Version, dp.Banner
                    FROM DevicePorts dp
                    JOIN Devices d ON dp.DeviceID = d.ID
                    WHERE d.IP = @IP
                    ORDER BY dp.Port", conn))
                {
                    cmd.Parameters.AddWithValue("@IP", ip);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            ports.Add(new NmapPort
                            {
                                Port = Convert.ToInt32(reader["Port"]),
                                Protocol = reader["Protocol"]?.ToString(),
                                State = reader["State"]?.ToString(),
                                Service = reader["Service"]?.ToString(),
                                Version = reader["Version"]?.ToString(),
                                Banner = reader["Banner"]?.ToString()
                            });
                }
            }
            return ports;
        }

        /// <summary>
        /// Speichert den Hardware-Info-Text (von WMI-Abfrage) für ein Gerät per IP.
        /// </summary>
        public void SaveHardwareInfo(string ip, string hardwareText)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                int deviceID = GetOrCreateDevice(conn, ip, null);
                using (var cmd = new SQLiteCommand(
                    "INSERT INTO DeviceHardwareInfo (DeviceID, HardwareText) VALUES (@DeviceID, @Text)", conn))
                {
                    cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                    cmd.Parameters.AddWithValue("@Text", hardwareText ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Lädt den zuletzt gespeicherten Hardware-Info-Text für ein Gerät.
        /// </summary>
        public string GetLatestHardwareInfo(string ip)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT h.HardwareText
                    FROM DeviceHardwareInfo h
                    JOIN Devices d ON h.DeviceID = d.ID
                    WHERE d.IP = @IP
                    ORDER BY h.QueryTime DESC LIMIT 1", conn))
                {
                    cmd.Parameters.AddWithValue("@IP", ip);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString();
                }
            }
        }

        /// <summary>
        /// Speichert einen manuell gesetzten Hostnamen.
        /// CustomHostname = 1 → wird bei späteren Scans nicht überschrieben.
        /// Aktualisiert auch LocationIPs.WorkstationName für Konsistenz.
        /// </summary>
        public void UpdateDeviceHostname(string ip, string newHostname)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();

                // Devices-Tabelle aktualisieren + als manuell markieren
                using (var cmd = new SQLiteCommand(
                    "UPDATE Devices SET Hostname=@Hostname, CustomHostname=1 WHERE IP=@IP", conn))
                {
                    cmd.Parameters.AddWithValue("@Hostname", newHostname);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.ExecuteNonQuery();
                }

                // Auch WorkstationName in LocationIPs aktualisieren
                using (var cmd = new SQLiteCommand(
                    "UPDATE LocationIPs SET WorkstationName=@Name WHERE IPAddress=@IP", conn))
                {
                    cmd.Parameters.AddWithValue("@Name", newHostname);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Setzt CustomHostname zurück → Hostname wird beim nächsten Scan wieder automatisch gesetzt.
        /// </summary>
        public void ResetCustomHostname(string ip)
        {
            ExecuteNonQuery(
                "UPDATE Devices SET CustomHostname=0 WHERE IP=@IP",
                new[] { ("@IP", ip) });
        }

        private string BuildPortsJson(List<NmapPort> ports)
        {
            if (ports == null || ports.Count == 0) return "[]";
            var sb = new System.Text.StringBuilder("[");
            for (int i = 0; i < ports.Count; i++)
            {
                var p = ports[i];
                sb.Append($"{{\"port\":{p.Port},\"proto\":\"{p.Protocol}\",\"state\":\"{p.State}\"," +
                           $"\"service\":\"{p.Service}\",\"version\":\"{p.Version?.Replace("\"", "'")}\"," +
                           $"\"banner\":\"{p.Banner?.Replace("\"", "'")}\"}}");
                if (i < ports.Count - 1) sb.Append(",");
            }
            sb.Append("]");
            return sb.ToString();
        }

        // =========================================================
        // === MAC-ADRESSEN ===
        // =========================================================

        // Hilfsmethode: Macro-Parameter für MAC-Werte. Gibt DBNull.Value zurück für leere/whitespace Strings.
        private object MacParam(string mac)
        {
            if (string.IsNullOrWhiteSpace(mac)) return DBNull.Value;
            return mac.Trim().ToUpperInvariant();
        }

        public void SaveDeviceMacAddress(string ip, string macAddress, string hostname)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                int deviceID = GetOrCreateDevice(conn, ip, hostname);

                using (var cmd = new SQLiteCommand("UPDATE Devices SET MacAddress=@MAC WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                    cmd.Parameters.AddWithValue("@ID", deviceID);
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = new SQLiteCommand("INSERT INTO DeviceMacHistory (DeviceID, MacAddress, IPAddress) VALUES (@DeviceID, @MAC, @IP)", conn))
                {
                    cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                    cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public (int ID, string IP, string Hostname)? GetDeviceByMacAddress(string macAddress)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, IP, Hostname FROM Devices WHERE MacAddress=@MAC", conn))
                {
                    cmd.Parameters.AddWithValue("@MAC", MacParam(macAddress));
                    using (var reader = cmd.ExecuteReader())
                        if (reader.Read())
                            return (Convert.ToInt32(reader["ID"]), reader["IP"].ToString(), reader["Hostname"]?.ToString());
                }
            }
            return null;
        }

        public List<(string MacAddress, string IPAddress, DateTime Timestamp)> GetDeviceMacHistory(int deviceID)
        {
            var history = new List<(string, string, DateTime)>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT DISTINCT MacAddress, IPAddress, Timestamp FROM DeviceMacHistory WHERE DeviceID=@DeviceID ORDER BY Timestamp DESC", conn))
                {
                    cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            history.Add((reader["MacAddress"].ToString(), reader["IPAddress"]?.ToString(), Convert.ToDateTime(reader["Timestamp"])));
                }
            }
            return history;
        }

        public List<DatabaseDevice> LoadDevices(string filter = "Alle")
        {
            var devices = new List<DatabaseDevice>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string whereClause = GetDateFilter(filter, "d.LastSeen");
                string query = $@"
                    SELECT DISTINCT 
                        d.ID, d.IP, d.Hostname, d.MacAddress,
                        d.LastSeen as Zeitstempel,
                        sh.Status, sh.Ports
                    FROM Devices d
                    LEFT JOIN DeviceScanHistory sh ON d.ID = sh.DeviceID
                    {whereClause}
                    ORDER BY d.LastSeen DESC";

                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        devices.Add(new DatabaseDevice
                        {
                            ID = Convert.ToInt32(reader["ID"]),
                            Zeitstempel = reader["Zeitstempel"]?.ToString(),
                            IP = reader["IP"].ToString(),
                            Hostname = reader["Hostname"]?.ToString(),
                            MacAddress = reader["MacAddress"]?.ToString(),
                            Status = reader["Status"]?.ToString(),
                            Ports = reader["Ports"]?.ToString()
                        });
            }
            return devices;
        }

        public void UpdateDevices(List<DatabaseDevice> devices)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var dev in devices)
                    using (var cmd = new SQLiteCommand("UPDATE Devices SET Hostname=@Hostname, MacAddress=@MAC WHERE ID=@ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", dev.ID);
                        cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                        cmd.Parameters.AddWithValue("@MAC", MacParam(dev.MacAddress));
                        cmd.ExecuteNonQuery();
                    }
            }
        }

        public void DeleteDevice(int id)
            => ExecuteNonQuery("DELETE FROM Devices WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        // =========================================================
        // === SOFTWARE ===
        // =========================================================

        public void SaveSoftware(List<SoftwareInfo> software)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var sw in software)
                {
                    int deviceID = GetDeviceIDByName(conn, sw.PCName);
                    if (deviceID <= 0) continue;

                    string oldVersion = GetLastSoftwareVersion(conn, deviceID, sw.Name);
                    if (!string.IsNullOrEmpty(oldVersion) && oldVersion != sw.Version)
                        SaveSoftwareHistory(conn, deviceID, sw.Name, oldVersion, sw.Version);

                    SaveOrUpdateSoftware(conn, deviceID, sw);
                }
            }
        }

        private int GetDeviceIDByName(SQLiteConnection conn, string nameOrIP)
        {
            using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP = @Name OR Hostname = @Name LIMIT 1", conn))
            {
                cmd.Parameters.AddWithValue("@Name", nameOrIP);
                var result = cmd.ExecuteScalar();
                return result != null ? Convert.ToInt32(result) : -1;
            }
        }

        private string GetLastSoftwareVersion(SQLiteConnection conn, int deviceID, string softwareName)
        {
            using (var cmd = new SQLiteCommand("SELECT Version FROM DeviceSoftware WHERE DeviceID = @DeviceID AND Name = @Name ORDER BY QueryTime DESC LIMIT 1", conn))
            {
                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                cmd.Parameters.AddWithValue("@Name", softwareName);
                return cmd.ExecuteScalar()?.ToString();
            }
        }

        private void SaveSoftwareHistory(SQLiteConnection conn, int deviceID, string name, string oldVersion, string newVersion)
        {
            using (var cmd = new SQLiteCommand("INSERT INTO SoftwareHistory (DeviceID, Name, OldVersion, NewVersion) VALUES (@DeviceID, @Name, @OldVersion, @NewVersion)", conn))
            {
                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                cmd.Parameters.AddWithValue("@Name", name);
                cmd.Parameters.AddWithValue("@OldVersion", oldVersion);
                cmd.Parameters.AddWithValue("@NewVersion", newVersion);
                cmd.ExecuteNonQuery();
            }
        }

        private void SaveOrUpdateSoftware(SQLiteConnection conn, int deviceID, SoftwareInfo sw)
        {
            using (var cmd = new SQLiteCommand("SELECT ID FROM DeviceSoftware WHERE DeviceID = @DeviceID AND Name = @Name", conn))
            {
                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                cmd.Parameters.AddWithValue("@Name", sw.Name);
                var result = cmd.ExecuteScalar();

                if (result != null)
                    using (var updateCmd = new SQLiteCommand("UPDATE DeviceSoftware SET Version=@Version, Publisher=@Publisher, InstallDate=@InstallDate, Source=@Source, QueryTime=CURRENT_TIMESTAMP WHERE ID=@ID", conn))
                    {
                        updateCmd.Parameters.AddWithValue("@ID", Convert.ToInt32(result));
                        updateCmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        updateCmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        updateCmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        updateCmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                        updateCmd.ExecuteNonQuery();
                    }
                else
                    using (var insertCmd = new SQLiteCommand("INSERT INTO DeviceSoftware (DeviceID, Name, Version, Publisher, InstallLocation, InstallDate, Source) VALUES (@DeviceID, @Name, @Version, @Publisher, @InstallLocation, @InstallDate, @Source)", conn))
                    {
                        insertCmd.Parameters.AddWithValue("@DeviceID", deviceID);
                        insertCmd.Parameters.AddWithValue("@Name", sw.Name);
                        insertCmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        insertCmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        insertCmd.Parameters.AddWithValue("@InstallLocation", sw.InstallLocation ?? "");
                        insertCmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        insertCmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                        insertCmd.ExecuteNonQuery();
                    }
            }
        }

        public List<DatabaseSoftware> LoadSoftware(string filter = "Alle")
        {
            var software = new List<DatabaseSoftware>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string whereClause = GetDateFilter(filter, "ds.QueryTime");
                string query = $@"
                    SELECT 
                        ds.ID, ds.QueryTime as Zeitstempel, d.Hostname as PCName,
                        ds.Name, ds.Version, ds.Publisher, ds.InstallDate,
                        ds.QueryTime as LastUpdate
                    FROM DeviceSoftware ds
                    JOIN Devices d ON ds.DeviceID = d.ID
                    {whereClause}
                    ORDER BY ds.QueryTime DESC";

                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        software.Add(new DatabaseSoftware
                        {
                            ID = Convert.ToInt32(reader["ID"]),
                            Zeitstempel = reader["Zeitstempel"].ToString(),
                            Name = reader["Name"].ToString(),
                            Version = reader["Version"]?.ToString(),
                            Publisher = reader["Publisher"]?.ToString(),
                            InstallDate = reader["InstallDate"]?.ToString(),
                            PCName = reader["PCName"]?.ToString(),
                            LastUpdate = reader["LastUpdate"]?.ToString()
                        });
            }
            return software;
        }

        public void UpdateSoftwareEntries(List<DatabaseSoftware> software)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var sw in software)
                    using (var cmd = new SQLiteCommand("UPDATE DeviceSoftware SET Name=@Name, Version=@Version, Publisher=@Publisher, InstallDate=@InstallDate WHERE ID=@ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", sw.ID);
                        cmd.Parameters.AddWithValue("@Name", sw.Name ?? "");
                        cmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        cmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        cmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        cmd.ExecuteNonQuery();
                    }
            }
        }

        public void DeleteSoftwareEntry(int id)
            => ExecuteNonQuery("DELETE FROM DeviceSoftware WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        // =========================================================
        // === CUSTOMERS ===
        // =========================================================

        public List<Customer> GetCustomers()
        {
            var customers = new List<Customer>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, Name, Address FROM Customers ORDER BY Name", conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        customers.Add(new Customer
                        {
                            ID = Convert.ToInt32(reader["ID"]),
                            Name = reader["Name"].ToString(),
                            Address = reader["Address"]?.ToString() ?? ""
                        });
            }
            return customers;
        }

        public void AddCustomer(string name, string address)
            => ExecuteNonQuery("INSERT INTO Customers (Name, Address) VALUES (@Name, @Address)", new[] { ("@Name", name), ("@Address", address ?? "") });

        public void UpdateCustomer(int id, string name, string address)
            => ExecuteNonQuery("UPDATE Customers SET Name=@Name, Address=@Address WHERE ID=@ID", new[] { ("@ID", id.ToString()), ("@Name", name), ("@Address", address ?? "") });

        public void DeleteCustomer(int id)
            => ExecuteNonQuery("DELETE FROM Customers WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        // =========================================================
        // === LOCATIONS (HIERARCHISCH / BAUM) ===
        //
        // Struktur:
        //   Customer
        //   └── Standort          (Level 0, ParentID = NULL)
        //       └── Abteilung     (Level 1, ParentID = Standort.ID)
        //           └── Unterabt. (Level 2+, ParentID = Abteilung.ID)
        //               └── ...   (beliebig tief)
        // =========================================================

        /// <summary>
        /// Laedt alle Root-Standorte eines Kunden (Level 0, kein Parent).
        /// </summary>
        public List<Location> GetLocationsByCustomer(int customerId)
            => GetLocationsByParent(customerId, null);

        /// <summary>
        /// Laedt alle direkten Kinder einer Location.
        /// </summary>
        public List<Location> GetChildLocations(int parentLocationId)
        {
            var locations = new List<Location>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE ParentID = @ParentID ORDER BY Name", conn))
                {
                    cmd.Parameters.AddWithValue("@ParentID", parentLocationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            locations.Add(MapLocation(reader));
                }
            }
            return locations;
        }

        private List<Location> GetLocationsByParent(int customerId, int? parentId)
        {
            var locations = new List<Location>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string query = parentId.HasValue
                    ? "SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE CustomerID = @CustomerID AND ParentID = @ParentID ORDER BY Name"
                    : "SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE CustomerID = @CustomerID AND ParentID IS NULL ORDER BY Name";

                using (var cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerId);
                    if (parentId.HasValue)
                        cmd.Parameters.AddWithValue("@ParentID", parentId.Value);

                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            locations.Add(MapLocation(reader));
                }
            }
            return locations;
        }

        /// <summary>
        /// Laedt den vollstaendigen Baum eines Kunden in einem einzigen rekursiven SQL-Query.
        /// Rueckgabe: flache Liste, nach Level und Name sortiert.
        /// </summary>
        /// Teil 3
        public List<Location> GetFullLocationTree(int customerId)
        {
            var locations = new List<Location>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                // Rekursives CTE: laedt den gesamten Baum in einem Query
                // Teil 2
                string query = @"
                    WITH RECURSIVE loc_tree AS (
                        -- Startknoten: alle Root-Standorte des Kunden
                        SELECT ID, CustomerID, ParentID, Name, Address, Level
                        FROM Locations
                        WHERE CustomerID = @CustomerID AND ParentID IS NULL

                        UNION ALL

                        -- Rekursiv alle Kinder laden
                        SELECT l.ID, l.CustomerID, l.ParentID, l.Name, l.Address, l.Level
                        FROM Locations l
                        INNER JOIN loc_tree lt ON l.ParentID = lt.ID
                    )
                    SELECT * FROM loc_tree ORDER BY Level, Name";

                using (var cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            locations.Add(MapLocation(reader));
                }
            }
            return locations;
        }

        /// <summary>
        /// Gibt eine Location als formatierten Baum-String aus (fuer Debug / Logs).
        /// Beispiel:
        ///   Werk Muenchen
        ///   ├── Produktion
        ///   │   └── Linie A
        ///   └── Verwaltung
        /// </summary>
        public string PrintLocationTree(int customerId)
        {
            var allLocations = GetFullLocationTree(customerId);
            var result = new System.Text.StringBuilder();

            void PrintNode(int? parentId, string prefix, bool isLast)
            {
                var children = allLocations.Where(l =>
                    parentId == null ? l.ParentID == -1 : l.ParentID == parentId).ToList();

                for (int i = 0; i < children.Count; i++)
                {
                    var loc = children[i];
                    bool last = i == children.Count - 1;
                    result.AppendLine($"{prefix}{(last ? "└── " : "├── ")}{loc.Name} (Level {loc.Level})");
                    PrintNode(loc.ID, prefix + (last ? "    " : "│   "), last);
                }
            }

            PrintNode(null, "", false);
            return result.ToString();
        }

        public void AddLocation(int customerId, string name, string address)
            => ExecuteNonQuery(
                "INSERT INTO Locations (CustomerID, ParentID, Name, Address, Level) VALUES (@CustomerID, NULL, @Name, @Address, 0)",
                new[] { ("@CustomerID", customerId.ToString()), ("@Name", name), ("@Address", address ?? "") });

        /// <summary>
        /// Fuegt eine Unterabteilung unter einem beliebigen Elternknoten hinzu.
        /// Level wird automatisch als ParentLevel + 1 berechnet.
        /// </summary>
        public void AddChildLocation(int parentLocationId, string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                int parentLevel = 0;
                int customerID = 0;

                using (var cmd = new SQLiteCommand("SELECT Level, CustomerID FROM Locations WHERE ID = @ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", parentLocationId);
                    using (var reader = cmd.ExecuteReader())
                        if (reader.Read())
                        {
                            parentLevel = Convert.ToInt32(reader["Level"]);
                            customerID = Convert.ToInt32(reader["CustomerID"]);
                        }
                }

                using (var cmd = new SQLiteCommand(
                    "INSERT INTO Locations (CustomerID, ParentID, Name, Address, Level) VALUES (@CustomerID, @ParentID, @Name, @Address, @Level)", conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerID);
                    cmd.Parameters.AddWithValue("@ParentID", parentLocationId);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.Parameters.AddWithValue("@Level", parentLevel + 1); // automatisch Level erhoehen
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void UpdateLocation(int id, string name, string address)
            => ExecuteNonQuery("UPDATE Locations SET Name=@Name, Address=@Address WHERE ID=@ID",
                new[] { ("@ID", id.ToString()), ("@Name", name), ("@Address", address ?? "") });

        /// <summary>
        /// Loescht eine Location inkl. aller Kinder und Geraete-Zuweisungen (ON DELETE CASCADE).
        /// </summary>
        public void DeleteLocation(int id)
            => ExecuteNonQuery("DELETE FROM Locations WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        public Location GetLocationByID(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE ID = @ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    using (var reader = cmd.ExecuteReader())
                        if (reader.Read()) return MapLocation(reader);
                }
            }
            return null;
        }

        private Location MapLocation(SQLiteDataReader reader) => new Location
        {
            ID = Convert.ToInt32(reader["ID"]),
            CustomerID = Convert.ToInt32(reader["CustomerID"]),
            ParentID = reader["ParentID"] == DBNull.Value ? -1 : Convert.ToInt32(reader["ParentID"]),
            Name = reader["Name"].ToString(),
            Address = reader["Address"]?.ToString() ?? "",
            Level = Convert.ToInt32(reader["Level"])
        };

        // =========================================================
        // === LOCATION DEVICES (neu: Geraete einem Knoten zuweisen) ===
        // =========================================================

        /// <summary>
        /// Weist ein Device (aus Devices-Tabelle) einer Location zu.
        /// </summary>
        public void AssignDeviceToLocation(int locationId, int deviceId)
            => ExecuteNonQuery(
                "INSERT OR IGNORE INTO LocationDevices (LocationID, DeviceID) VALUES (@LocationID, @DeviceID)",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });

        /// <summary>
        /// Entfernt die Zuweisung eines Devices von einer Location.
        /// </summary>
        public void UnassignDeviceFromLocation(int locationId, int deviceId)
            => ExecuteNonQuery(
                "DELETE FROM LocationDevices WHERE LocationID=@LocationID AND DeviceID=@DeviceID",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });

        /// <summary>
        /// Laedt alle direkt zugewiesenen Geraete einer Location (nicht rekursiv).
        /// </summary>
        public List<DatabaseDevice> GetDevicesByLocation(int locationId)
        {
            var devices = new List<DatabaseDevice>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string query = @"
                    SELECT d.ID, d.IP, d.Hostname, d.MacAddress, d.LastSeen as Zeitstempel
                    FROM Devices d
                    INNER JOIN LocationDevices ld ON d.ID = ld.DeviceID
                    WHERE ld.LocationID = @LocationID
                    ORDER BY d.Hostname, d.IP";

                using (var cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            devices.Add(new DatabaseDevice
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                IP = reader["IP"].ToString(),
                                Hostname = reader["Hostname"]?.ToString(),
                                MacAddress = reader["MacAddress"]?.ToString(),
                                Zeitstempel = reader["Zeitstempel"]?.ToString()
                            });
                }
            }
            return devices;
        }

        /// <summary>
        /// Laedt alle Geraete einer Location UND aller ihrer Unterabteilungen (rekursiv).
        /// </summary>
        public List<DatabaseDevice> GetDevicesByLocationRecursive(int locationId)
        {
            var devices = new List<DatabaseDevice>();
            devices.AddRange(GetDevicesByLocation(locationId));

            foreach (var child in GetChildLocations(locationId))
                devices.AddRange(GetDevicesByLocationRecursive(child.ID));

            return devices;
        }

        // =========================================================
        // === LOCATION IPs (Rueckwaertskompatibilitaet) ===
        // =========================================================

        public List<LocationIP> GetIPsWithWorkstationByLocation(int locationId)
        {
            var ips = new List<LocationIP>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT ID, LocationID, IPAddress, WorkstationName FROM LocationIPs WHERE LocationID=@LocationID ORDER BY WorkstationName, IPAddress", conn))
                {
                    cmd.Parameters.AddWithValue("@LocationID", locationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            ips.Add(new LocationIP
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                LocationID = Convert.ToInt32(reader["LocationID"]),
                                IPAddress = reader["IPAddress"].ToString(),
                                WorkstationName = reader["WorkstationName"]?.ToString()
                            });
                }
            }
            return ips;
        }

        public List<(string IP, string Hostname)> GetAllIPsFromDevices()
        {
            var ips = new List<(string, string)>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT IP, Hostname FROM Devices WHERE IP IS NOT NULL AND IP != '' ORDER BY IP", conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        ips.Add((reader["IP"].ToString(), reader["Hostname"]?.ToString() ?? ""));
            }
            return ips;
        }

        public List<(int LocationID, string LocationName, string IPAddress, string WorkstationName)> GetAllDevicesFromDB()
        {
            var devices = new List<(int, string, string, string)>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT 
                        d.IP,
                        d.Hostname,
                        0 as DummyLocationID,
                        'Aus DB gescannte Geräte' as DummyLocationName
                    FROM Devices d
                    WHERE d.IP IS NOT NULL AND d.IP != ''
                    ORDER BY d.IP", conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        devices.Add((
                            0,
                            "Aus DB gescannte Geräte",
                            reader["IP"].ToString(),
                            reader["Hostname"]?.ToString() ?? ""
                        ));
            }
            return devices;
        }

        public void CheckForUpdates(List<SoftwareInfo> softwareList, string pcName)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                int deviceID = GetDeviceIDByName(conn, pcName);
                if (deviceID <= 0) return;

                foreach (var sw in softwareList)
                {
                    using (var cmd = new SQLiteCommand(@"
                        SELECT sh.ChangeTime 
                        FROM SoftwareHistory sh 
                        WHERE sh.DeviceID = @DeviceID AND sh.Name = @Name 
                        ORDER BY sh.ChangeTime DESC LIMIT 1", conn))
                    {
                        cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                        cmd.Parameters.AddWithValue("@Name", sw.Name);
                        var result = cmd.ExecuteScalar();
                        sw.LastUpdate = result != null
                            ? Convert.ToDateTime(result).ToString("dd.MM.yyyy HH:mm")
                            : "-";
                    }
                }
            }
        }

        public void AddIPToLocation(int locationId, string ip, string workstationName = "")
            => ExecuteNonQuery(
                "INSERT INTO LocationIPs (LocationID, IPAddress, WorkstationName) VALUES (@LocationID, @IPAddress, @WorkstationName)",
                new[] { ("@LocationID", locationId.ToString()), ("@IPAddress", ip), ("@WorkstationName", workstationName ?? "") });

        public void RemoveIPFromLocation(int locationId, string ip)
            => ExecuteNonQuery(
                "DELETE FROM LocationIPs WHERE LocationID=@LocationID AND IPAddress=@IPAddress",
                new[] { ("@LocationID", locationId.ToString()), ("@IPAddress", ip) });

        /// <summary>
        /// Laedt alle IPs einer Location und aller Unterabteilungen (rekursiv).
        /// </summary>
        public List<LocationIP> GetAllIPsRecursive(int locationId)
        {
            var allIPs = new List<LocationIP>();
            allIPs.AddRange(GetIPsWithWorkstationByLocation(locationId));

            foreach (var child in GetChildLocations(locationId))
                allIPs.AddRange(GetAllIPsRecursive(child.ID));

            return allIPs;
        }

        // =========================================================
        // === HILFSMETHODEN ===
        // =========================================================

        private void ExecuteNonQuery(string query, (string, string)[] parameters)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(query, conn))
                {
                    foreach (var (key, value) in parameters)
                        cmd.Parameters.AddWithValue(key, value ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private string GetDateFilter(string filter, string dateColumn = "Zeitstempel")
        {
            switch (filter)
            {
                case "Heute": return $"WHERE DATE({dateColumn}) = DATE('now')";
                case "Diese Woche": return $"WHERE DATE({dateColumn}) >= DATE('now', '-7 days')";
                case "Dieser Monat": return $"WHERE DATE({dateColumn}) >= DATE('now', 'start of month')";
                case "Dieses Jahr": return $"WHERE DATE({dateColumn}) >= DATE('now', 'start of year')";
                default: return "";
            }
        }
    }
}