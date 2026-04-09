using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;

namespace NmapInventory
{
    public class DatabaseManager
    {
        private static readonly string APP_DATA_DIR = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Inventarisierung");

        private static readonly string DB_PATH = System.IO.Path.Combine(APP_DATA_DIR, "nmap_inventory.db");

        // Per-customer DB files will be named: nmap_customer_{customerId}.db
        private string GetCustomerDbPath(int customerId)
            => System.IO.Path.Combine(APP_DATA_DIR, $"nmap_customer_{customerId}.db");

        private void EnsureCustomerDatabaseExists(int customerId)
        {
            var path = GetCustomerDbPath(customerId);
            using (var conn = new SQLiteConnection($"Data Source={path};Version=3;"))
            {
                conn.Open();
                string create = @"
                    CREATE TABLE IF NOT EXISTS Devices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        IP TEXT UNIQUE NOT NULL,
                        Hostname TEXT,
                        MacAddress TEXT UNIQUE,
                        StandortID INTEGER,
                        CustomHostname INTEGER DEFAULT 0,
                        DeviceType INTEGER DEFAULT 0,
                        Vendor TEXT DEFAULT '',
                        OS TEXT,
                        Comment TEXT DEFAULT '',
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

                    CREATE TABLE IF NOT EXISTS DeviceNmapDetails (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        OS TEXT,
                        OSDetails TEXT,
                        Vendor TEXT,
                        Ports TEXT,
                        PortsJson TEXT,
                        RawOutput TEXT,
                        ScanTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS DeviceHardwareInfo (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        HardwareText TEXT,
                        QueryTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS DevicePorts (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Port INTEGER NOT NULL,
                        Protocol TEXT,
                        State TEXT,
                        Service TEXT,
                        Version TEXT,
                        Banner TEXT,
                        ScanTime DATETIME DEFAULT CURRENT_TIMESTAMP,
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

                    CREATE INDEX IF NOT EXISTS idx_device_ports ON DevicePorts(DeviceID);

                ";
                using (var cmd = new SQLiteCommand(create, conn)) cmd.ExecuteNonQuery();

                // Migrations: fehlende Spalten nachrüsten
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN CustomHostname INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Comment TEXT DEFAULT ''");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Vendor TEXT DEFAULT ''");
                TryAlterTable(conn, "ALTER TABLE DeviceSoftware ADD COLUMN PCName TEXT");
                TryAlterTable(conn, "ALTER TABLE DeviceSoftware ADD COLUMN Timestamp DATETIME");
                TryAlterTable(conn, "ALTER TABLE DeviceSoftware ADD COLUMN LastUpdate TEXT");
                TryAlterTable(conn, "ALTER TABLE DeviceSoftware ADD COLUMN DeviceID INTEGER");
                TryAlterTable(conn, "ALTER TABLE DeviceHardwareInfo ADD COLUMN IPAddress TEXT");
            }
        }

        /// <summary>
        /// Bringt eine bestehende Kunden-DB auf den aktuellen Stand (fehlende Spalten + Tabellen).
        /// Kann jederzeit sicher aufgerufen werden — bereits vorhandene Objekte werden übersprungen.
        /// </summary>
        public void MigrateCustomerDatabase(int customerId)
        {
            EnsureCustomerDatabaseExists(customerId);
        }

        /// <summary>
        /// Gibt den Dateipfad der Haupt-Datenbank zurück.
        /// </summary>
        public string GetMainDatabasePath() => DB_PATH;

        /// <summary>
        /// Gibt den Dateipfad einer Kunden-Datenbank zurück.
        /// </summary>
        public string GetCustomerDatabasePath(int customerId) => GetCustomerDbPath(customerId);

        // Public wrapper so other parts of the app can request syncing a device into the per-customer DB
        public void SyncDeviceToCustomerDb(DeviceInfo dev, int customerId)
        {
            if (dev == null) return;
            AddOrUpdateDeviceInCustomerDb(dev, customerId);
        }

        /// <summary>
        /// Aggregiert Geräte-Informationen aus allen per-Kunden Datenbanken.
        /// Liefert pro Kunde ein paar Kennzahlen (Anzahl Geräte, Geräte mit MAC, RDP/SSH Vorkommen).
        /// </summary>
        public List<(int CustomerID, string CustomerName, int DeviceCount, int DevicesWithMac, int RdpDevices, int SshDevices)> GetPivotDeviceSummary()
        {
            var result = new List<(int, string, int, int, int, int)>();
            var customers = GetCustomers();
            foreach (var c in customers)
            {
                int deviceCount = 0, devicesWithMac = 0, rdp = 0, ssh = 0;
                try
                {
                    var path = GetCustomerDbPath(c.ID);
                    if (!System.IO.File.Exists(path))
                    {
                        result.Add((c.ID, c.Name, 0, 0, 0, 0));
                        continue;
                    }

                    using (var conn = new SQLiteConnection($"Data Source={path};Version=3;"))
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM Devices", conn))
                        {
                            var r = cmd.ExecuteScalar(); deviceCount = r != null ? Convert.ToInt32(r) : 0;
                        }
                        using (var cmd = new SQLiteCommand("SELECT COUNT(*) FROM Devices WHERE MacAddress IS NOT NULL AND MacAddress != ''", conn))
                        {
                            var r = cmd.ExecuteScalar(); devicesWithMac = r != null ? Convert.ToInt32(r) : 0;
                        }
                        using (var cmd = new SQLiteCommand(@"SELECT COUNT(DISTINCT DeviceID) FROM DevicePorts WHERE Port = 3389", conn))
                        {
                            var r = cmd.ExecuteScalar(); rdp = r != null ? Convert.ToInt32(r) : 0;
                        }
                        using (var cmd = new SQLiteCommand(@"SELECT COUNT(DISTINCT DeviceID) FROM DevicePorts WHERE Port = 22", conn))
                        {
                            var r = cmd.ExecuteScalar(); ssh = r != null ? Convert.ToInt32(r) : 0;
                        }
                    }
                }
                catch { /* ignore per-customer DB errors */ }
                result.Add((c.ID, c.Name, deviceCount, devicesWithMac, rdp, ssh));
            }
            return result;
        }

        /// <summary>
        /// Lädt Gerätezeilen aus allen per-Kunden Datenbanken.
        /// Unterstützte Filter: "All" (keine), "HasMac" (nur Geräte mit MAC), "Port" (filterValue = Portnummer als string)
        /// Liefert Tuple: CustomerID, CustomerName, IP, Hostname, MacAddress, Ports (kommagetrennt), HasRdp, HasSsh
        /// </summary>
        public List<(int CustomerID, string CustomerName, string IP, string Hostname, string MacAddress, string Ports, bool HasRdp, bool HasSsh, string HardwareText, string SoftwareNames)> GetDevicesAcrossCustomers(string filterType = "All", string filterValue = null)
        {
            var rows = new List<(int, string, string, string, string, string, bool, bool, string, string)>();
            var customers = GetCustomers();
            foreach (var c in customers)
            {
                try
                {
                    var path = GetCustomerDbPath(c.ID);
                    if (!System.IO.File.Exists(path)) continue;

                    using (var conn = new SQLiteConnection($"Data Source={path};Version=3;"))
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand(@"SELECT IP, Hostname, MacAddress, ID FROM Devices", conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string ip = reader["IP"]?.ToString();
                                string hostname = reader["Hostname"]?.ToString();
                                string mac = reader["MacAddress"]?.ToString();
                                int deviceId = reader["ID"] != null ? Convert.ToInt32(reader["ID"]) : -1;

                                // Load ports for this device
                                var portList = new List<int>();
                                using (var pcmd = new SQLiteCommand("SELECT Port FROM DevicePorts WHERE DeviceID=@ID", conn))
                                {
                                    pcmd.Parameters.AddWithValue("@ID", deviceId);
                                    using (var pr = pcmd.ExecuteReader())
                                        while (pr.Read()) portList.Add(Convert.ToInt32(pr["Port"]));
                                }

                                bool hasRdp = portList.Contains(3389);
                                bool hasSsh = portList.Contains(22);
                                string ports = portList.Count > 0 ? string.Join(",", portList) : "";

                                // Load latest hardware text (if any)
                                string hwText = "";
                                using (var hcmd = new SQLiteCommand("SELECT HardwareText FROM DeviceHardwareInfo WHERE DeviceID=@ID ORDER BY QueryTime DESC LIMIT 1", conn))
                                {
                                    hcmd.Parameters.AddWithValue("@ID", deviceId);
                                    var hres = hcmd.ExecuteScalar();
                                    if (hres != null) hwText = hres.ToString();
                                }

                                // Load installed software names (comma separated)
                                var swNames = new List<string>();
                                using (var scmd = new SQLiteCommand("SELECT Name FROM DeviceSoftware WHERE DeviceID=@ID", conn))
                                {
                                    scmd.Parameters.AddWithValue("@ID", deviceId);
                                    using (var sr = scmd.ExecuteReader())
                                        while (sr.Read()) swNames.Add(sr["Name"]?.ToString() ?? "");
                                }
                                string softwareNames = swNames.Count > 0 ? string.Join(",", swNames.Where(s => !string.IsNullOrEmpty(s))) : "";

                                // Apply filter
                                bool include = true;
                                if (filterType == "HasMac") include = !string.IsNullOrWhiteSpace(mac);
                                else if (filterType == "Port")
                                {
                                    if (int.TryParse(filterValue, out int p)) include = portList.Contains(p);
                                    else include = false;
                                }

                                if (include)
                                    rows.Add((c.ID, c.Name, ip, hostname, mac, ports, hasRdp, hasSsh, hwText, softwareNames));
                            }
                        }
                    }
                }
                catch { /* ignore per-customer DB errors */ }
            }
            return rows;
        }

        /// <summary>
        /// Führt eine SQL-Abfrage (WHERE-Klausel) auf der per-customer DB aus und liefert Gerätedaten zurück.
        /// </summary>
        public List<(string IP, string Hostname, string MacAddress, string Ports, bool HasRdp, bool HasSsh, string HardwareText, string SoftwareNames)> RunDeviceQueryOnCustomerDb(int customerId, string whereClause, Dictionary<string, object> parameters = null)
        {
            var path = GetCustomerDbPath(customerId);
            // If missing, fall back to central
            if (!System.IO.File.Exists(path)) path = DB_PATH;
            return RunDeviceQueryOnPath(path, whereClause, parameters);
        }

        /// <summary>
        /// Führt die gleiche Abfrage gegen die zentrale DB aus.
        /// </summary>
        public List<(string IP, string Hostname, string MacAddress, string Ports, bool HasRdp, bool HasSsh, string HardwareText, string SoftwareNames)> RunDeviceQueryOnCentralDb(string whereClause, Dictionary<string, object> parameters = null)
        {
            return RunDeviceQueryOnPath(DB_PATH, whereClause, parameters);
        }

        public List<(string IP, string Hostname, string MacAddress, string Ports, bool HasRdp, bool HasSsh, string HardwareText, string SoftwareNames)> RunDeviceQueryOnPath(string path, string whereClause, Dictionary<string, object> parameters = null)
        {
            var results = new List<(string, string, string, string, bool, bool, string, string)>();
            if (!System.IO.File.Exists(path)) return results;

            using (var conn = new SQLiteConnection($"Data Source={path};Version=3;"))
            {
                conn.Open();
                string sql = @"SELECT d.IP, d.Hostname, d.MacAddress,
                    (SELECT group_concat(Port, ',') FROM DevicePorts WHERE DeviceID=d.ID) AS Ports,
                    EXISTS(SELECT 1 FROM DevicePorts WHERE DeviceID=d.ID AND Port=3389) AS HasRdp,
                    EXISTS(SELECT 1 FROM DevicePorts WHERE DeviceID=d.ID AND Port=22) AS HasSsh,
                    (SELECT HardwareText FROM DeviceHardwareInfo WHERE DeviceID=d.ID ORDER BY QueryTime DESC LIMIT 1) AS HardwareText,
                    (SELECT group_concat(Name, ',') FROM DeviceSoftware WHERE DeviceID=d.ID) AS SoftwareNames
                  FROM Devices d";

                if (!string.IsNullOrWhiteSpace(whereClause)) sql += " WHERE " + whereClause;

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    if (parameters != null)
                        foreach (var kv in parameters)
                            cmd.Parameters.AddWithValue(kv.Key, kv.Value ?? "");

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string ip = reader["IP"]?.ToString();
                            string hostname = reader["Hostname"]?.ToString();
                            string mac = reader["MacAddress"]?.ToString();
                            string ports = reader["Ports"]?.ToString() ?? "";
                            bool hasRdp = Convert.ToInt32(reader["HasRdp"] ?? 0) == 1;
                            bool hasSsh = Convert.ToInt32(reader["HasSsh"] ?? 0) == 1;
                            string hw = reader["HardwareText"]?.ToString() ?? "";
                            string sw = reader["SoftwareNames"]?.ToString() ?? "";
                            results.Add((ip, hostname, mac, ports, hasRdp, hasSsh, hw, sw));
                        }
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// Findet alle nmap_*.db Dateien im Anwendungsverzeichnis (rekursiv) und gibt die Pfade zurück.
        /// </summary>
        public List<string> GetAllDatabaseFiles()
        {
            try
            {
                return System.IO.Directory.GetFiles(APP_DATA_DIR, "nmap_*.db").ToList();
            }
            catch { return new List<string>(); }
        }

        /// <summary>
        /// Extracts customerId from filename if pattern matches nmap_customer_{id}.db
        /// </summary>
        public int? TryGetCustomerIdFromPath(string path)
        {
            try
            {
                var name = System.IO.Path.GetFileNameWithoutExtension(path);
                if (name.StartsWith("nmap_customer_"))
                {
                    var part = name.Substring("nmap_customer_".Length);
                    if (int.TryParse(part.Split('_')[0], out int id)) return id;
                }
            }
            catch { }
            return null;
        }

        private void AddOrUpdateDeviceInCustomerDb(DeviceInfo dev, int customerId)
        {
            try
            {
                EnsureCustomerDatabaseExists(customerId);
                var path = GetCustomerDbPath(customerId);
                using (var conn = new SQLiteConnection($"Data Source={path};Version=3;"))
                {
                    conn.Open();
                    // Try find by IP
                    int existingID = -1;
                    using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP = @IP", conn))
                    {
                        cmd.Parameters.AddWithValue("@IP", dev.IP);
                        var r = cmd.ExecuteScalar();
                        if (r != null) existingID = Convert.ToInt32(r);
                    }

                    // If not found and MAC present, try by MAC
                    if (existingID < 0 && !string.IsNullOrWhiteSpace(dev.MacAddress))
                    {
                        using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE MacAddress = @MAC", conn))
                        {
                            cmd.Parameters.AddWithValue("@MAC", dev.MacAddress.Trim().ToUpperInvariant());
                            var r = cmd.ExecuteScalar();
                            if (r != null) existingID = Convert.ToInt32(r);
                        }
                    }

                    if (existingID > 0)
                    {
                        using (var cmd = new SQLiteCommand(@"UPDATE Devices SET Hostname=@Hostname, MacAddress=@MAC, IP=@IP, LastSeen=CURRENT_TIMESTAMP WHERE ID=@ID", conn))
                        {
                            cmd.Parameters.AddWithValue("@ID", existingID);
                            cmd.Parameters.AddWithValue("@IP", dev.IP);
                            cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                            cmd.Parameters.AddWithValue("@MAC", string.IsNullOrWhiteSpace(dev.MacAddress) ? (object)DBNull.Value : dev.MacAddress.Trim().ToUpperInvariant());
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SQLiteCommand("INSERT INTO Devices (IP, Hostname, MacAddress) VALUES (@IP, @Hostname, @MAC)", conn))
                        {
                            cmd.Parameters.AddWithValue("@IP", dev.IP);
                            cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                            cmd.Parameters.AddWithValue("@MAC", string.IsNullOrWhiteSpace(dev.MacAddress) ? (object)DBNull.Value : dev.MacAddress.Trim().ToUpperInvariant());
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch { /* non-fatal */ }
        }

        /// <summary>
        /// Importiert Kunden, Standorte und Geräte aus einer alten DB-Datei in die aktuelle DB.
        /// Bereits vorhandene Einträge (gleicher Name/IP) werden übersprungen.
        /// Gibt die Anzahl importierter Kunden, Standorte und Geräte zurück.
        /// </summary>
        public bool LegacyDbHasTable(string dbPath, string tableName)
        {
            try
            {
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmd = new SQLiteCommand(
                        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@T", conn))
                    {
                        cmd.Parameters.AddWithValue("@T", tableName);
                        return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
                    }
                }
            }
            catch { return false; }
        }

        public (int Kunden, int Standorte, int Geraete) ImportFromLegacyDatabase(string sourcePath, int? fallbackCustomerId = null)
        {
            int kunden = 0, standorte = 0, geraete = 0;
            using (var src  = new SQLiteConnection($"Data Source={sourcePath};Version=3;"))
            using (var dest = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                src.Open();
                dest.Open();

                // ── Kunden ────────────────────────────────────
                var customerMap = new Dictionary<int, int>(); // alte ID → neue ID
                bool hasCustomers;
                using (var chk = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='Customers'", src))
                    hasCustomers = Convert.ToInt32(chk.ExecuteScalar()) > 0;

                if (hasCustomers)
                    using (var cmd = new SQLiteCommand("SELECT ID, Name, Address FROM Customers ORDER BY ID", src))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            int  oldId  = Convert.ToInt32(r["ID"]);
                            string name = r["Name"]?.ToString() ?? "";
                            string addr = r["Address"]?.ToString() ?? "";
                            int newId;
                            using (var chk = new SQLiteCommand("SELECT ID FROM Customers WHERE Name=@N", dest))
                            {
                                chk.Parameters.AddWithValue("@N", name);
                                var ex = chk.ExecuteScalar();
                                if (ex != null) { newId = Convert.ToInt32(ex); }
                                else
                                {
                                    using (var ins = new SQLiteCommand(
                                        "INSERT INTO Customers (Name, Address) VALUES (@N, @A)", dest))
                                    {
                                        ins.Parameters.AddWithValue("@N", name);
                                        ins.Parameters.AddWithValue("@A", addr);
                                        ins.ExecuteNonQuery();
                                    }
                                    newId = (int)dest.LastInsertRowId;
                                    kunden++;
                                }
                            }
                            customerMap[oldId] = newId;
                        }

                // ── Standorte ─────────────────────────────────
                var locationMap = new Dictionary<int, int>(); // alte ID → neue ID
                bool hasLocations;
                using (var chk = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='Locations'", src))
                    hasLocations = Convert.ToInt32(chk.ExecuteScalar()) > 0;

                if (hasLocations)
                    using (var cmd = new SQLiteCommand("SELECT ID, CustomerID, Name, Address, Level, ParentID FROM Locations ORDER BY Level, ID", src))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            int oldId     = Convert.ToInt32(r["ID"]);
                            int oldCustId = r["CustomerID"] != DBNull.Value ? Convert.ToInt32(r["CustomerID"]) : 0;
                            string name   = r["Name"]?.ToString() ?? "";
                            string addr   = r["Address"]?.ToString() ?? "";
                            int level     = r["Level"] != DBNull.Value ? Convert.ToInt32(r["Level"]) : 0;
                            int oldParent = r["ParentID"] != DBNull.Value ? Convert.ToInt32(r["ParentID"]) : 0;
                            // Kunden-ID: aus Mapping, Fallback, oder direkt übernehmen
                            int newCustId = customerMap.TryGetValue(oldCustId, out int mc) ? mc
                                          : fallbackCustomerId.HasValue ? fallbackCustomerId.Value
                                          : oldCustId;
                            int newParent = locationMap.TryGetValue(oldParent, out int mp) ? mp : 0;

                            int newId;
                            using (var chk = new SQLiteCommand(
                                "SELECT ID FROM Locations WHERE CustomerID=@C AND Name=@N", dest))
                            {
                                chk.Parameters.AddWithValue("@C", newCustId);
                                chk.Parameters.AddWithValue("@N", name);
                                var ex = chk.ExecuteScalar();
                                if (ex != null) { newId = Convert.ToInt32(ex); }
                                else
                                {
                                    using (var ins = new SQLiteCommand(
                                        "INSERT INTO Locations (CustomerID, Name, Address, Level, ParentID) VALUES (@C,@N,@A,@L,@P)", dest))
                                    {
                                        ins.Parameters.AddWithValue("@C", newCustId);
                                        ins.Parameters.AddWithValue("@N", name);
                                        ins.Parameters.AddWithValue("@A", addr);
                                        ins.Parameters.AddWithValue("@L", level);
                                        ins.Parameters.AddWithValue("@P", newParent > 0 ? (object)newParent : DBNull.Value);
                                        ins.ExecuteNonQuery();
                                    }
                                    newId = (int)dest.LastInsertRowId;
                                    standorte++;
                                }
                            }
                            locationMap[oldId] = newId;
                        }

                // ── Geräte ────────────────────────────────────
                bool hasDevices;
                using (var chk = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='Devices'", src))
                    hasDevices = Convert.ToInt32(chk.ExecuteScalar()) > 0;

                if (hasDevices)
                    using (var cmd = new SQLiteCommand("SELECT IP, Hostname, MacAddress, DeviceType, Vendor, OS FROM Devices", src))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            string ip  = r["IP"]?.ToString() ?? "";
                            if (string.IsNullOrEmpty(ip)) continue;
                            using (var ins = new SQLiteCommand(@"
                                INSERT OR IGNORE INTO Devices (IP, Hostname, MacAddress, DeviceType, Vendor, OS)
                                VALUES (@IP,@H,@MAC,@DT,@V,@OS)", dest))
                            {
                                ins.Parameters.AddWithValue("@IP",  ip);
                                ins.Parameters.AddWithValue("@H",   r["Hostname"]?.ToString() ?? "");
                                ins.Parameters.AddWithValue("@MAC", r["MacAddress"]?.ToString() ?? "");
                                ins.Parameters.AddWithValue("@DT",  r["DeviceType"] != DBNull.Value ? r["DeviceType"] : 0);
                                ins.Parameters.AddWithValue("@V",   r["Vendor"]?.ToString() ?? "");
                                ins.Parameters.AddWithValue("@OS",  r["OS"]?.ToString() ?? "");
                                if (ins.ExecuteNonQuery() > 0) geraete++;
                            }
                        }

                // ── LocationIPs ───────────────────────────────
                bool hasLocationIPs;
                using (var chk = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='LocationIPs'", src))
                    hasLocationIPs = Convert.ToInt32(chk.ExecuteScalar()) > 0;

                if (hasLocationIPs && locationMap.Count > 0)
                    using (var cmd = new SQLiteCommand("SELECT LocationID, IPAddress, WorkstationName FROM LocationIPs", src))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            int oldLoc = Convert.ToInt32(r["LocationID"]);
                            if (!locationMap.TryGetValue(oldLoc, out int newLoc)) continue;
                            using (var ins = new SQLiteCommand(@"
                                INSERT OR IGNORE INTO LocationIPs (LocationID, IPAddress, WorkstationName)
                                VALUES (@L, @IP, @WS)", dest))
                            {
                                ins.Parameters.AddWithValue("@L",  newLoc);
                                ins.Parameters.AddWithValue("@IP", r["IPAddress"]?.ToString() ?? "");
                                ins.Parameters.AddWithValue("@WS", r["WorkstationName"]?.ToString() ?? "");
                                ins.ExecuteNonQuery();
                            }
                        }
            }
            return (kunden, standorte, geraete);
        }

        public void InitializeDatabase()
        {
            System.IO.Directory.CreateDirectory(APP_DATA_DIR);
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Devices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        IP TEXT UNIQUE NOT NULL,
                        Hostname TEXT,
                        MacAddress TEXT UNIQUE,
                        StandortID INTEGER,
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
                    );

                    -- ═══════════════════════════════════════════════════
                    -- NORMALISIERTES GERÄTEMODELL (nach ER-Diagramm)
                    -- ═══════════════════════════════════════════════════

                    CREATE TABLE IF NOT EXISTS Ip_Type (
                        id_ip_type    INTEGER PRIMARY KEY AUTOINCREMENT,
                        ip_type_description VARCHAR(255) NOT NULL UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS Ip_Version (
                        id_ip_version INTEGER PRIMARY KEY AUTOINCREMENT,
                        ip_version_description VARCHAR(50) NOT NULL UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS Network_Type (
                        id_network_type INTEGER PRIMARY KEY AUTOINCREMENT,
                        network_type_description VARCHAR(50) NOT NULL UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS Software_Type (
                        id_software_type INTEGER PRIMARY KEY AUTOINCREMENT,
                        software_type_description VARCHAR(50) NOT NULL UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS Hardware_Type (
                        id_hardware_type INTEGER PRIMARY KEY AUTOINCREMENT,
                        hardware_type_description VARCHAR(50) NOT NULL UNIQUE
                    );

                    CREATE TABLE IF NOT EXISTS Device (
                        id_device                       INTEGER PRIMARY KEY AUTOINCREMENT,
                        device_name                     VARCHAR(255),
                        device_operatingsystem          VARCHAR(50),
                        device_operatingsystem_version  VARCHAR(50),
                        device_last_systemscan          TIMESTAMP,
                        device_last_hardwarescan        TIMESTAMP,
                        device_last_softwarescan        TIMESTAMP
                    );

                    CREATE TABLE IF NOT EXISTS Network (
                        id_network          INTEGER PRIMARY KEY AUTOINCREMENT,
                        network_name        VARCHAR(255),
                        network_description VARCHAR(45),
                        network_subnet      VARCHAR(50),
                        id_network_type     INTEGER,
                        FOREIGN KEY(id_network_type) REFERENCES Network_Type(id_network_type)
                    );

                    CREATE TABLE IF NOT EXISTS Device_Network (
                        id_device_network INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_device         INTEGER NOT NULL,
                        id_network        INTEGER,
                        FOREIGN KEY(id_device)  REFERENCES Device(id_device)  ON DELETE CASCADE,
                        FOREIGN KEY(id_network) REFERENCES Network(id_network) ON DELETE SET NULL
                    );

                    CREATE TABLE IF NOT EXISTS Device_Mac_Address (
                        id_mac_address    INTEGER PRIMARY KEY AUTOINCREMENT,
                        mac_address       VARCHAR(50) NOT NULL,
                        id_device_network INTEGER NOT NULL,
                        FOREIGN KEY(id_device_network) REFERENCES Device_Network(id_device_network) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS Device_ip_address (
                        id_device_ip_address INTEGER PRIMARY KEY AUTOINCREMENT,
                        ip_address           VARCHAR(255) NOT NULL,
                        id_ip_type           INTEGER,
                        id_ip_version        INTEGER,
                        id_device_network    INTEGER NOT NULL,
                        FOREIGN KEY(id_ip_type)        REFERENCES Ip_Type(id_ip_type),
                        FOREIGN KEY(id_ip_version)     REFERENCES Ip_Version(id_ip_version),
                        FOREIGN KEY(id_device_network) REFERENCES Device_Network(id_device_network) ON DELETE CASCADE
                    );

                    CREATE TABLE IF NOT EXISTS Software (
                        id_software       INTEGER PRIMARY KEY AUTOINCREMENT,
                        software_name     VARCHAR(50) NOT NULL,
                        software_publisher VARCHAR(50),
                        software_size     FLOAT,
                        software_version  VARCHAR(50),
                        id_software_type  INTEGER,
                        FOREIGN KEY(id_software_type) REFERENCES Software_Type(id_software_type)
                    );

                    CREATE TABLE IF NOT EXISTS Device_Software (
                        id_device_software          INTEGER PRIMARY KEY AUTOINCREMENT,
                        software_installation_date  TIMESTAMP,
                        id_device                   INTEGER NOT NULL,
                        id_software                 INTEGER NOT NULL,
                        FOREIGN KEY(id_device)  REFERENCES Device(id_device)   ON DELETE CASCADE,
                        FOREIGN KEY(id_software) REFERENCES Software(id_software)
                    );

                    CREATE TABLE IF NOT EXISTS Hardware (
                        id_hardware              INTEGER PRIMARY KEY AUTOINCREMENT,
                        hardware_name            VARCHAR(50),
                        hardware_manufacturer    VARCHAR(50),
                        hardware_description     VARCHAR(50),
                        id_hardware_type         INTEGER,
                        FOREIGN KEY(id_hardware_type) REFERENCES Hardware_Type(id_hardware_type)
                    );

                    CREATE TABLE IF NOT EXISTS Device_Hardware (
                        id_device_hardware INTEGER PRIMARY KEY AUTOINCREMENT,
                        id_device          INTEGER NOT NULL,
                        id_hardware        INTEGER NOT NULL,
                        FOREIGN KEY(id_device)  REFERENCES Device(id_device)   ON DELETE CASCADE,
                        FOREIGN KEY(id_hardware) REFERENCES Hardware(id_hardware)
                    );

                    -- Indizes für das normalisierte Modell
                    CREATE INDEX IF NOT EXISTS idx_dn_device   ON Device_Network(id_device);
                    CREATE INDEX IF NOT EXISTS idx_dn_network  ON Device_Network(id_network);
                    CREATE INDEX IF NOT EXISTS idx_dma_dn      ON Device_Mac_Address(id_device_network);
                    CREATE INDEX IF NOT EXISTS idx_dia_dn      ON Device_ip_address(id_device_network);
                    CREATE INDEX IF NOT EXISTS idx_ds_device   ON Device_Software(id_device);
                    CREATE INDEX IF NOT EXISTS idx_dh_device   ON Device_Hardware(id_device);
                ";

                using (var cmd = new SQLiteCommand(createTableQuery, conn))
                    cmd.ExecuteNonQuery();

                TryAlterTable(conn, "ALTER TABLE DeviceSoftware ADD COLUMN PCName TEXT");
                TryAlterTable(conn, "ALTER TABLE DeviceHardwareInfo ADD COLUMN IPAddress TEXT");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN MacAddress TEXT UNIQUE");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN CustomHostname INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN DeviceType INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Vendor TEXT");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN OS TEXT");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Comment TEXT DEFAULT ''");
                TryAlterTable(conn, "ALTER TABLE Customers ADD COLUMN CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN Level INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN ParentID INTEGER");
                TryAlterTable(conn, "ALTER TABLE LocationIPs ADD COLUMN AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlterTable(conn, "ALTER TABLE LocationIPs ADD COLUMN DeviceID INTEGER REFERENCES Devices(ID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_mac_address ON Devices(MacAddress)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_device_mac_history ON DeviceMacHistory(MacAddress)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_parent ON Locations(ParentID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_customer ON Locations(CustomerID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_devices ON LocationDevices(LocationID)");
                TryAlterTable(conn, "CREATE INDEX IF NOT EXISTS idx_location_ips_device ON LocationIPs(DeviceID)");
                TryAlterTable(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_location_ips_unique ON LocationIPs(LocationID, IPAddress)");

                // Migration: StandortID auf Devices setzen (aus LocationDevices ableiten)
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN StandortID INTEGER");
                TryAlterTable(conn, @"
                    UPDATE Devices SET StandortID = (
                        SELECT ld.LocationID FROM LocationDevices ld WHERE ld.DeviceID = Devices.ID LIMIT 1
                    ) WHERE StandortID IS NULL AND EXISTS (SELECT 1 FROM LocationDevices ld WHERE ld.DeviceID = Devices.ID)");

                // UNIQUE-Index auf DeviceSoftware (verhindert Duplikate)
                TryAlterTable(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_device_software_unique ON DeviceSoftware(DeviceID, Name)");

                // ── Inventar_* als VIEWs (immer live, kein Refresh nötig) ──
                // Alte Views entfernen, dann frische Views anlegen
                // WICHTIG: Jedes Statement einzeln ausführen, da SQLite sonst einen Fehler wirft
                // Zuerst sicherstellen, dass Views (nicht Tables) gelöscht werden
                TryDropViewSafely(conn, "Inventar_Geraete");
                TryDropViewSafely(conn, "Inventar_Software");
                TryDropViewSafely(conn, "Inventar_Ports");

                using (var viewCmd = new SQLiteCommand(conn))
                {
                    // Create Inventar_Geraete
                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Geraete AS
                        SELECT
                            c.ID  AS KundeID,     c.Name AS KundeName,
                            l.ID  AS StandortID,  l.Name AS StandortName,
                            d.ID  AS GeraetID,    d.IP,  d.Hostname,  d.MacAddress,
                            d.DeviceType AS GeraetTyp,
                            CASE d.DeviceType
                                WHEN 1  THEN 'Windows PC'    WHEN 2  THEN 'Windows Server'
                                WHEN 3  THEN 'Linux'         WHEN 4  THEN 'Drucker'
                                WHEN 5  THEN 'Smartphone'    WHEN 6  THEN 'Tablet'
                                WHEN 7  THEN 'Netzwerkgerät' WHEN 8  THEN 'NAS'
                                WHEN 9  THEN 'Smart TV'      WHEN 10 THEN 'IoT-Gerät'
                                WHEN 11 THEN 'macOS'         WHEN 12 THEN 'Laptop'
                                ELSE 'Unbekannt'
                            END AS GeraetTypName,
                            d.Vendor, d.OS, d.Comment,
                            d.FirstSeen AS ErsterScan, d.LastSeen AS LetzterScan,
                            (SELECT COUNT(*) FROM DevicePorts    WHERE DeviceID = d.ID) AS AnzahlPorts,
                            (SELECT COUNT(*) FROM DeviceSoftware WHERE DeviceID = d.ID) AS AnzahlSoftware
                        FROM Devices d
                        LEFT JOIN Locations l  ON l.ID = d.StandortID
                        LEFT JOIN Customers c  ON c.ID = l.CustomerID;";
                    viewCmd.ExecuteNonQuery();

                    // Create Inventar_Software
                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Software AS
                        SELECT
                            c.ID  AS KundeID,     c.Name AS KundeName,
                            l.ID  AS StandortID,  l.Name AS StandortName,
                            d.ID  AS GeraetID,    d.IP,  d.Hostname,  d.MacAddress,
                            d.DeviceType AS GeraetTyp,
                            CASE d.DeviceType
                                WHEN 1  THEN 'Windows PC'    WHEN 2  THEN 'Windows Server'
                                WHEN 3  THEN 'Linux'         WHEN 4  THEN 'Drucker'
                                WHEN 5  THEN 'Smartphone'    WHEN 6  THEN 'Tablet'
                                WHEN 7  THEN 'Netzwerkgerät' WHEN 8  THEN 'NAS'
                                WHEN 9  THEN 'Smart TV'      WHEN 10 THEN 'IoT-Gerät'
                                WHEN 11 THEN 'macOS'         WHEN 12 THEN 'Laptop'
                                ELSE 'Unbekannt'
                            END AS GeraetTypName,
                            sw.ID AS SoftwareID, sw.Name AS SoftwareName,
                            sw.Version AS SoftwareVersion,
                            sw.Publisher AS Hersteller, sw.InstallDate AS InstallDatum,
                            sw.Source AS Quelle
                        FROM DeviceSoftware sw
                        JOIN Devices d         ON d.ID = sw.DeviceID
                        LEFT JOIN Locations l  ON l.ID = d.StandortID
                        LEFT JOIN Customers c  ON c.ID = l.CustomerID;";
                    viewCmd.ExecuteNonQuery();

                    // Create Inventar_Ports
                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Ports AS
                        SELECT
                            c.ID  AS KundeID,     c.Name AS KundeName,
                            l.ID  AS StandortID,  l.Name AS StandortName,
                            d.ID  AS GeraetID,    d.IP,  d.Hostname,
                            d.DeviceType AS GeraetTyp,
                            CASE d.DeviceType
                                WHEN 1  THEN 'Windows PC'    WHEN 2  THEN 'Windows Server'
                                WHEN 3  THEN 'Linux'         WHEN 4  THEN 'Drucker'
                                WHEN 5  THEN 'Smartphone'    WHEN 6  THEN 'Tablet'
                                WHEN 7  THEN 'Netzwerkgerät' WHEN 8  THEN 'NAS'
                                WHEN 9  THEN 'Smart TV'      WHEN 10 THEN 'IoT-Gerät'
                                WHEN 11 THEN 'macOS'         WHEN 12 THEN 'Laptop'
                                ELSE 'Unbekannt'
                            END AS GeraetTypName,
                            p.ID AS PortID, p.Port, p.Protocol AS Protokoll,
                            p.State AS Status, p.Service AS Dienst, p.Version
                        FROM DevicePorts p
                        JOIN Devices d         ON d.ID = p.DeviceID
                        LEFT JOIN Locations l  ON l.ID = d.StandortID
                        LEFT JOIN Customers c  ON c.ID = l.CustomerID;";
                    viewCmd.ExecuteNonQuery();
                }

                // Grunddaten für normalisiertes Modell (idempotent via INSERT OR IGNORE)
                using (var cmd2 = new SQLiteCommand(conn))
                {
                    cmd2.CommandText = @"
                        INSERT OR IGNORE INTO Ip_Version (ip_version_description) VALUES ('IPv4'), ('IPv6');
                        INSERT OR IGNORE INTO Ip_Type    (ip_type_description)    VALUES ('Statisch'), ('Dynamisch (DHCP)'), ('Link-Local'), ('Loopback');
                        INSERT OR IGNORE INTO Network_Type (network_type_description) VALUES ('LAN'), ('WLAN'), ('WAN'), ('VPN'), ('DMZ');
                        INSERT OR IGNORE INTO Software_Type (software_type_description) VALUES ('Betriebssystem'), ('Anwendung'), ('Treiber'), ('Sicherheit'), ('Entwicklungswerkzeug'), ('Sonstiges');
                        INSERT OR IGNORE INTO Hardware_Type (hardware_type_description) VALUES ('CPU'), ('RAM'), ('Festplatte'), ('Grafikkarte'), ('Netzwerkkarte'), ('Motherboard'), ('Netzteil'), ('Monitor'), ('Peripherie'), ('Sonstiges');
                    ";
                    cmd2.ExecuteNonQuery();
                }

                // Migration: LocationIPs.DeviceID anhand IP-Adresse befüllen
                TryAlterTable(conn, @"
                    UPDATE LocationIPs SET DeviceID = (
                        SELECT d.ID FROM Devices d WHERE d.IP = LocationIPs.IPAddress LIMIT 1
                    ) WHERE DeviceID IS NULL");

                // Migration: LocationDevices aus LocationIPs ableiten (einmalig)
                TryAlterTable(conn, @"
                    INSERT OR IGNORE INTO LocationDevices (LocationID, DeviceID)
                    SELECT LocationID, DeviceID FROM LocationIPs
                    WHERE DeviceID IS NOT NULL");

                // Migration: DeviceSoftware.DeviceID reparieren wo PCName eine IP ist
                TryAlterTable(conn, @"
                    UPDATE DeviceSoftware SET DeviceID = (
                        SELECT d.ID FROM Devices d WHERE d.IP = DeviceSoftware.PCName LIMIT 1
                    ) WHERE PCName IS NOT NULL AND PCName != ''
                      AND EXISTS (SELECT 1 FROM Devices WHERE IP = DeviceSoftware.PCName)");

                // Migration: alte "Vendor (AA:BB:CC:DD:EE:FF)" und reine MAC-Hostnamen bereinigen
                // Vendor+MAC Muster: beliebiger Text gefolgt von " (XX:XX:XX:XX:XX:XX)"
                TryAlterTable(conn, @"
                    UPDATE Devices SET Hostname = '', CustomHostname = 0
                    WHERE CustomHostname = 0
                      AND (
                        Hostname GLOB '*(*:*:*:*:*:*)*'
                        OR Hostname GLOB '[0-9A-Fa-f][0-9A-Fa-f]:[0-9A-Fa-f][0-9A-Fa-f]:*'
                        OR Hostname = IP
                      )");
                TryAlterTable(conn, @"
                    UPDATE LocationIPs SET WorkstationName = NULL
                    WHERE (
                        WorkstationName GLOB '*(*:*:*:*:*:*)*'
                        OR WorkstationName GLOB '[0-9A-Fa-f][0-9A-Fa-f]:[0-9A-Fa-f][0-9A-Fa-f]:*'
                        OR WorkstationName = IPAddress
                    )");
            }
        }

        private void TryAlterTable(SQLiteConnection conn, string sql)
        {
            try { using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery(); }
            catch { }
        }

        private void TryDropViewSafely(SQLiteConnection conn, string viewName)
        {
            try
            {
                using (var cmd = new SQLiteCommand($"DROP VIEW IF EXISTS {viewName};", conn))
                    cmd.ExecuteNonQuery();
            }
            catch
            {
                // Falls als View nicht vorhanden, versuche als Tabelle zu löschen (für alte DBs)
                try
                {
                    using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS {viewName};", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }
            }
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

                    // DeviceType nur überschreiben wenn er bisher Unbekannt war
                    // (manuell gesetzter Typ bleibt erhalten)
                    using (var cmd = new SQLiteCommand(
                        @"UPDATE Devices SET
                            DeviceType = CASE WHEN DeviceType = 0 THEN @DT ELSE DeviceType END,
                            Vendor = CASE WHEN @Vendor != '' THEN @Vendor ELSE Vendor END,
                            OS     = CASE WHEN @OS != ''     THEN @OS     ELSE OS     END
                          WHERE ID = @ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@DT", (int)dev.DeviceType);
                        cmd.Parameters.AddWithValue("@Vendor", dev.Vendor ?? "");
                        cmd.Parameters.AddWithValue("@OS", dev.OS ?? "");
                        cmd.Parameters.AddWithValue("@ID", deviceID);
                        cmd.ExecuteNonQuery();
                    }

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
            // Nur eintragen wenn sich Status oder Ports geändert haben
            string lastStatus = null, lastPorts = null;
            using (var cmd = new SQLiteCommand(
                "SELECT Status, Ports FROM DeviceScanHistory WHERE DeviceID=@ID ORDER BY ScanTime DESC LIMIT 1", conn))
            {
                cmd.Parameters.AddWithValue("@ID", deviceID);
                using (var r = cmd.ExecuteReader())
                    if (r.Read()) { lastStatus = r["Status"]?.ToString(); lastPorts = r["Ports"]?.ToString(); }
            }

            if (lastStatus == (status ?? "") && lastPorts == (ports ?? ""))
                return; // Keine Änderung — kein neuer Eintrag

            using (var cmd = new SQLiteCommand(
                "INSERT INTO DeviceScanHistory (DeviceID, Status, Ports) VALUES (@DeviceID, @Status, @Ports)", conn))
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
                    "INSERT INTO DeviceHardwareInfo (DeviceID, IPAddress, HardwareText) VALUES (@DeviceID, @IP, @Text)", conn))
                {
                    cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                    cmd.Parameters.AddWithValue("@IP", ip ?? "");
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
        /// Speichert einen Kommentar für ein Gerät (bleibt bei Scans erhalten).
        /// </summary>
        /// <summary>
        /// Speichert den Hersteller aus einem HW-Scan (WMI/SSH/ADB) — überschreibt nur wenn bisher leer.
        /// </summary>
        public void SaveVendorFromScan(string ip, string vendor)
        {
            if (string.IsNullOrWhiteSpace(vendor)) return;
            try
            {
                ExecuteNonQuery(
                    "UPDATE Devices SET Vendor=@V WHERE IP=@IP AND (Vendor IS NULL OR Vendor='')",
                    new[] { ("@V", vendor), ("@IP", ip) });
            }
            catch { }
        }

        public void SaveComment(string ip, string comment)
        {
            ExecuteNonQuery(
                "UPDATE Devices SET Comment=@C WHERE IP=@IP",
                new[] { ("@C", comment ?? ""), ("@IP", ip) });
        }

        /// <summary>
        /// Setzt den Gerätetyp manuell und markiert ihn als fixiert (wird bei Scans nicht überschrieben).
        /// </summary>
        public void SetDeviceType(string ip, DeviceType type)
        {
            ExecuteNonQuery(
                "UPDATE Devices SET DeviceType=@DT WHERE IP=@IP",
                new[] { ("@DT", ((int)type).ToString()), ("@IP", ip) });
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

        /// <summary>
        /// Setzt Hostname + WorkstationName nur wenn noch kein manueller Name vergeben wurde (CustomHostname = 0)
        /// und der aktuelle Hostname leer ist. Wird z.B. nach ADB/SSH-Scan aufgerufen.
        /// </summary>
        public void SetHostnameIfEmpty(string ip, string hostname)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(
                    @"UPDATE Devices SET Hostname=@Hostname
                      WHERE IP=@IP AND CustomHostname=0 AND (Hostname IS NULL OR Hostname='')", conn))
                {
                    cmd.Parameters.AddWithValue("@Hostname", hostname);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.ExecuteNonQuery();
                }
                // LocationIPs.WorkstationName ebenfalls setzen wenn leer
                using (var cmd = new SQLiteCommand(
                    @"UPDATE LocationIPs SET WorkstationName=@Name
                      WHERE IPAddress=@IP AND (WorkstationName IS NULL OR WorkstationName='')", conn))
                {
                    cmd.Parameters.AddWithValue("@Name", hostname);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.ExecuteNonQuery();
                }
            }
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

                // Fehlende Spalten sicher nachrüsten (falls Migration noch nicht gelaufen)
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Comment TEXT DEFAULT ''");
                TryAlterTable(conn, "ALTER TABLE Devices ADD COLUMN Vendor  TEXT DEFAULT ''");

                string whereClause = GetDateFilter(filter, "d.LastSeen");
                string query = $@"
                    SELECT DISTINCT
                        d.ID, d.IP, d.Hostname, d.MacAddress,
                        d.DeviceType, d.Vendor, d.OS, d.Comment,
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
                            Vendor = reader["Vendor"]?.ToString(),
                            Comment = reader["Comment"]?.ToString(),
                            Status = reader["Status"]?.ToString(),
                            Ports = reader["Ports"]?.ToString(),
                            DeviceType = reader["DeviceType"] != DBNull.Value
                                ? (DeviceType)Convert.ToInt32(reader["DeviceType"])
                                : DeviceType.Unbekannt
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
            if (string.IsNullOrWhiteSpace(nameOrIP)) return -1;

            // 1. Exakter Treffer auf IP oder Hostname
            using (var cmd = new SQLiteCommand(
                "SELECT ID FROM Devices WHERE IP = @Name OR Hostname = @Name LIMIT 1", conn))
            {
                cmd.Parameters.AddWithValue("@Name", nameOrIP);
                var r = cmd.ExecuteScalar();
                if (r != null) return Convert.ToInt32(r);
            }

            // 2. Hostname ohne Domain-Suffix vergleichen (z.B. "PC-01" findet "PC-01.fritz.box")
            using (var cmd = new SQLiteCommand(
                "SELECT ID FROM Devices WHERE Hostname LIKE @Name || '.%' OR Hostname = @Name LIMIT 1", conn))
            {
                cmd.Parameters.AddWithValue("@Name", nameOrIP);
                var r = cmd.ExecuteScalar();
                if (r != null) return Convert.ToInt32(r);
            }

            // 3. Falls nameOrIP eine IP ist → Gerät anlegen (verhindert Datenverlust)
            if (System.Net.IPAddress.TryParse(nameOrIP, out _))
                return GetOrCreateDevice(conn, nameOrIP, null, null);

            return -1;
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
            // UPSERT: Falls (DeviceID, Name) existiert → Update, sonst Insert
            using (var cmd = new SQLiteCommand(@"
                INSERT INTO DeviceSoftware (DeviceID, PCName, Name, Version, Publisher, InstallLocation, InstallDate, Source)
                VALUES (@DeviceID, @PCName, @Name, @Version, @Publisher, @InstallLocation, @InstallDate, @Source)
                ON CONFLICT(DeviceID, Name) DO UPDATE SET
                    Version   = excluded.Version,
                    Publisher = excluded.Publisher,
                    InstallDate = excluded.InstallDate,
                    Source    = excluded.Source,
                    QueryTime = CURRENT_TIMESTAMP", conn))
            {
                cmd.Parameters.AddWithValue("@DeviceID", deviceID);
                cmd.Parameters.AddWithValue("@PCName", sw.PCName ?? "");
                cmd.Parameters.AddWithValue("@Name", sw.Name);
                cmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                cmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                cmd.Parameters.AddWithValue("@InstallLocation", sw.InstallLocation ?? "");
                cmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                cmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                cmd.ExecuteNonQuery();
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
                        ds.ID, ds.DeviceID, ds.QueryTime as Zeitstempel,
                        COALESCE(d.Hostname, ds.PCName) as PCName,
                        ds.Name, ds.Version, ds.Publisher, ds.InstallDate,
                        ds.QueryTime as LastUpdate
                    FROM DeviceSoftware ds
                    LEFT JOIN Devices d ON d.ID = ds.DeviceID
                                       OR (ds.DeviceID IS NULL AND (d.IP = ds.PCName OR d.Hostname = ds.PCName))
                    {whereClause}
                    ORDER BY ds.QueryTime DESC";

                using (var cmd = new SQLiteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        software.Add(new DatabaseSoftware
                        {
                            ID         = Convert.ToInt32(reader["ID"]),
                            DeviceID   = reader["DeviceID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DeviceID"]),
                            Zeitstempel = reader["Zeitstempel"].ToString(),
                            Name       = reader["Name"].ToString(),
                            Version    = reader["Version"]?.ToString(),
                            Publisher  = reader["Publisher"]?.ToString(),
                            InstallDate = reader["InstallDate"]?.ToString(),
                            PCName     = reader["PCName"]?.ToString(),
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
        {
            ExecuteNonQuery(
                "INSERT OR IGNORE INTO LocationDevices (LocationID, DeviceID) VALUES (@LocationID, @DeviceID)",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });
            ExecuteNonQuery(
                "UPDATE Devices SET StandortID = @LocationID WHERE ID = @DeviceID",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });
        }

        /// <summary>
        /// Entfernt die Zuweisung eines Devices von einer Location.
        /// </summary>
        public void UnassignDeviceFromLocation(int locationId, int deviceId)
        {
            ExecuteNonQuery(
                "DELETE FROM LocationDevices WHERE LocationID=@LocationID AND DeviceID=@DeviceID",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });
            ExecuteNonQuery(
                "UPDATE Devices SET StandortID = NULL WHERE ID = @DeviceID AND StandortID = @LocationID",
                new[] { ("@LocationID", locationId.ToString()), ("@DeviceID", deviceId.ToString()) });
        }

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
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();

                // DeviceID anhand der IP-Adresse ermitteln (oder Gerät neu anlegen)
                int? deviceId = null;
                using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP=@IP LIMIT 1", conn))
                {
                    cmd.Parameters.AddWithValue("@IP", ip);
                    var r = cmd.ExecuteScalar();
                    if (r != null) deviceId = Convert.ToInt32(r);
                }

                // LocationIPs befüllen (IP-basiert, für Rückwärtskompatibilität)
                using (var cmd = new SQLiteCommand(
                    "INSERT OR IGNORE INTO LocationIPs (LocationID, IPAddress, WorkstationName, DeviceID) VALUES (@LID, @IP, @WS, @DID)", conn))
                {
                    cmd.Parameters.AddWithValue("@LID", locationId);
                    cmd.Parameters.AddWithValue("@IP", ip);
                    cmd.Parameters.AddWithValue("@WS", workstationName ?? "");
                    cmd.Parameters.AddWithValue("@DID", deviceId.HasValue ? (object)deviceId.Value : DBNull.Value);
                    cmd.ExecuteNonQuery();
                }

                // LocationDevices befüllen (ID-basiert, für saubere FK-Verknüpfung)
                if (deviceId.HasValue)
                {
                    using (var cmd = new SQLiteCommand(
                        "INSERT OR IGNORE INTO LocationDevices (LocationID, DeviceID) VALUES (@LID, @DID)", conn))
                    {
                        cmd.Parameters.AddWithValue("@LID", locationId);
                        cmd.Parameters.AddWithValue("@DID", deviceId.Value);
                        cmd.ExecuteNonQuery();
                    }
                }

                // Bestehende LocationIPs-Einträge ohne DeviceID nachträglich verknüpfen
                if (deviceId.HasValue)
                {
                    using (var cmd = new SQLiteCommand(
                        "UPDATE LocationIPs SET DeviceID=@DID WHERE IPAddress=@IP AND DeviceID IS NULL", conn))
                    {
                        cmd.Parameters.AddWithValue("@DID", deviceId.Value);
                        cmd.Parameters.AddWithValue("@IP", ip);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

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

        // RefreshInventarTables() entfällt — Inventar_* sind jetzt VIEWs (live, immer aktuell)
    }
}