using System;
using System.Data.SQLite;

namespace NmapInventory
{
    /// <summary>
    /// Enthält die gesamte Datenbank-Initialisierung und Migration.
    /// Wird von DatabaseManager aufgerufen.
    /// </summary>
    public class DatabaseInitializer
    {
        private readonly string _appDataDir;
        private readonly string _dbPath;

        public DatabaseInitializer(string appDataDir, string dbPath)
        {
            _appDataDir = appDataDir;
            _dbPath = dbPath;
        }

        // =========================================================
        // Haupt-Datenbank initialisieren
        // =========================================================
        public void InitializeDatabase()
        {
            System.IO.Directory.CreateDirectory(_appDataDir);
            using (var conn = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
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

                    CREATE TABLE IF NOT EXISTS LocationDevices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        LocationID INTEGER NOT NULL,
                        DeviceID INTEGER NOT NULL,
                        AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(LocationID) REFERENCES Locations(ID) ON DELETE CASCADE,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE,
                        UNIQUE(LocationID, DeviceID)
                    );

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

                    CREATE TABLE IF NOT EXISTS DeviceHardwareInfo (
                        ID          INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID    INTEGER NOT NULL,
                        HardwareText TEXT,
                        QueryTime   DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

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

                    CREATE INDEX IF NOT EXISTS idx_device_ports ON DevicePorts(DeviceID);

                    CREATE TABLE IF NOT EXISTS LocationIPs (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        LocationID INTEGER NOT NULL,
                        IPAddress TEXT NOT NULL,
                        WorkstationName TEXT,
                        AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(LocationID) REFERENCES Locations(ID) ON DELETE CASCADE
                    );
                ";

                using (var cmd = new SQLiteCommand(createTableQuery, conn))
                    cmd.ExecuteNonQuery();

                // Migrationen: fehlende Spalten nachrüsten
                TryAlter(conn, "ALTER TABLE DeviceSoftware ADD COLUMN PCName TEXT");
                TryAlter(conn, "ALTER TABLE DeviceHardwareInfo ADD COLUMN IPAddress TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN MacAddress TEXT UNIQUE");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CustomHostname INTEGER DEFAULT 0");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN DeviceType INTEGER DEFAULT 0");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN Vendor TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN OS TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN Comment TEXT DEFAULT ''");
                TryAlter(conn, "ALTER TABLE Customers ADD COLUMN CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlter(conn, "ALTER TABLE Locations ADD COLUMN Level INTEGER DEFAULT 0");
                TryAlter(conn, "ALTER TABLE Locations ADD COLUMN ParentID INTEGER");
                TryAlter(conn, "ALTER TABLE LocationIPs ADD COLUMN AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlter(conn, "ALTER TABLE LocationIPs ADD COLUMN DeviceID INTEGER REFERENCES Devices(ID)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_mac_address ON Devices(MacAddress)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_device_mac_history ON DeviceMacHistory(MacAddress)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_location_parent ON Locations(ParentID)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_location_customer ON Locations(CustomerID)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_location_devices ON LocationDevices(LocationID)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_location_ips_device ON LocationIPs(DeviceID)");
                TryAlter(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_location_ips_unique ON LocationIPs(LocationID, IPAddress)");

                // Migration: StandortID auf Devices setzen (aus LocationDevices ableiten)
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN StandortID INTEGER");
                TryAlter(conn, @"
                    UPDATE Devices SET StandortID = (
                        SELECT ld.LocationID FROM LocationDevices ld WHERE ld.DeviceID = Devices.ID LIMIT 1
                    ) WHERE StandortID IS NULL AND EXISTS (SELECT 1 FROM LocationDevices ld WHERE ld.DeviceID = Devices.ID)");

                // UNIQUE-Index auf DeviceSoftware
                TryAlter(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_device_software_unique ON DeviceSoftware(DeviceID, Name)");

                using (var ctCmd = new SQLiteCommand(@"
                    CREATE TABLE IF NOT EXISTS CredentialTemplates (
                        ID              INTEGER PRIMARY KEY AUTOINCREMENT,
                        Name            TEXT NOT NULL,
                        Protocol        TEXT NOT NULL,
                        Username        TEXT,
                        EncryptedPass   TEXT NOT NULL,
                        Port            INTEGER,
                        DeviceTypes     TEXT,
                        Priority        INTEGER DEFAULT 0,
                        Created         DATETIME DEFAULT CURRENT_TIMESTAMP
                    )", conn))
                    ctCmd.ExecuteNonQuery();

                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CredentialTemplateID INTEGER");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CredentialVerified DATETIME");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN UniqueID TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN UniqueIDSource TEXT");

                using (var diCmd = new SQLiteCommand(@"
                    CREATE TABLE IF NOT EXISTS DeviceInterfaces (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        MacAddress TEXT NOT NULL,
                        IPAddress TEXT,
                        InterfaceType TEXT,
                        IsPrimary INTEGER DEFAULT 0,
                        FirstSeen DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    )", conn))
                    diCmd.ExecuteNonQuery();
                TryAlter(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_device_interfaces_mac ON DeviceInterfaces(DeviceID, MacAddress)");
                TryAlter(conn, "CREATE INDEX IF NOT EXISTS idx_unique_id ON Devices(UniqueID)");

                // ── Inventar_* als VIEWs ──
                TryDropView(conn, "Inventar_Geraete");
                TryDropView(conn, "Inventar_Software");
                TryDropView(conn, "Inventar_Ports");
                TryDropView(conn, "Inventar_Hardware");

                using (var viewCmd = new SQLiteCommand(conn))
                {
                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Geraete AS
                        SELECT
                            c.ID  AS KundeID,     d.ID  AS GeraetID,
                            c.Name AS KundeName,
                            l.ID  AS StandortID,  l.Name AS StandortName,
                            d.IP,  d.Hostname,  d.MacAddress,
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
                            d.UniqueID, d.UniqueIDSource,
                            d.FirstSeen AS ErsterScan, d.LastSeen AS LetzterScan,
                            (SELECT COUNT(*) FROM DevicePorts    WHERE DeviceID = d.ID) AS AnzahlPorts,
                            (SELECT COUNT(*) FROM DeviceSoftware WHERE DeviceID = d.ID) AS AnzahlSoftware
                        FROM Devices d
                        LEFT JOIN Locations l  ON l.ID = d.StandortID
                        LEFT JOIN Customers c  ON c.ID = l.CustomerID;";
                    viewCmd.ExecuteNonQuery();

                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Software AS
                        SELECT
                            sw.ID AS SoftwareID,  d.ID  AS GeraetID,
                            d.IP,  d.Hostname,  d.MacAddress,
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
                            sw.Name AS SoftwareName,
                            sw.Version AS SoftwareVersion,
                            sw.Publisher AS Hersteller, sw.InstallDate AS InstallDatum,
                            sw.Source AS Quelle
                        FROM DeviceSoftware sw
                        JOIN Devices d         ON d.ID = sw.DeviceID;";
                    viewCmd.ExecuteNonQuery();

                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Ports AS
                        SELECT
                            p.ID AS PortID,       d.ID  AS GeraetID,
                            d.IP,  d.Hostname,
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
                            p.Port, p.Protocol AS Protokoll,
                            p.State AS Status, p.Service AS Dienst, p.Version
                        FROM DevicePorts p
                        JOIN Devices d         ON d.ID = p.DeviceID;";
                    viewCmd.ExecuteNonQuery();

                    viewCmd.CommandText = @"
                        CREATE VIEW Inventar_Hardware AS
                        SELECT
                            h.ID AS HardwareID,   d.ID  AS GeraetID,
                            d.IP,  d.Hostname,  d.MacAddress,
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
                            h.HardwareText,
                            h.QueryTime AS AbfrageZeit
                        FROM DeviceHardwareInfo h
                        JOIN Devices d         ON d.ID = h.DeviceID;";
                    viewCmd.ExecuteNonQuery();
                }

                // Daten-Migrationen
                TryAlter(conn, @"
                    UPDATE LocationIPs SET DeviceID = (
                        SELECT d.ID FROM Devices d WHERE d.IP = LocationIPs.IPAddress LIMIT 1
                    ) WHERE DeviceID IS NULL");

                TryAlter(conn, @"
                    INSERT OR IGNORE INTO LocationDevices (LocationID, DeviceID)
                    SELECT LocationID, DeviceID FROM LocationIPs
                    WHERE DeviceID IS NOT NULL");

                TryAlter(conn, @"
                    UPDATE DeviceSoftware SET DeviceID = (
                        SELECT d.ID FROM Devices d WHERE d.IP = DeviceSoftware.PCName LIMIT 1
                    ) WHERE PCName IS NOT NULL AND PCName != ''
                      AND EXISTS (SELECT 1 FROM Devices WHERE IP = DeviceSoftware.PCName)");

                TryAlter(conn, @"
                    UPDATE Devices SET Hostname = '', CustomHostname = 0
                    WHERE CustomHostname = 0
                      AND (
                        Hostname GLOB '*(*:*:*:*:*:*)*'
                        OR Hostname GLOB '[0-9A-Fa-f][0-9A-Fa-f]:[0-9A-Fa-f][0-9A-Fa-f]:*'
                        OR Hostname = IP
                      )");
                TryAlter(conn, @"
                    UPDATE LocationIPs SET WorkstationName = NULL
                    WHERE (
                        WorkstationName GLOB '*(*:*:*:*:*:*)*'
                        OR WorkstationName GLOB '[0-9A-Fa-f][0-9A-Fa-f]:[0-9A-Fa-f][0-9A-Fa-f]:*'
                        OR WorkstationName = IPAddress
                    )");
            }
        }

        // =========================================================
        // Kunden-Datenbank initialisieren / migrieren
        // =========================================================
        public void EnsureCustomerDatabaseExists(int customerId)
        {
            var path = System.IO.Path.Combine(_appDataDir, $"nmap_customer_{customerId}.db");
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

                // Migrationen
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CustomHostname INTEGER DEFAULT 0");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN Comment TEXT DEFAULT ''");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN Vendor TEXT DEFAULT ''");
                TryAlter(conn, "ALTER TABLE DeviceSoftware ADD COLUMN PCName TEXT");
                TryAlter(conn, "ALTER TABLE DeviceSoftware ADD COLUMN Timestamp DATETIME");
                TryAlter(conn, "ALTER TABLE DeviceSoftware ADD COLUMN LastUpdate TEXT");
                TryAlter(conn, "ALTER TABLE DeviceSoftware ADD COLUMN DeviceID INTEGER");
                TryAlter(conn, "ALTER TABLE DeviceHardwareInfo ADD COLUMN IPAddress TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CredentialTemplateID INTEGER");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN CredentialVerified DATETIME");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN UniqueID TEXT");
                TryAlter(conn, "ALTER TABLE Devices ADD COLUMN UniqueIDSource TEXT");

                using (var diCmd = new SQLiteCommand(@"
                    CREATE TABLE IF NOT EXISTS DeviceInterfaces (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        MacAddress TEXT NOT NULL,
                        IPAddress TEXT,
                        InterfaceType TEXT,
                        IsPrimary INTEGER DEFAULT 0,
                        FirstSeen DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    )", conn))
                    diCmd.ExecuteNonQuery();
                TryAlter(conn, "CREATE UNIQUE INDEX IF NOT EXISTS idx_device_interfaces_mac ON DeviceInterfaces(DeviceID, MacAddress)");
            }
        }

        // ── Hilfsmethoden ──────────────────────────────────────
        private static void TryAlter(SQLiteConnection conn, string sql)
        {
            try { using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery(); }
            catch { }
        }

        private static void TryDropView(SQLiteConnection conn, string viewName)
        {
            try
            {
                using (var cmd = new SQLiteCommand($"DROP VIEW IF EXISTS {viewName};", conn))
                    cmd.ExecuteNonQuery();
            }
            catch
            {
                try
                {
                    using (var cmd = new SQLiteCommand($"DROP TABLE IF EXISTS {viewName};", conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }
            }
        }
    }
}
