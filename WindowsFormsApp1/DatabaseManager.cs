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
                    -- Geräte Tabelle (zentral)
                    CREATE TABLE IF NOT EXISTS Devices (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        IP TEXT UNIQUE NOT NULL,
                        Hostname TEXT,
                        FirstSeen DATETIME DEFAULT CURRENT_TIMESTAMP,
                        LastSeen DATETIME DEFAULT CURRENT_TIMESTAMP
                    );

                    -- Scan-Verlauf für jedes Gerät
                    CREATE TABLE IF NOT EXISTS DeviceScanHistory (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Status TEXT,
                        Ports TEXT,
                        ScanTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    -- Software für jedes Gerät mit Zeitstempel
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

                    -- Software-Verlauf (für Update-Tracking)
                    CREATE TABLE IF NOT EXISTS SoftwareHistory (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        DeviceID INTEGER NOT NULL,
                        Name TEXT NOT NULL,
                        OldVersion TEXT,
                        NewVersion TEXT,
                        ChangeTime DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(DeviceID) REFERENCES Devices(ID) ON DELETE CASCADE
                    );

                    -- Kunden (Top-Level)
                    CREATE TABLE IF NOT EXISTS Customers (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        Name TEXT NOT NULL,
                        Address TEXT,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP
                    );

                    -- Hierarchische Struktur: Kunde > Standort > Abteilung
                    CREATE TABLE IF NOT EXISTS Locations (
                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                        CustomerID INTEGER NOT NULL,
                        ParentID INTEGER,
                        Name TEXT NOT NULL,
                        Address TEXT,
                        Level INTEGER DEFAULT 0,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(CustomerID) REFERENCES Customers(ID) ON DELETE CASCADE,
                        FOREIGN KEY(ParentID) REFERENCES Locations(ID) ON DELETE CASCADE
                    );

                    -- IPs zugeordnet zu Abteilungen/Standorten
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

                TryAlterTable(conn, "ALTER TABLE Customers ADD COLUMN CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN Level INTEGER DEFAULT 0");
                TryAlterTable(conn, "ALTER TABLE Locations ADD COLUMN ParentID INTEGER");
                TryAlterTable(conn, "ALTER TABLE LocationIPs ADD COLUMN AssignedDate DATETIME DEFAULT CURRENT_TIMESTAMP");
            }
        }

        private void TryAlterTable(SQLiteConnection conn, string sql)
        {
            try { using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery(); }
            catch { }
        }

        // === DEVICES ===
        public void SaveDevices(List<DeviceInfo> devices)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var dev in devices)
                {
                    int deviceID = GetOrCreateDevice(conn, dev.IP, dev.Hostname);
                    SaveDeviceScanHistory(conn, deviceID, dev.Status, dev.Ports);
                    UpdateDeviceLastSeen(conn, deviceID);
                }
            }
        }

        private int GetOrCreateDevice(SQLiteConnection conn, string ip, string hostname)
        {
            using (var cmd = new SQLiteCommand("SELECT ID FROM Devices WHERE IP = @IP", conn))
            {
                cmd.Parameters.AddWithValue("@IP", ip);
                var result = cmd.ExecuteScalar();
                if (result != null)
                {
                    return Convert.ToInt32(result);
                }
            }

            using (var cmd = new SQLiteCommand("INSERT INTO Devices (IP, Hostname) VALUES (@IP, @Hostname)", conn))
            {
                cmd.Parameters.AddWithValue("@IP", ip);
                cmd.Parameters.AddWithValue("@Hostname", hostname ?? "");
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

        public List<DatabaseDevice> LoadDevices(string filter = "Alle")
        {
            var devices = new List<DatabaseDevice>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                string whereClause = GetDateFilter(filter, "d.LastSeen");
                string query = $@"
                    SELECT DISTINCT 
                        d.ID,
                        d.IP,
                        d.Hostname,
                        d.LastSeen as Zeitstempel,
                        sh.Status,
                        sh.Ports
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
                {
                    using (var cmd = new SQLiteCommand("UPDATE Devices SET Hostname=@Hostname WHERE ID=@ID", conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", dev.ID);
                        cmd.Parameters.AddWithValue("@Hostname", dev.Hostname ?? "");
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        public void DeleteDevice(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM Devices WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // === SOFTWARE ===
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
                    {
                        SaveSoftwareHistory(conn, deviceID, sw.Name, oldVersion, sw.Version);
                    }

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
                var result = cmd.ExecuteScalar();
                return result?.ToString();
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
                {
                    using (var updateCmd = new SQLiteCommand("UPDATE DeviceSoftware SET Version=@Version, Publisher=@Publisher, InstallDate=@InstallDate, Source=@Source, QueryTime=CURRENT_TIMESTAMP WHERE ID=@ID", conn))
                    {
                        updateCmd.Parameters.AddWithValue("@ID", Convert.ToInt32(result));
                        updateCmd.Parameters.AddWithValue("@Version", sw.Version ?? "");
                        updateCmd.Parameters.AddWithValue("@Publisher", sw.Publisher ?? "");
                        updateCmd.Parameters.AddWithValue("@InstallDate", sw.InstallDate ?? "");
                        updateCmd.Parameters.AddWithValue("@Source", sw.Source ?? "");
                        updateCmd.ExecuteNonQuery();
                    }
                }
                else
                {
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
                        ds.ID,
                        ds.QueryTime as Zeitstempel,
                        d.Hostname as PCName,
                        ds.Name,
                        ds.Version,
                        ds.Publisher,
                        ds.InstallDate,
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

                        if (result != null)
                        {
                            sw.LastUpdate = Convert.ToDateTime(result).ToString("dd.MM.yyyy HH:mm");
                        }
                        else
                        {
                            sw.LastUpdate = "-";
                        }
                    }
                }
            }
        }

        public void UpdateSoftwareEntries(List<DatabaseSoftware> software)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                foreach (var sw in software)
                {
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
        }

        public void DeleteSoftwareEntry(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("DELETE FROM DeviceSoftware WHERE ID=@ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // === CUSTOMERS ===
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
            => ExecuteNonQuery($"INSERT INTO Customers (Name, Address) VALUES (@Name, @Address)", new[] { ("@Name", name), ("@Address", address ?? "") });

        public void UpdateCustomer(int id, string name, string address)
            => ExecuteNonQuery($"UPDATE Customers SET Name=@Name, Address=@Address WHERE ID=@ID", new[] { ("@ID", id.ToString()), ("@Name", name), ("@Address", address ?? "") });

        public void DeleteCustomer(int id)
            => ExecuteNonQuery($"DELETE FROM Customers WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        // === LOCATIONS (HIERARCHISCH) ===
        /// <summary>
        /// Holt alle Top-Level Standorte für einen Kunden (ParentID = NULL)
        /// </summary>
        public List<Location> GetLocationsByCustomer(int customerId)
        {
            return GetLocationsByCustomerAndParent(customerId, null);
        }

        /// <summary>
        /// Holt alle Kinder-Standorte/Abteilungen für einen Parent-Standort
        /// </summary>
        public List<Location> GetChildLocations(int parentLocationId)
        {
            var locations = new List<Location>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE ParentID = @ParentID ORDER BY Name", conn))
                {
                    cmd.Parameters.AddWithValue("@ParentID", parentLocationId);
                    using (var reader = cmd.ExecuteReader())
                        while (reader.Read())
                            locations.Add(new Location
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                CustomerID = Convert.ToInt32(reader["CustomerID"]),
                                ParentID = Convert.ToInt32(reader["ParentID"]),
                                Name = reader["Name"].ToString(),
                                Address = reader["Address"]?.ToString() ?? "",
                                Level = Convert.ToInt32(reader["Level"])
                            });
                }
            }
            return locations;
        }

        /// <summary>
        /// Holt Standorte/Abteilungen mit optionalem Parent-Filter
        /// </summary>
        private List<Location> GetLocationsByCustomerAndParent(int customerId, int? parentId)
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
                            locations.Add(new Location
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                CustomerID = Convert.ToInt32(reader["CustomerID"]),
                                ParentID = reader["ParentID"] == DBNull.Value ? -1 : Convert.ToInt32(reader["ParentID"]),
                                Name = reader["Name"].ToString(),
                                Address = reader["Address"]?.ToString() ?? "",
                                Level = Convert.ToInt32(reader["Level"])
                            });
                }
            }
            return locations;
        }

        /// <summary>
        /// Erstellt einen neuen Top-Level Standort
        /// </summary>
        public void AddLocation(int customerId, string name, string address)
            => ExecuteNonQuery(
                "INSERT INTO Locations (CustomerID, ParentID, Name, Address, Level) VALUES (@CustomerID, NULL, @Name, @Address, 0)",
                new[] { ("@CustomerID", customerId.ToString()), ("@Name", name), ("@Address", address ?? "") });

        /// <summary>
        /// Erstellt eine Abteilung unter einem Standort/Parent
        /// </summary>
        public void AddChildLocation(int parentLocationId, string name, string address)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();

                // Hole Parent-Infos um Level zu erhöhen
                int parentLevel = 0;
                int customerID = 0;
                using (var cmd = new SQLiteCommand("SELECT Level, CustomerID FROM Locations WHERE ID = @ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", parentLocationId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            parentLevel = Convert.ToInt32(reader["Level"]);
                            customerID = Convert.ToInt32(reader["CustomerID"]);
                        }
                    }
                }

                // Erstelle Kind mit erhöhtem Level
                using (var cmd = new SQLiteCommand("INSERT INTO Locations (CustomerID, ParentID, Name, Address, Level) VALUES (@CustomerID, @ParentID, @Name, @Address, @Level)", conn))
                {
                    cmd.Parameters.AddWithValue("@CustomerID", customerID);
                    cmd.Parameters.AddWithValue("@ParentID", parentLocationId);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Address", address ?? "");
                    cmd.Parameters.AddWithValue("@Level", parentLevel + 1);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void UpdateLocation(int id, string name, string address)
            => ExecuteNonQuery($"UPDATE Locations SET Name=@Name, Address=@Address WHERE ID=@ID", new[] { ("@ID", id.ToString()), ("@Name", name), ("@Address", address ?? "") });

        public void DeleteLocation(int id)
            => ExecuteNonQuery($"DELETE FROM Locations WHERE ID=@ID", new[] { ("@ID", id.ToString()) });

        public Location GetLocationByID(int id)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, CustomerID, ParentID, Name, Address, Level FROM Locations WHERE ID = @ID", conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new Location
                            {
                                ID = Convert.ToInt32(reader["ID"]),
                                CustomerID = Convert.ToInt32(reader["CustomerID"]),
                                ParentID = reader["ParentID"] == DBNull.Value ? -1 : Convert.ToInt32(reader["ParentID"]),
                                Name = reader["Name"].ToString(),
                                Address = reader["Address"]?.ToString() ?? "",
                                Level = Convert.ToInt32(reader["Level"])
                            };
                        }
                    }
                }
            }
            return null;
        }

        // === LOCATION IPs ===
        public List<LocationIP> GetIPsWithWorkstationByLocation(int locationId)
        {
            var ips = new List<LocationIP>();
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT ID, LocationID, IPAddress, WorkstationName FROM LocationIPs WHERE LocationID=@LocationID ORDER BY WorkstationName, IPAddress", conn))
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

        public void AddIPToLocation(int locationId, string ip, string workstationName = "")
            => ExecuteNonQuery($"INSERT INTO LocationIPs (LocationID, IPAddress, WorkstationName) VALUES (@LocationID, @IPAddress, @WorkstationName)", new[] { ("@LocationID", locationId.ToString()), ("@IPAddress", ip), ("@WorkstationName", workstationName ?? "") });

        public void RemoveIPFromLocation(int locationId, string ip)
            => ExecuteNonQuery($"DELETE FROM LocationIPs WHERE LocationID=@LocationID AND IPAddress=@IPAddress", new[] { ("@LocationID", locationId.ToString()), ("@IPAddress", ip) });

        /// <summary>
        /// Holt alle IPs einer Location PLUS aller Child-Locations rekursiv
        /// </summary>
        public List<LocationIP> GetAllIPsRecursive(int locationId)
        {
            var allIPs = new List<LocationIP>();

            // IPs der aktuellen Location
            allIPs.AddRange(GetIPsWithWorkstationByLocation(locationId));

            // Rekursiv alle Child-Locations durchsuchen
            var children = GetChildLocations(locationId);
            foreach (var child in children)
            {
                allIPs.AddRange(GetAllIPsRecursive(child.ID));
            }

            return allIPs;
        }

        private void ExecuteNonQuery(string query, (string, string)[] parameters)
        {
            using (var conn = new SQLiteConnection($"Data Source={DB_PATH};Version=3;"))
            {
                conn.Open();
                ExecuteCommand(conn, query, parameters);
            }
        }

        private void ExecuteCommand(SQLiteConnection conn, string query, (string, string)[] parameters)
        {
            using (var cmd = new SQLiteCommand(query, conn))
            {
                foreach (var (key, value) in parameters)
                    cmd.Parameters.AddWithValue(key, value ?? "");
                cmd.ExecuteNonQuery();
            }
        }

        private string GetDateFilter(string filter, string dateColumn = "Zeitstempel")
        {
            switch (filter)
            {
                case "Heute":
                    return $"WHERE DATE({dateColumn}) = DATE('now')";
                case "Diese Woche":
                    return $"WHERE DATE({dateColumn}) >= DATE('now', '-7 days')";
                case "Dieser Monat":
                    return $"WHERE DATE({dateColumn}) >= DATE('now', 'start of month')";
                case "Dieses Jahr":
                    return $"WHERE DATE({dateColumn}) >= DATE('now', 'start of year')";
                default:
                    return "";
            }
        }
    }
}