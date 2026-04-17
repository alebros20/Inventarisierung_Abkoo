using System;

namespace NmapInventory
{
    public class SoftwareInfo
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string InstallLocation { get; set; }
        public string Source { get; set; }
        public string PCName { get; set; }
        public DateTime Timestamp { get; set; }
        public string LastUpdate { get; set; }
    }

    public class DatabaseSoftware
    {
        public int ID { get; set; }
        public int DeviceID { get; set; }
        public string Zeitstempel { get; set; }
        public string PCName { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string LastUpdate { get; set; }
    }
}
