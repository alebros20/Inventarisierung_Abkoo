using System;
using System.Collections.Generic;

namespace NmapInventory
{
    // === DEVICE DATA ===
    public class DeviceInfo 
    { 
        public string IP { get; set; } 
        public string Hostname { get; set; } 
        public string Status { get; set; } 
        public string Ports { get; set; } 
    }
    
    public class DatabaseDevice 
    { 
        public int ID { get; set; }
        public string Zeitstempel { get; set; } 
        public string IP { get; set; } 
        public string Hostname { get; set; } 
        public string Status { get; set; } 
        public string Ports { get; set; } 
    }
    
    public class NmapScanResult 
    { 
        public string RawOutput { get; set; } 
        public List<DeviceInfo> Devices { get; set; } 
    }

    // === SOFTWARE DATA ===
    public class SoftwareInfo 
    { 
        public string Name { get; set; } 
        public string Version { get; set; } 
        public string Publisher { get; set; } 
        public string InstallLocation { get; set; } 
        public string Source { get; set; } 
        public string InstallDate { get; set; }
        public string PCName { get; set; }
        public DateTime Timestamp { get; set; }
        public string LastUpdate { get; set; }
    }
    
    public class DatabaseSoftware 
    { 
        public int ID { get; set; }
        public string Zeitstempel { get; set; } 
        public string Name { get; set; } 
        public string Version { get; set; } 
        public string Publisher { get; set; } 
        public string InstallDate { get; set; }
        public string PCName { get; set; }
        public string LastUpdate { get; set; }
    }

    // === CUSTOMER/LOCATION DATA ===
    public class Customer
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }
    
    public class Location
    {
        public int ID { get; set; }
        public int CustomerID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }
    
    public class LocationIP
    {
        public int ID { get; set; }
        public int LocationID { get; set; }
        public string IPAddress { get; set; }
        public string WorkstationName { get; set; }
    }
    
    public class NodeData
    {
        public string Type { get; set; } // "Customer", "Location", "IP"
        public int ID { get; set; }
        public object Data { get; set; }
    }
}
