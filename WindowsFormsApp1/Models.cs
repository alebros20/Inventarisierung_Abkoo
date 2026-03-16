using System;
using System.Collections.Generic;

namespace NmapInventory
{
    public class DeviceInfo
    {
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
        public string MacAddress { get; set; }

        // Erweiterte Nmap-Felder
        public string OS { get; set; }   // OS-Erkennung
        public string OSDetails { get; set; }   // Detail-String z.B. "Windows 10 / Server 2016"
        public string Vendor { get; set; }   // MAC-Hersteller
        public List<NmapPort> OpenPorts { get; set; } = new List<NmapPort>();
    }

    // Einzelner Port mit Service-Info
    public class NmapPort
    {
        public int Port { get; set; }
        public string Protocol { get; set; }   // tcp / udp
        public string State { get; set; }   // open / closed / filtered
        public string Service { get; set; }   // http, ssh, smb ...
        public string Version { get; set; }   // Service-Version
        public string Banner { get; set; }   // Banner-Text
    }

    // Vollständige Nmap-Scan-Details pro Gerät — wird in DB gespeichert
    public class NmapScanDetail
    {
        public int ID { get; set; }
        public int DeviceID { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Vendor { get; set; }
        public string OS { get; set; }
        public string OSDetails { get; set; }
        public string Ports { get; set; }   // Kompakter String: "80/tcp open http"
        public string PortsJson { get; set; }   // JSON für detaillierte Ansicht
        public string RawOutput { get; set; }   // Roher Nmap-Block für dieses Gerät
        public DateTime ScanTime { get; set; }
    }

    public class DatabaseDevice
    {
        public int ID { get; set; }
        public string Zeitstempel { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
        public string OS { get; set; }
    }

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
        public string Zeitstempel { get; set; }
        public string PCName { get; set; }
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }
        public string InstallDate { get; set; }
        public string LastUpdate { get; set; }
    }

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
        public int ParentID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public int Level { get; set; }
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
        public string Type { get; set; }
        public int ID { get; set; }
        public object Data { get; set; }
    }

    public class ScanResult
    {
        public List<DeviceInfo> Devices { get; set; }
        public string RawOutput { get; set; }
    }
}