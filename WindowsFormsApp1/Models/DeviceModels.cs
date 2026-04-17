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

        public string OS { get; set; }
        public string OSDetails { get; set; }
        public string Vendor { get; set; }
        public List<NmapPort> OpenPorts { get; set; } = new List<NmapPort>();

        public DeviceType DeviceType { get; set; } = DeviceType.Unbekannt;
        public string DeviceTypeIcon => DeviceTypeHelper.GetIcon(DeviceType);

        public SnmpResult SnmpData { get; set; }
    }

    public class NmapPort
    {
        public int Port { get; set; }
        public string Protocol { get; set; }
        public string State { get; set; }
        public string Service { get; set; }
        public string Version { get; set; }
        public string Banner { get; set; }
    }

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
        public string Ports { get; set; }
        public string PortsJson { get; set; }
        public string RawOutput { get; set; }
        public DateTime ScanTime { get; set; }
    }

    public class DatabaseDevice
    {
        public int ID { get; set; }
        public string Zeitstempel { get; set; }
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string MacAddress { get; set; }
        public string Vendor { get; set; }
        public string Status { get; set; }
        public string Ports { get; set; }
        public string OS { get; set; }
        public string Comment { get; set; }
        public DeviceType DeviceType { get; set; } = DeviceType.Unbekannt;
        public string DeviceTypeIcon => DeviceTypeHelper.GetIcon(DeviceType);
        public int? CredentialTemplateID { get; set; }
        public string UniqueID { get; set; }
        public string UniqueIDSource { get; set; }
    }

    public class DeviceInterface
    {
        public int ID { get; set; }
        public int DeviceID { get; set; }
        public string MacAddress { get; set; }
        public string IPAddress { get; set; }
        public string InterfaceType { get; set; }
        public bool IsPrimary { get; set; }
    }
}
