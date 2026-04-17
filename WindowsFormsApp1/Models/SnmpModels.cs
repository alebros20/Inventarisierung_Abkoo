using System;
using System.Collections.Generic;

namespace NmapInventory
{
    public class ScanResult
    {
        public List<DeviceInfo> Devices { get; set; }
        public string RawOutput { get; set; }
    }

    public class SnmpSettings
    {
        public int Version { get; set; } = 2;
        public string Community { get; set; } = "public";
        public int Port { get; set; } = 161;
        public int TimeoutMs { get; set; } = 2000;

        public string Username { get; set; } = "";
        public string AuthPassword { get; set; } = "";
        public string PrivPassword { get; set; } = "";
        public string AuthProtocol { get; set; } = "SHA256";
        public string PrivProtocol { get; set; } = "AES128";
        public string SecurityLevel { get; set; } = "AuthPriv";
    }

    public class SnmpResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string SysDescr { get; set; }
        public string SysObjectID { get; set; }
        public string SysUpTime { get; set; }
        public string SysContact { get; set; }
        public string SysName { get; set; }
        public string SysLocation { get; set; }
        public string SnmpVersion { get; set; }
        public DateTime QueryTime { get; set; } = DateTime.Now;
    }
}
