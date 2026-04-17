using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using Lextm.SharpSnmpLib.Security;

namespace NmapInventory
{
    public class SnmpManager
    {
        private static readonly string OID_SysDescr    = "1.3.6.1.2.1.1.1.0";
        private static readonly string OID_SysObjectID = "1.3.6.1.2.1.1.2.0";
        private static readonly string OID_SysUpTime   = "1.3.6.1.2.1.1.3.0";
        private static readonly string OID_SysContact  = "1.3.6.1.2.1.1.4.0";
        private static readonly string OID_SysName     = "1.3.6.1.2.1.1.5.0";
        private static readonly string OID_SysLocation = "1.3.6.1.2.1.1.6.0";

        public SnmpResult Query(string ipAddress, SnmpSettings settings)
        {
            try
            {
                switch (settings.Version)
                {
                    case 1:  return QueryV1(ipAddress, settings);
                    case 2:  return QueryV2c(ipAddress, settings);
                    case 3:  return QueryV3(ipAddress, settings);
                    default: return Error($"Ungültige SNMP-Version: {settings.Version}");
                }
            }
            catch (Exception ex)
            {
                return Error(ex.Message);
            }
        }

        private SnmpResult QueryV1(string ip, SnmpSettings s)
        {
            var community = new OctetString(s.Community);
            var endpoint  = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids      = BuildOidList();
            var result = Messenger.Get(VersionCode.V1, endpoint, community, oids, s.TimeoutMs);
            return ParseGetResult(result, "v1");
        }

        private SnmpResult QueryV2c(string ip, SnmpSettings s)
        {
            var community = new OctetString(s.Community);
            var endpoint  = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids      = BuildOidList();
            var result = Messenger.Get(VersionCode.V2, endpoint, community, oids, s.TimeoutMs);
            return ParseGetResult(result, "v2c");
        }

        private SnmpResult QueryV3(string ip, SnmpSettings s)
        {
            var endpoint = new IPEndPoint(IPAddress.Parse(ip), s.Port);
            var oids     = BuildOidList();

            IAuthenticationProvider authProvider = BuildAuthProvider(s);
            IPrivacyProvider privProvider = BuildPrivacyProvider(s, authProvider);

            var discovery = Messenger.GetNextDiscovery(SnmpType.GetRequestPdu);
            var report    = discovery.GetResponse(s.TimeoutMs, endpoint);

            var request = new GetRequestMessage(
                VersionCode.V3,
                Messenger.NextMessageId,
                Messenger.NextRequestId,
                new OctetString(s.Username),
                OctetString.Empty,
                oids,
                privProvider,
                Messenger.MaxMessageSize,
                report);

            var reply = request.GetResponse(s.TimeoutMs, endpoint);

            if (reply.Pdu().ErrorStatus.ToInt32() != 0)
                return Error($"SNMP v3 Fehler: ErrorStatus={reply.Pdu().ErrorStatus}");

            return ParseGetResult(reply.Pdu().Variables.ToList(), "v3");
        }

        private static List<Variable> BuildOidList() => new List<Variable>
        {
            new Variable(new ObjectIdentifier(OID_SysDescr)),
            new Variable(new ObjectIdentifier(OID_SysObjectID)),
            new Variable(new ObjectIdentifier(OID_SysUpTime)),
            new Variable(new ObjectIdentifier(OID_SysContact)),
            new Variable(new ObjectIdentifier(OID_SysName)),
            new Variable(new ObjectIdentifier(OID_SysLocation)),
        };

        private static SnmpResult ParseGetResult(IList<Variable> vars, string version)
        {
            var r = new SnmpResult { Success = true, SnmpVersion = version };
            foreach (var v in vars)
            {
                string oid = v.Id.ToString();
                string val = v.Data.ToString();

                if (oid == OID_SysDescr)    r.SysDescr    = val;
                else if (oid == OID_SysObjectID) r.SysObjectID = val;
                else if (oid == OID_SysUpTime)   r.SysUpTime   = FormatUptime(val);
                else if (oid == OID_SysContact)  r.SysContact  = val;
                else if (oid == OID_SysName)     r.SysName     = val;
                else if (oid == OID_SysLocation) r.SysLocation = val;
            }
            return r;
        }

        private static string FormatUptime(string raw)
        {
            if (uint.TryParse(raw, out uint ticks))
            {
                var ts = TimeSpan.FromSeconds(ticks / 100.0);
                return $"{(int)ts.TotalDays}d {ts.Hours:D2}h {ts.Minutes:D2}m";
            }
            return raw;
        }

        private static IAuthenticationProvider BuildAuthProvider(SnmpSettings s)
        {
            if (s.SecurityLevel == "NoAuth")
                return DefaultAuthenticationProvider.Instance;

            switch (s.AuthProtocol?.ToUpper())
            {
                case "SHA384": return new SHA384AuthenticationProvider(new OctetString(s.AuthPassword));
                case "SHA512": return new SHA512AuthenticationProvider(new OctetString(s.AuthPassword));
                default:       return new SHA256AuthenticationProvider(new OctetString(s.AuthPassword));
            }
        }

        private static IPrivacyProvider BuildPrivacyProvider(SnmpSettings s, IAuthenticationProvider auth)
        {
            if (s.SecurityLevel == "NoAuth" || s.SecurityLevel == "AuthNoPriv")
                return new DefaultPrivacyProvider(auth);

            switch (s.PrivProtocol?.ToUpper())
            {
                case "AES256": return new AES256PrivacyProvider(new OctetString(s.PrivPassword), auth);
                case "AES192": return new AES192PrivacyProvider(new OctetString(s.PrivPassword), auth);
                default:       return new AESPrivacyProvider(new OctetString(s.PrivPassword), auth);
            }
        }

        private static SnmpResult Error(string msg)
            => new SnmpResult { Success = false, ErrorMessage = msg };

        public void QueryDevices(IEnumerable<DeviceInfo> devices, SnmpSettings settings,
            Action<string> progress = null)
        {
            foreach (var device in devices)
            {
                progress?.Invoke($"SNMP → {device.IP}");
                try
                {
                    device.SnmpData = Query(device.IP, settings);
                    if (device.SnmpData.Success &&
                        string.IsNullOrEmpty(device.Hostname) &&
                        !string.IsNullOrEmpty(device.SnmpData.SysName))
                        device.Hostname = device.SnmpData.SysName;
                }
                catch
                {
                    device.SnmpData = new SnmpResult { Success = false, ErrorMessage = "Timeout / nicht erreichbar" };
                }
            }
        }
    }
}
