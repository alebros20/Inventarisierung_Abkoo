using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;

namespace NmapInventory
{
    public class CredentialScanner
    {
        public class ScanResult
        {
            public DatabaseDevice Device { get; set; }
            public CredentialTemplate MatchedTemplate { get; set; }
            public bool Success { get; set; }
            public string Message { get; set; }
            public DeviceIdentifier.IdentifyResult Identity { get; set; }
        }

        public event Action<ScanResult> DeviceScanned;
        public event Action<int, int> ProgressChanged;

        private readonly int _maxParallel;
        private CancellationTokenSource _cts;

        public CredentialScanner(int maxParallel = 5)
        {
            _maxParallel = maxParallel;
        }

        public void Cancel() => _cts?.Cancel();

        public async Task<List<ScanResult>> ScanDevicesAsync(
            List<DatabaseDevice> devices,
            List<CredentialTemplate> templates,
            bool retestExisting)
        {
            _cts = new CancellationTokenSource();
            var results = new List<ScanResult>();
            var semaphore = new SemaphoreSlim(_maxParallel);
            int progress = 0;
            int total = devices.Count;

            var tasks = devices.Select(async device =>
            {
                if (_cts.Token.IsCancellationRequested) return;

                await semaphore.WaitAsync(_cts.Token);
                try
                {
                    var result = await Task.Run(() => TestDevice(device, templates, retestExisting), _cts.Token);
                    lock (results) results.Add(result);
                    DeviceScanned?.Invoke(result);
                    var p = Interlocked.Increment(ref progress);
                    ProgressChanged?.Invoke(p, total);
                }
                finally
                {
                    semaphore.Release();
                }
            });

            try { await Task.WhenAll(tasks); }
            catch (OperationCanceledException) { }

            return results;
        }

        private static readonly (string Protocol, int Port)[] ALL_PROTOCOLS = {
            ("WMI",    135),
            ("SSH",     22),
            ("SNMP",   161),
            ("Telnet",  23),
            ("HTTP",    80),
            ("HTTPS",  443)
        };

        private ScanResult TestDevice(DatabaseDevice device, List<CredentialTemplate> templates, bool retestExisting)
        {
            if (!retestExisting && device.CredentialTemplateID.HasValue)
                return new ScanResult { Device = device, Success = false, Message = "Übersprungen (bereits getestet)" };

            foreach (var template in templates)
            {
                try
                {
                    string password = CredentialStore.Decrypt(template.EncryptedPass);

                    foreach (var (protocol, port) in ALL_PROTOCOLS)
                    {
                        try
                        {
                            bool success = TestProtocol(device.IP, protocol, template.Username, password, port);
                            if (success)
                            {
                                DeviceIdentifier.IdentifyResult identity = null;
                                try
                                {
                                    var identifier = new DeviceIdentifier();
                                    if (protocol == "WMI")
                                        identity = identifier.IdentifyWindows(device.IP, template.Username, password);
                                    else if (protocol == "SSH")
                                        identity = identifier.IdentifyViaSsh(device.IP, template.Username, password, port);
                                }
                                catch { }

                                string msg = $"{protocol} {template.Username}:*** → Erfolg";
                                if (identity != null && !string.IsNullOrEmpty(identity.UniqueID))
                                    msg += $" [ID: {identity.Source}]";

                                return new ScanResult
                                {
                                    Device = device,
                                    MatchedTemplate = template,
                                    Success = true,
                                    Message = msg,
                                    Identity = identity
                                };
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return new ScanResult { Device = device, Success = false, Message = "Keine Vorlage passte" };
        }

        private bool TestProtocol(string ip, string protocol, string username, string password, int port)
        {
            switch (protocol)
            {
                case "WMI":    return TestWmi(ip, username, password);
                case "SSH":    return TestSsh(ip, username, password, port);
                case "SNMP":   return TestSnmp(ip, password, port);
                case "Telnet": return TestTelnet(ip, username, password, port);
                case "HTTP":   return TestHttp(ip, username, password, port, false);
                case "HTTPS":  return TestHttp(ip, username, password, port, true);
                default:       return false;
            }
        }

        private bool TestWmi(string ip, string username, string password)
        {
            try
            {
                var options = HardwareManager.BuildConnectionOptions(username, password);
                var scope = new System.Management.ManagementScope($@"\\{ip}\root\cimv2", options)
                {
                    Options = { Timeout = TimeSpan.FromSeconds(10) }
                };
                scope.Connect();
                return scope.IsConnected;
            }
            catch { return false; }
        }

        private bool TestSsh(string ip, string username, string password, int port)
        {
            try
            {
                using (var client = new Renci.SshNet.SshClient(
                    new Renci.SshNet.ConnectionInfo(ip, port, username,
                        new Renci.SshNet.PasswordAuthenticationMethod(username, password))
                    { Timeout = TimeSpan.FromSeconds(10) }))
                {
                    client.Connect();
                    bool connected = client.IsConnected;
                    client.Disconnect();
                    return connected;
                }
            }
            catch { return false; }
        }

        private bool TestSnmp(string ip, string community, int port)
        {
            try
            {
                using (var udp = new UdpClient())
                {
                    udp.Client.ReceiveTimeout = 5000;
                    udp.Client.SendTimeout = 5000;

                    byte[] getRequest = BuildSnmpGetRequest(community, "1.3.6.1.2.1.1.1.0");
                    udp.Send(getRequest, getRequest.Length, new IPEndPoint(IPAddress.Parse(ip), port));

                    var remoteEP = new IPEndPoint(IPAddress.Any, 0);
                    byte[] response = udp.Receive(ref remoteEP);
                    return response != null && response.Length > 0;
                }
            }
            catch { return false; }
        }

        private byte[] BuildSnmpGetRequest(string community, string oid)
        {
            var communityBytes = System.Text.Encoding.ASCII.GetBytes(community);
            var oidParts = oid.Split('.').Select(s => int.Parse(s)).ToArray();

            var oidBytes = new List<byte>();
            oidBytes.Add((byte)(oidParts[0] * 40 + oidParts[1]));
            for (int i = 2; i < oidParts.Length; i++)
            {
                int val = oidParts[i];
                if (val < 128) oidBytes.Add((byte)val);
                else
                {
                    var chunks = new List<byte>();
                    while (val > 0) { chunks.Insert(0, (byte)(val & 0x7F)); val >>= 7; }
                    for (int j = 0; j < chunks.Count - 1; j++) chunks[j] |= 0x80;
                    oidBytes.AddRange(chunks);
                }
            }

            byte[] oidTlv = TlvEncode(0x06, oidBytes.ToArray());
            byte[] nullTlv = new byte[] { 0x05, 0x00 };
            byte[] varBind = TlvEncode(0x30, Combine(oidTlv, nullTlv));
            byte[] varBindList = TlvEncode(0x30, varBind);

            byte[] requestId = TlvEncode(0x02, new byte[] { 0x01 });
            byte[] errorStatus = TlvEncode(0x02, new byte[] { 0x00 });
            byte[] errorIndex = TlvEncode(0x02, new byte[] { 0x00 });
            byte[] pdu = TlvEncode(0xA0, Combine(requestId, errorStatus, errorIndex, varBindList));

            byte[] version = TlvEncode(0x02, new byte[] { 0x01 });
            byte[] comm = TlvEncode(0x04, communityBytes);
            byte[] message = TlvEncode(0x30, Combine(version, comm, pdu));

            return message;
        }

        private static byte[] TlvEncode(byte tag, byte[] value)
        {
            var result = new List<byte> { tag };
            if (value.Length < 128) result.Add((byte)value.Length);
            else
            {
                int lenBytes = value.Length < 256 ? 1 : 2;
                result.Add((byte)(0x80 | lenBytes));
                if (lenBytes == 2) result.Add((byte)(value.Length >> 8));
                result.Add((byte)(value.Length & 0xFF));
            }
            result.AddRange(value);
            return result.ToArray();
        }

        private static byte[] Combine(params byte[][] arrays)
        {
            int len = arrays.Sum(a => a.Length);
            var result = new byte[len];
            int offset = 0;
            foreach (var a in arrays) { Buffer.BlockCopy(a, 0, result, offset, a.Length); offset += a.Length; }
            return result;
        }

        private bool TestTelnet(string ip, string username, string password, int port)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    client.ReceiveTimeout = 5000;
                    client.SendTimeout = 5000;
                    var connectTask = client.ConnectAsync(ip, port);
                    if (!connectTask.Wait(5000)) return false;

                    using (var stream = client.GetStream())
                    {
                        byte[] buffer = new byte[4096];
                        stream.ReadTimeout = 5000;
                        int bytes = stream.Read(buffer, 0, buffer.Length);
                        string prompt = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();

                        if (prompt.Contains("login") || prompt.Contains("username") || prompt.Contains("user"))
                        {
                            byte[] userBytes = System.Text.Encoding.ASCII.GetBytes(username + "\r\n");
                            stream.Write(userBytes, 0, userBytes.Length);
                            System.Threading.Thread.Sleep(1000);
                            bytes = stream.Read(buffer, 0, buffer.Length);
                            prompt = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();
                        }

                        if (prompt.Contains("pass"))
                        {
                            byte[] passBytes = System.Text.Encoding.ASCII.GetBytes(password + "\r\n");
                            stream.Write(passBytes, 0, passBytes.Length);
                            System.Threading.Thread.Sleep(1000);
                            bytes = stream.Read(buffer, 0, buffer.Length);
                            string response = System.Text.Encoding.ASCII.GetString(buffer, 0, bytes).ToLower();

                            return !response.Contains("fail") && !response.Contains("denied") &&
                                   !response.Contains("incorrect") && !response.Contains("invalid") &&
                                   (response.Contains("#") || response.Contains(">") || response.Contains("$"));
                        }
                        return false;
                    }
                }
            }
            catch { return false; }
        }

        private bool TestHttp(string ip, string username, string password, int port, bool https)
        {
            try
            {
                string scheme = https ? "https" : "http";
                var request = (HttpWebRequest)WebRequest.Create($"{scheme}://{ip}:{port}/");
                request.Timeout = 5000;
                request.Credentials = new NetworkCredential(username, password);
                request.PreAuthenticate = true;

                string auth = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($"{username}:{password}"));
                request.Headers.Add("Authorization", "Basic " + auth);

                if (https)
                    System.Net.ServicePointManager.ServerCertificateValidationCallback =
                        (sender, cert, chain, errors) => true;

                using (var response = (HttpWebResponse)request.GetResponse())
                    return response.StatusCode == HttpStatusCode.OK ||
                           response.StatusCode == HttpStatusCode.Found ||
                           response.StatusCode == HttpStatusCode.Redirect;
            }
            catch (WebException ex)
            {
                if (ex.Response is HttpWebResponse resp)
                    return resp.StatusCode != HttpStatusCode.Unauthorized &&
                           resp.StatusCode != HttpStatusCode.Forbidden;
                return false;
            }
            catch { return false; }
        }
    }
}
