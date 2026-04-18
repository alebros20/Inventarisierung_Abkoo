using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Windows.Forms;

namespace NmapInventory
{
    public partial class MainForm
    {
        /// <summary>
        /// Ermittelt das lokale Netzwerk in CIDR-Notation (z.B. "192.168.2.0/24")
        /// durch Auswertung der aktiven Netzwerk-Interfaces.
        /// Gibt einen Fallback zurück, falls keine Netzwerk-Info gefunden wird.
        /// </summary>
        private static string GetLocalNetworkCidr()
        {
            try
            {
                var candidates = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(ni => ni.OperationalStatus == OperationalStatus.Up)
                    .Where(ni => ni.NetworkInterfaceType != NetworkInterfaceType.Loopback
                              && ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                    // Interfaces mit echtem Gateway bevorzugen (vermeidet virtuelle Adapter)
                    .OrderByDescending(ni => ni.GetIPProperties().GatewayAddresses
                        .Any(g => g.Address.AddressFamily == AddressFamily.InterNetwork
                               && !g.Address.Equals(IPAddress.Any)));

                foreach (var ni in candidates)
                {
                    foreach (var addr in ni.GetIPProperties().UnicastAddresses)
                    {
                        if (addr.Address.AddressFamily != AddressFamily.InterNetwork) continue;
                        if (IPAddress.IsLoopback(addr.Address)) continue;
                        if (addr.IPv4Mask == null) continue;

                        var ipBytes = addr.Address.GetAddressBytes();
                        var maskBytes = addr.IPv4Mask.GetAddressBytes();
                        if (maskBytes.All(b => b == 0)) continue; // Ungültige Maske

                        var netBytes = new byte[4];
                        for (int i = 0; i < 4; i++) netBytes[i] = (byte)(ipBytes[i] & maskBytes[i]);
                        int prefix = maskBytes.Sum(b => Convert.ToString(b, 2).Count(c => c == '1'));

                        return $"{netBytes[0]}.{netBytes[1]}.{netBytes[2]}.{netBytes[3]}/{prefix}";
                    }
                }
            }
            catch { }
            return "192.168.1.0/24";
        }

        // Setzt Icon + Text auf einen Button (Icon links, Text rechts)
        private static void SetButtonIcon(Button btn, string text, int iconIndex, string dll = "imageres.dll")
        {
            try
            {
                var bmp = ExtractDllIcon(dll, iconIndex);
                var img = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
                img.Images.Add(new Bitmap(bmp, 16, 16));
                btn.ImageList = img;
                btn.ImageIndex = 0;
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextAlign = ContentAlignment.MiddleRight;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            }
            catch { }
            btn.Text = text;
        }

        // Zeichnet ein einfaches Männchen-Icon als Bitmap
        private static Bitmap DrawPersonIcon(int size)
        {
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                var color = Color.FromArgb(70, 130, 180);
                using (var brush = new SolidBrush(color))
                using (var pen = new Pen(color, 1.5f))
                {
                    float s = size;
                    float headR = s * 0.18f, headX = s * 0.5f - s * 0.18f, headY = s * 0.05f;
                    g.FillEllipse(brush, headX, headY, headR * 2, headR * 2);
                    float bodyTop = headY + headR * 2 + s * 0.02f, bodyBottom = s * 0.72f;
                    float bodyW = s * 0.28f, bodyX = s * 0.5f - bodyW / 2;
                    g.FillRectangle(brush, bodyX, bodyTop, bodyW, bodyBottom - bodyTop);
                    pen.Width = size * 0.09f;
                    pen.StartCap = pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    float armY = bodyTop + (bodyBottom - bodyTop) * 0.25f;
                    g.DrawLine(pen, bodyX, armY, s * 0.08f, armY + s * 0.18f);
                    g.DrawLine(pen, bodyX + bodyW, armY, s * 0.92f, armY + s * 0.18f);
                    g.DrawLine(pen, s * 0.5f - bodyW * 0.2f, bodyBottom, s * 0.28f, s * 0.97f);
                    g.DrawLine(pen, s * 0.5f + bodyW * 0.2f, bodyBottom, s * 0.72f, s * 0.97f);
                }
            }
            return bmp;
        }

        // Standard Windows-Benutzer-Icon
        private static Bitmap GetStandardUserIcon(int size)
        {
            try
            {
                string userPic = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    @"Microsoft\User Account Pictures\user.bmp");
                if (File.Exists(userPic))
                {
                    using (var orig = new Bitmap(userPic))
                    {
                        var bmp = new Bitmap(size, size);
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.DrawImage(orig, 0, 0, size, size);
                        }
                        return bmp;
                    }
                }
                return ExtractDllIcon("shell32.dll", 265);
            }
            catch { return ExtractDllIcon("shell32.dll", 265); }
        }

        [System.Runtime.InteropServices.DllImport("shell32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        private static Bitmap ExtractDllIcon(string dllName, int index)
        {
            try
            {
                string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), dllName);
                IntPtr hIcon = ExtractIcon(IntPtr.Zero, path, index);
                if (hIcon != IntPtr.Zero)
                {
                    using (var icon = System.Drawing.Icon.FromHandle(hIcon))
                    {
                        var bmp = new Bitmap(24, 24);
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.Clear(Color.Transparent);
                            g.DrawIcon(icon, new Rectangle(0, 0, 24, 24));
                        }
                        return bmp;
                    }
                }
            }
            catch { }
            return new Bitmap(24, 24);
        }

        private static Bitmap ExtractShell32Icon(int index) => ExtractDllIcon("shell32.dll", index);

        private Label CreateStatusLabel()
            => new Label { Dock = DockStyle.Bottom, Height = 30, Text = "Bereit", BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(5) };

        private T FindControl<T>(string name) where T : Control
            => Controls.Find(name, true).FirstOrDefault() as T;
    }
}
