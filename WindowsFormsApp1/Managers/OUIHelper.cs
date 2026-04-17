using System.Collections.Generic;
using System.Linq;

namespace NmapInventory
{
    // OUI helper: simple vendor lookup from MAC OUI (first 3 bytes)
    public static class OUIHelper
    {
        private static readonly Dictionary<string, string> builtIn = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
        {
            { "00:1A:2B", "Apple, Inc." },
            { "00:1A:79", "Samsung Electronics" },
            { "00:0C:29", "VMware, Inc." },
            { "84:38:35", "Google, Inc." },
            { "3C:5A:B4", "Xiaomi Communications" }
        };

        public static string Lookup(string mac)
        {
            if (string.IsNullOrWhiteSpace(mac)) return null;
            try
            {
                var clean = mac.Trim().ToUpperInvariant().Replace('-', ':').Replace('.', ':');
                clean = new string(clean.Where(c => c == ':' || (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F')).ToArray());
                if (!clean.Contains(":") && clean.Length >= 12)
                    clean = string.Join(":", System.Linq.Enumerable.Range(0, 6).Select(i => clean.Substring(i * 2, 2)));
                var parts = clean.Split(new[] { ':' }, System.StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 3) return null;
                var key = string.Join(":", parts.Take(3));
                if (builtIn.TryGetValue(key, out var vendor)) return vendor;
                return null;
            }
            catch { return null; }
        }
    }
}
