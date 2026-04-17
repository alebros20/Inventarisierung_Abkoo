using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace NmapInventory
{
    public class CredentialEntry
    {
        public string Alias { get; set; }
        public string IP { get; set; }
        public string Username { get; set; }
        public string EncryptedPassword { get; set; }

        public override string ToString() => $"{Alias}  ({Username}@{IP})";
    }

    public static class CredentialStore
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NmapInventory", "credentials.dat");

        // Verifikationsdatei — enthält verschlüsselten Prüftext
        private static readonly string VerifyPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NmapInventory", "credentials.verify");

        private const string VerifyMagic = "NMAP_INVENTORY_OK"; // bekannter Klartext

        // ── PBKDF2 + AES-256 ─────────────────────────────────
        private const int KeySize = 32;
        private const int IVSize = 16;
        private const int SaltSize = 16;
        private const int Iterations = 10000;

        private static byte[] DeriveKey(string password, byte[] salt)
        {
            using (var kdf = new Rfc2898DeriveBytes(password, salt, Iterations, HashAlgorithmName.SHA256))
                return kdf.GetBytes(KeySize);
        }

        // Prüft ob das Passwort korrekt ist.
        // Beim allerersten Aufruf: Verifikationsdatei anlegen und true zurückgeben.
        // Danach: Nur true wenn HMAC-Signatur stimmt.
        // Erster Aufruf — Verifikationsdatei mit Standardpasswort 1234 anlegen
        public static bool VerifyPassword(string password)
        {
            if (!File.Exists(VerifyPath))
            {
                CreateVerifyFile("1234");
                // Prüfen ob eingegebenes Passwort dem Standard entspricht  Veschlüsselung
                byte[] data = Convert.FromBase64String(File.ReadAllText(VerifyPath).Trim());
                byte[] salt = new byte[SaltSize]; Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                byte[] stored = new byte[32]; Buffer.BlockCopy(data, SaltSize, stored, 0, 32);
                byte[] computed = ComputeHmac(password, salt);
                int diff = 0;
                for (int i = 0; i < computed.Length; i++) diff |= computed[i] ^ stored[i];
                return diff == 0;
            }
            try
            {
                // Datei lesen: Salt(16) + HMAC(32)
                byte[] data = Convert.FromBase64String(File.ReadAllText(VerifyPath).Trim());
                if (data.Length < SaltSize + 32) return false;

                byte[] salt = new byte[SaltSize];
                byte[] stored = new byte[32];
                Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                Buffer.BlockCopy(data, SaltSize, stored, 0, 32);

                byte[] computed = ComputeHmac(password, salt);
                // Zeitkonstanter Vergleich gegen Timing-Angriffe
                if (computed.Length != stored.Length) return false;
                int diff = 0;
                for (int i = 0; i < computed.Length; i++)
                    diff |= computed[i] ^ stored[i];
                return diff == 0;
            }
            catch { return false; }
        }

        // Erzeugt HMAC-SHA256 aus Passwort + Salt (PBKDF2-abgeleiteter Schlüssel)
        private static byte[] ComputeHmac(string password, byte[] salt)
        {
            byte[] key = DeriveKey(password, salt);
            using (var hmac = new HMACSHA256(key))
                return hmac.ComputeHash(Encoding.UTF8.GetBytes(VerifyMagic));
        }

        // Legt die Verifikationsdatei an
        private static void CreateVerifyFile(string password)
        {
            try
            {
                byte[] salt = new byte[SaltSize];
                using (var rng = RandomNumberGenerator.Create())
                    rng.GetBytes(salt);

                byte[] hmac = ComputeHmac(password, salt);
                byte[] result = new byte[SaltSize + hmac.Length];
                Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
                Buffer.BlockCopy(hmac, 0, result, SaltSize, hmac.Length);

                Directory.CreateDirectory(Path.GetDirectoryName(VerifyPath));
                File.WriteAllText(VerifyPath, Convert.ToBase64String(result));
            }
            catch { }
        }

        // Speichert das Passwort als Verifikationsdatei (beim ersten Speichern)
        private static void SaveVerification(string password)
        {
            if (File.Exists(VerifyPath)) return;
            CreateVerifyFile(password);
        }

        public static string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            string password = SessionKey.Current;
            if (string.IsNullOrEmpty(password)) return "";
            try
            {
                byte[] salt = new byte[SaltSize];
                using (var rng = RandomNumberGenerator.Create())
                    rng.GetBytes(salt);

                byte[] key = DeriveKey(password, salt);
                using (var aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.GenerateIV();
                    using (var enc = aes.CreateEncryptor())
                    {
                        byte[] plain = Encoding.UTF8.GetBytes(plainText);
                        byte[] cipher = enc.TransformFinalBlock(plain, 0, plain.Length);
                        // Format: Salt(16) + IV(16) + CipherText
                        byte[] result = new byte[SaltSize + IVSize + cipher.Length];
                        Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
                        Buffer.BlockCopy(aes.IV, 0, result, SaltSize, IVSize);
                        Buffer.BlockCopy(cipher, 0, result, SaltSize + IVSize, cipher.Length);
                        return Convert.ToBase64String(result);
                    }
                }
            }
            catch { return ""; }
        }

        public static string Decrypt(string encryptedBase64)
        {
            if (string.IsNullOrEmpty(encryptedBase64)) return "";
            string password = SessionKey.Current;
            if (string.IsNullOrEmpty(password)) return "";
            try
            {
                byte[] data = Convert.FromBase64String(encryptedBase64);
                byte[] salt = new byte[SaltSize];
                byte[] iv = new byte[IVSize];
                byte[] cipher = new byte[data.Length - SaltSize - IVSize];
                Buffer.BlockCopy(data, 0, salt, 0, SaltSize);
                Buffer.BlockCopy(data, SaltSize, iv, 0, IVSize);
                Buffer.BlockCopy(data, SaltSize + IVSize, cipher, 0, cipher.Length);

                byte[] key = DeriveKey(password, salt);
                using (var aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    using (var dec = aes.CreateDecryptor())
                    {
                        byte[] plain = dec.TransformFinalBlock(cipher, 0, cipher.Length);
                        return Encoding.UTF8.GetString(plain);
                    }
                }
            }
            catch { return ""; }
        }

        // Liest alle Einträge roh (ohne Entschlüsselung) — für Existenzprüfung
        public static List<CredentialEntry> GetAllRaw()
        {
            var list = new List<CredentialEntry>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                foreach (var line in File.ReadAllLines(FilePath))
                {
                    var parts = line.Split('\t');
                    if (parts.Length == 4)
                        list.Add(new CredentialEntry
                        {
                            Alias = parts[0],
                            IP = parts[1],
                            Username = parts[2],
                            EncryptedPassword = parts[3]
                        });
                }
            }
            catch { }
            return list;
        }

        public static List<CredentialEntry> GetAll() => GetAllRaw();

        public static void Save(CredentialEntry entry)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));

            // Verifikationsdatei beim ersten Speichern anlegen
            SaveVerification(SessionKey.Current);
            var all = GetAll();

            // Bestehenden Eintrag mit gleichem Alias überschreiben
            var existing = all.FirstOrDefault(e =>
                e.Alias == entry.Alias || (e.IP == entry.IP && e.Username == entry.Username));
            if (existing != null) all.Remove(existing);
            all.Add(entry);

            File.WriteAllLines(FilePath, all.Select(e =>
                $"{e.Alias}\t{e.IP}\t{e.Username}\t{e.EncryptedPassword}"));
        }

        public static void Delete(string alias)
        {
            if (!File.Exists(FilePath)) return;
            var all = GetAll().Where(e => e.Alias != alias).ToList();
            File.WriteAllLines(FilePath, all.Select(e =>
                $"{e.Alias}\t{e.IP}\t{e.Username}\t{e.EncryptedPassword}"));
        }
    }

    // =========================================================
    // === SESSION-SCHLÜSSEL ===
    // Hält das Benutzerpasswort für die laufende Session im RAM.
    // Wird nie auf Disk geschrieben.
    // =========================================================
    public static class SessionKey
    {
        private static string _key = null;

        public static string Current => _key;
        public static bool IsSet => !string.IsNullOrEmpty(_key);

        public static void Set(string password) => _key = password;
        public static void Clear() => _key = null;
    }
}
