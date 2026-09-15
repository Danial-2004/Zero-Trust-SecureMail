using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SecureMailAddin
{
    public static class CryptoHelper
    {
        private static readonly string KeyFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SecureOutlookAddin", "keys");

        static CryptoHelper()
        {
            if (!Directory.Exists(KeyFolder)) Directory.CreateDirectory(KeyFolder);
        }

        // ---------------- RSA XML helpers ----------------
        public static string RsaToXmlString(RSA rsa, bool includePrivateParams)
        {
            var csp = new RSACryptoServiceProvider();
            csp.ImportParameters(rsa.ExportParameters(includePrivateParams));
            return csp.ToXmlString(includePrivateParams);
        }

        public static void RsaFromXmlString(RSA rsa, string xml)
        {
            if (rsa == null) throw new ArgumentNullException(nameof(rsa));
            if (string.IsNullOrWhiteSpace(xml)) throw new ArgumentNullException(nameof(xml));

            bool hasPrivate = xml.Contains("<D>") || xml.Contains("<P>") || xml.Contains("<Q>");

            var csp = new RSACryptoServiceProvider();
            try
            {
                csp.FromXmlString(xml);
                RSAParameters parameters = csp.ExportParameters(hasPrivate);
                rsa.ImportParameters(parameters);
            }
            finally
            {
                csp.PersistKeyInCsp = false;
                csp.Clear();
            }
        }

        // ---------------- Key management ----------------
        public static void GenerateAndSaveRsaKeyPair(string username, int keySize = 2048)
        {
            using (RSA rsa = RSA.Create(keySize))
            {
                string pub = RsaToXmlString(rsa, false);
                string priv = RsaToXmlString(rsa, true);

                File.WriteAllText(GetPubKeyPath(username), pub);

                byte[] privBytes = Encoding.UTF8.GetBytes(priv);
                byte[] protectedPriv = ProtectedData.Protect(privBytes, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(GetPrivKeyPath(username), protectedPriv);
            }
        }

        public static bool HasKeyPairFor(string username)
        {
            return File.Exists(GetPubKeyPath(username)) &&
                   File.Exists(GetPrivKeyPath(username));
        }

        public static string LoadPublicKeyXml(string username)
        {
            string path = GetPubKeyPath(username);
            if (!File.Exists(path)) throw new FileNotFoundException($"Public key not found for {username}");
            return File.ReadAllText(path);
        }

        public static string LoadPrivateKeyXml(string username)
        {
            string path = GetPrivKeyPath(username);
            if (!File.Exists(path)) throw new FileNotFoundException($"Private key not found for {username}");

            byte[] protectedPriv = File.ReadAllBytes(path);
            byte[] privBytes;
            try
            {
                privBytes = ProtectedData.Unprotect(protectedPriv, null, DataProtectionScope.CurrentUser);
            }
            catch (CryptographicException ex)
            {
                throw new CryptographicException("Unable to unprotect private key. Make sure you're running under the same Windows user account that created the key. " + ex.Message, ex);
            }

            return Encoding.UTF8.GetString(privBytes);
        }

        private static string GetPubKeyPath(string username) =>
            Path.Combine(KeyFolder, MakeSafeFileName(username) + ".pub.xml");

        private static string GetPrivKeyPath(string username) =>
            Path.Combine(KeyFolder, MakeSafeFileName(username) + ".priv.prot");

        private static string MakeSafeFileName(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            return s;
        }

        public static RSA GetOrCreateRsaFor(string username)
        {
            RSA rsa = RSA.Create();
            if (HasKeyPairFor(username))
            {
                string privXml = LoadPrivateKeyXml(username);
                RsaFromXmlString(rsa, privXml);
            }
            else
            {
                GenerateAndSaveRsaKeyPair(username);
                string privXml = LoadPrivateKeyXml(username);
                RsaFromXmlString(rsa, privXml);
            }
            return rsa;
        }

        // ================= HYBRID ENCRYPTION (TEXT) =================
        public static string HybridEncryptString(string text, string recipientPublicKeyXml)
        {
            byte[] data = Encoding.UTF8.GetBytes(text);
            return HybridEncryptBytes(data, recipientPublicKeyXml);
        }

        public static string HybridDecryptString(string base64, string recipientPrivateKeyXml)
        {
            byte[] bytes = HybridDecryptBytes(base64, recipientPrivateKeyXml);
            return Encoding.UTF8.GetString(bytes);
        }

        // ================= HYBRID ENCRYPTION (BYTES / FILES) =================
        //
        // Format:
        // [int encKeyLen][encKey]
        // [int ivLen][iv]
        // [int cipherLen][cipher]
        //

        public static string HybridEncryptBytes(byte[] data, string recipientPublicKeyXml)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrWhiteSpace(recipientPublicKeyXml)) throw new ArgumentNullException(nameof(recipientPublicKeyXml));

            byte[] encryptedMessage;
            byte[] aesKey;
            byte[] aesIV;

            using (Aes aes = Aes.Create())
            {
                aes.KeySize = 256;
                aes.GenerateKey();
                aes.GenerateIV();
                aesKey = aes.Key;
                aesIV = aes.IV;

                using (MemoryStream ms = new MemoryStream())
                using (CryptoStream cs = new CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(data, 0, data.Length);
                    cs.FlushFinalBlock();
                    encryptedMessage = ms.ToArray();
                }
            }

            byte[] encryptedKey = null;

            // Try OAEP SHA-256 first (preferred), fallback to OAEP SHA-1 if not supported
            using (RSA rsa = RSA.Create())
            {
                RsaFromXmlString(rsa, recipientPublicKeyXml);

                try
                {
                    encryptedKey = rsa.Encrypt(aesKey, RSAEncryptionPadding.OaepSHA256);
                }
                catch (CryptographicException)
                {
                    encryptedKey = rsa.Encrypt(aesKey, RSAEncryptionPadding.OaepSHA1);
                }
            }

            using (MemoryStream final = new MemoryStream())
            using (BinaryWriter bw = new BinaryWriter(final))
            {
                bw.Write(encryptedKey.Length);
                bw.Write(encryptedKey);

                bw.Write(aesIV.Length);
                bw.Write(aesIV);

                bw.Write(encryptedMessage.Length);
                bw.Write(encryptedMessage);

                return Convert.ToBase64String(final.ToArray());
            }
        }

        public static byte[] HybridDecryptBytes(string base64Blob, string recipientPrivateKeyXml)
        {
            if (string.IsNullOrWhiteSpace(base64Blob)) throw new ArgumentNullException(nameof(base64Blob));
            if (string.IsNullOrWhiteSpace(recipientPrivateKeyXml)) throw new ArgumentNullException(nameof(recipientPrivateKeyXml));

            byte[] full = Convert.FromBase64String(base64Blob);
            using (MemoryStream ms = new MemoryStream(full))
            using (BinaryReader br = new BinaryReader(ms))
            {
                int encKeyLen = br.ReadInt32();
                byte[] encryptedKey = br.ReadBytes(encKeyLen);

                int ivLen = br.ReadInt32();
                byte[] iv = br.ReadBytes(ivLen);

                int cipherLen = br.ReadInt32();
                byte[] cipher = br.ReadBytes(cipherLen);

                byte[] aesKey = null;

                // Try decrypt with OAEP-SHA256, fallback to OAEP-SHA1
                using (RSA rsa = RSA.Create())
                {
                    RsaFromXmlString(rsa, recipientPrivateKeyXml);

                    try
                    {
                        aesKey = rsa.Decrypt(encryptedKey, RSAEncryptionPadding.OaepSHA256);
                    }
                    catch (CryptographicException)
                    {
                        aesKey = rsa.Decrypt(encryptedKey, RSAEncryptionPadding.OaepSHA1);
                    }
                }

                using (Aes aes = Aes.Create())
                {
                    aes.Key = aesKey;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;

                    using (var cms = new MemoryStream(cipher))
                    using (var cs = new CryptoStream(cms, aes.CreateDecryptor(), CryptoStreamMode.Read))
                    using (var outMs = new MemoryStream())
                    {
                        cs.CopyTo(outMs);
                        return outMs.ToArray();
                    }
                }
            }
        }

        // ---------------- Backward compatibility ----------------
        public static string NormalizeKeyId(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                throw new ArgumentNullException(nameof(email));

            email = email.Trim().ToLowerInvariant();

            foreach (char c in Path.GetInvalidFileNameChars())
                email = email.Replace(c, '_');

            return email;
        }

        public static string GetPublicKeyXml(string email)
        {
            string keyId = NormalizeKeyId(email);
            return LoadPublicKeyXml(keyId);
        }

        public static void SavePublicKeyFor(string email, string publicKeyXml)
        {
            Directory.CreateDirectory(KeyFolder);
            string keyId = NormalizeKeyId(email);
            File.WriteAllText(GetPubKeyPath(keyId), publicKeyXml);
        }

        // ================== ENTERPRISE SECURE MAIL FUNCTIONS ==================

        public static string HybridEncryptForRecipients(
            string plainText,
            Dictionary<string, string> recipientPublicKeys,
            string senderEmail = null,
            string senderPublicKeyXml = null)
        {
            if (plainText == null) plainText = "";
            if (recipientPublicKeys == null || recipientPublicKeys.Count == 0)
                throw new Exception("No recipients provided");

            // Include sender automatically
            if (!string.IsNullOrWhiteSpace(senderEmail) && !string.IsNullOrWhiteSpace(senderPublicKeyXml))
            {
                string senderKeyId = NormalizeEmail(senderEmail);
                if (!recipientPublicKeys.ContainsKey(senderKeyId))
                    recipientPublicKeys[senderKeyId] = senderPublicKeyXml;
            }

            byte[] data = Encoding.UTF8.GetBytes(plainText);

            byte[] encryptedMessage;
            byte[] aesKey;
            byte[] aesIV;

            // --- Generate AES key/IV and encrypt message ---
            using (Aes aes = Aes.Create())
            {
                aes.KeySize = 256;
                aes.GenerateKey();
                aes.GenerateIV();
                aesKey = aes.Key;
                aesIV = aes.IV;

                using (MemoryStream ms = new MemoryStream())
                using (CryptoStream cs = new CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(data, 0, data.Length);
                    cs.FlushFinalBlock();
                    encryptedMessage = ms.ToArray();
                }
            }

            // --- Compute HMAC for integrity ---
            byte[] hmac;
            using (var h = new HMACSHA256(aesKey))
            {
                byte[] combined = new byte[aesIV.Length + encryptedMessage.Length];
                Buffer.BlockCopy(aesIV, 0, combined, 0, aesIV.Length);
                Buffer.BlockCopy(encryptedMessage, 0, combined, aesIV.Length, encryptedMessage.Length);
                hmac = h.ComputeHash(combined);
            }

            // --- Wrap AES key for each recipient ---
            StringBuilder keyBlock = new StringBuilder();
            foreach (var r in recipientPublicKeys)
            {
                string email = r.Key;
                string pubKeyXml = r.Value;

                using (RSA rsa = RSA.Create())
                {
                    RsaFromXmlString(rsa, pubKeyXml);
                    byte[] wrappedKey;

                    try { wrappedKey = rsa.Encrypt(aesKey, RSAEncryptionPadding.OaepSHA256); }
                    catch { wrappedKey = rsa.Encrypt(aesKey, RSAEncryptionPadding.OaepSHA1); }

                    string wrappedKeyB64 = Convert.ToBase64String(wrappedKey, Base64FormattingOptions.None);
                    keyBlock.AppendLine($"{NormalizeEmail(email)}:{wrappedKeyB64}");
                }
            }

            // --- Build secure email package ---
            StringBuilder package = new StringBuilder();
            package.AppendLine("---BEGIN SECURE EMAIL---");
            package.AppendLine("VERSION:1");
            package.AppendLine("ALGO:AES256");
            package.AppendLine("HMAC:" + Convert.ToBase64String(hmac));
            package.AppendLine("RECIPIENT_KEYS");
            package.Append(keyBlock.ToString());
            package.AppendLine("IV:" + Convert.ToBase64String(aesIV, Base64FormattingOptions.None));
            string dataB64 = Convert.ToBase64String(encryptedMessage, Base64FormattingOptions.None);
            dataB64 = dataB64.Replace("\r", "").Replace("\n", "");

            package.AppendLine("DATA:" + dataB64); package.AppendLine("---END SECURE EMAIL---");

            return package.ToString();
        }

        // --- Single-recipient version with HMAC ---
        public static string HybridEncryptForRecipient(
            string plainText,
            string recipientPublicKeyXml,
            string senderEmail = null,
            string senderPublicKeyXml = null)
        {
            var recipients = new Dictionary<string, string>();
            recipients["recipient"] = recipientPublicKeyXml;
            return HybridEncryptForRecipients(plainText, recipients, senderEmail, senderPublicKeyXml);
        }

        // ================== DECRYPTION WRAPPER FOR LEGACY CALLS ==================
        public static (string PlainText, bool SignatureValid, string SenderPublicKeyXml) HybridDecryptAsRecipient(
            string secureMessage,
            string myEmail,
            string recipientPrivateKeyXml)
        {
            // Reuse full HMAC-verified decryption
            string plainText = HybridDecryptForRecipient(secureMessage, myEmail, recipientPrivateKeyXml);

            // SignatureValid is true because HMAC check passed
            // SenderPublicKeyXml is empty for now (extend if you add signing)
            return (plainText, true, string.Empty);
        }

        // ================== FULL SECUREMAIL DECRYPTION ==================
        public static string HybridDecryptForRecipient(
    string secureMessage,
    string myEmail,
    string recipientPrivateKeyXml)
        {
            if (string.IsNullOrWhiteSpace(secureMessage))
                throw new ArgumentNullException(nameof(secureMessage));

            if (!secureMessage.Contains("---BEGIN SECURE EMAIL---"))
                throw new Exception("Invalid secure email format");

            string normalizedUser = NormalizeEmail(myEmail);

            string encryptedKey = null;
            StringBuilder currentKey = new StringBuilder();
            bool collectingKey = false;
            string currentKeyEmail = null;
            string iv = null;
            string hmacValue = null;

            StringBuilder dataBuilder = new StringBuilder();

            ParseSection section = ParseSection.None;

            string[] lines = secureMessage.Split(new[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (string raw in lines)
            {
                string line = raw.Trim();

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // ----------------------------
                // Section switches
                // ----------------------------

                if (line.Equals("RECIPIENT_KEYS", StringComparison.OrdinalIgnoreCase))
                {
                    section = ParseSection.RecipientKeys;
                    continue;
                }

                if (line.StartsWith("IV:", StringComparison.OrdinalIgnoreCase))
                {
                    // Save last wrapped key before leaving RECIPIENT_KEYS
                    if (collectingKey && currentKeyEmail == normalizedUser)
                        encryptedKey = currentKey.ToString();

                    collectingKey = false;

                    iv = line.Substring(3).Trim();
                    section = ParseSection.None;
                    continue;
                }

                if (line.StartsWith("DATA:", StringComparison.OrdinalIgnoreCase))
                {
                    section = ParseSection.Data;
                    dataBuilder.Append(line.Substring(5).Trim());
                    continue;
                }

                if (line.StartsWith("VERSION:", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (line.StartsWith("ALGO:", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (line.StartsWith("HMAC:", StringComparison.OrdinalIgnoreCase))
                {
                    hmacValue = line.Substring(5).Trim();
                    continue;
                }

                if (line.StartsWith("---END"))
                    break;

                // ----------------------------
                // Recipient Keys
                // ----------------------------

                if (section == ParseSection.RecipientKeys)
                {
                    int idx = line.IndexOf(':');

                    if (idx > 0)
                    {
                        // Store previous recipient key
                        if (collectingKey && currentKeyEmail == normalizedUser)
                            encryptedKey = currentKey.ToString();

                        currentKey.Clear();

                        currentKeyEmail = NormalizeEmail(line.Substring(0, idx));

                        currentKey.Append(line.Substring(idx + 1).Trim());

                        collectingKey = true;

                        continue;
                    }

                    // Wrapped Base64 continuation
                    if (collectingKey)
                    {
                        currentKey.Append(line.Trim());
                    }

                    continue;
                }

                // ----------------------------
                // DATA
                // ----------------------------

                if (section == ParseSection.Data)
                {
                    dataBuilder.Append(line.Trim());
                }
            }

            // Save last recipient key if file ended immediately before IV
            if (collectingKey && currentKeyEmail == normalizedUser)
                encryptedKey = currentKey.ToString();

            if (string.IsNullOrWhiteSpace(encryptedKey))
                throw new Exception("No encrypted key found for this recipient.");

            string data = dataBuilder.ToString();

            // Clean Base64
            encryptedKey = Regex.Replace(encryptedKey, @"\s+", "");
            iv = Regex.Replace(iv ?? "", @"\s+", "");
            data = Regex.Replace(data, @"\s+", "");
            hmacValue = Regex.Replace(hmacValue ?? "", @"\s+", "");

            //---------------------------------------------
            // Debug output
            //---------------------------------------------

            Debug.WriteLine("=========== Secure Mail Debug ===========");
            Debug.WriteLine("Recipient : " + normalizedUser);
            Debug.WriteLine("EncryptedKey Length : " + encryptedKey.Length);
            Debug.WriteLine("EncryptedKey Mod4   : " + (encryptedKey.Length % 4));
            Debug.WriteLine("IV Length           : " + iv.Length);
            Debug.WriteLine("Data Length         : " + data.Length);
            Debug.WriteLine("HMAC Length         : " + hmacValue.Length);

            //---------------------------------------------
            // Validate
            //---------------------------------------------

            try
            {
                Convert.FromBase64String(encryptedKey);
            }
            catch (FormatException)
            {
                throw new Exception("Recipient key is corrupted or not valid Base64.");
            }

            if (iv.Length % 4 != 0)
                throw new Exception("IV is not valid Base64.");

            if (data.Length % 4 != 0)
                throw new Exception("Encrypted data is not valid Base64.");

            if (hmacValue.Length % 4 != 0)
                throw new Exception("HMAC is not valid Base64.");

            //---------------------------------------------
            // RSA decrypt AES key
            //---------------------------------------------

            byte[] aesKey;

            using (RSA rsa = RSA.Create())
            {
                RsaFromXmlString(rsa, recipientPrivateKeyXml);

                try
                {
                    aesKey = rsa.Decrypt(
                        Convert.FromBase64String(encryptedKey),
                        RSAEncryptionPadding.OaepSHA256);
                }
                catch
                {
                    aesKey = rsa.Decrypt(
                        Convert.FromBase64String(encryptedKey),
                        RSAEncryptionPadding.OaepSHA1);
                }
            }

            //---------------------------------------------
            // Decode Base64
            //---------------------------------------------

            byte[] ivBytes = Convert.FromBase64String(iv);
            byte[] cipher = Convert.FromBase64String(data);

            //---------------------------------------------
            // Verify HMAC
            //---------------------------------------------

            byte[] expectedHmac = Convert.FromBase64String(hmacValue);

            byte[] combined = new byte[ivBytes.Length + cipher.Length];

            Buffer.BlockCopy(ivBytes, 0, combined, 0, ivBytes.Length);
            Buffer.BlockCopy(cipher, 0, combined, ivBytes.Length, cipher.Length);

            byte[] computedHmac;

            using (var h = new HMACSHA256(aesKey))
            {
                computedHmac = h.ComputeHash(combined);
            }

            if (!FixedTimeEquals(expectedHmac, computedHmac))
                throw new Exception("Message integrity check failed.");

            //---------------------------------------------
            // AES decrypt
            //---------------------------------------------

            using (Aes aes = Aes.Create())
            {
                aes.Key = aesKey;
                aes.IV = ivBytes;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (MemoryStream ms = new MemoryStream(cipher))
                using (CryptoStream cs = new CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Read))
                using (MemoryStream outMs = new MemoryStream())
                {
                    cs.CopyTo(outMs);
                    return Encoding.UTF8.GetString(outMs.ToArray());
                }
            }
        }

        // ================== FIXED-TIME COMPARISON ==================
        private static bool FixedTimeEquals(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length)
                return false;

            int diff = 0;
            for (int i = 0; i < a.Length; i++)
                diff |= a[i] ^ b[i];

            return diff == 0;
        }

        // ================= HYBRID ENCRYPTION (FILES) =================

        public static void HybridEncryptFile(
            string inputFilePath,
            string outputFilePath,
            string recipientPublicKeyXml)
        {
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException("Input file not found", inputFilePath);

            byte[] fileBytes = File.ReadAllBytes(inputFilePath);

            // Encrypt bytes using existing hybrid logic
            string encryptedBase64 = HybridEncryptBytes(fileBytes, recipientPublicKeyXml);

            // Store as text (Base64)
            File.WriteAllText(outputFilePath, encryptedBase64, Encoding.UTF8);
        }

        public static void HybridDecryptFile(
            string encryptedFilePath,
            string outputFilePath,
            string recipientPrivateKeyXml)
        {
            if (!File.Exists(encryptedFilePath))
                throw new FileNotFoundException("Encrypted file not found", encryptedFilePath);

            string encryptedBase64 = File.ReadAllText(encryptedFilePath, Encoding.UTF8);

            // Decrypt back to bytes
            byte[] decryptedBytes = HybridDecryptBytes(encryptedBase64, recipientPrivateKeyXml);

            File.WriteAllBytes(outputFilePath, decryptedBytes);
        }

        public static string LoadPublicKeyFromFile(string keyFilePath)
        {
            if (!File.Exists(keyFilePath))
                throw new FileNotFoundException("Public key file not found", keyFilePath);

            return File.ReadAllText(keyFilePath);
        }

        public static string GetPublicKeyForEmailFromFolder(string email)
        {
            email = NormalizeEmail(email);

            string[] files = Directory.GetFiles(KeyFolder, "*.pub.xml");

            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file).ToLowerInvariant();

                if (fileName.StartsWith(email.ToLowerInvariant()))
                {
                    return File.ReadAllText(file);
                }
            }

            throw new FileNotFoundException($"Public key not found for {email} in {KeyFolder}");
        }

        private static string NormalizeEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return null;
            email = email.Trim().ToLowerInvariant();
            if (email.StartsWith("smtp:")) email = email.Substring(5);
            if (email.Contains("/o="))
            {
                int i = email.LastIndexOf("cn=");
                if (i != -1) email = email.Substring(i + 3);
            }
            return email.Trim();
        }

        private enum ParseSection
        {
            None,
            RecipientKeys,
            Data
        }
    }
}
