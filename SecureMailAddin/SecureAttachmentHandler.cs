using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace SecureMailAddin
{
    public static class SecureAttachmentHandler
    {
        /// <summary>
        /// Encrypts every attachment and removes both the plaintext and the
        /// temporary encrypted copy from disk. The sender is also added to the
        /// recipient-key set so the sender can decrypt the sent copy later.
        /// </summary>
        public static void EncryptAttachments(
            MailItem mail,
            Dictionary<string, string> recipientPublicKeys)
        {
            if (mail == null || mail.Attachments == null || mail.Attachments.Count == 0)
                return;

            if (recipientPublicKeys == null || recipientPublicKeys.Count == 0)
                throw new ArgumentException("At least one trusted recipient public key is required.", nameof(recipientPublicKeys));

            string senderEmail = GetCurrentUserEmail();
            string senderPublicKeyXml = CryptoHelper.GetPublicKeyXml(senderEmail);

            if (string.IsNullOrWhiteSpace(senderPublicKeyXml))
                throw new InvalidOperationException("The sender public key is unavailable; encryption was aborted.");

            string tempDir = Path.Combine(Path.GetTempPath(), "SecureMail");
            Directory.CreateDirectory(tempDir);

            // Iterate backwards for COM collection safety.
            for (int i = mail.Attachments.Count; i >= 1; i--)
            {
                string originalPath = null;
                string encryptedPath = null;

                try
                {
                    Attachment attachment = mail.Attachments[i];
                    if (attachment == null || string.IsNullOrWhiteSpace(attachment.FileName))
                        throw new InvalidOperationException("An attachment has no usable file name.");

                    string safeName = Path.GetFileName(attachment.FileName);
                    if (string.IsNullOrWhiteSpace(safeName))
                        throw new InvalidOperationException("An attachment has an unsafe file name.");

                    originalPath = Path.Combine(tempDir, Guid.NewGuid().ToString("N") + "_" + safeName);
                    encryptedPath = originalPath + ".enc";

                    attachment.SaveAsFile(originalPath);

                    byte[] fileBytes = File.ReadAllBytes(originalPath);
                    string encryptedBlob = CryptoHelper.HybridEncryptForRecipients(
                        Convert.ToBase64String(fileBytes),
                        recipientPublicKeys,
                        senderEmail,
                        senderPublicKeyXml);

                    File.WriteAllText(encryptedPath, encryptedBlob);

                    attachment.Delete();

                    mail.Attachments.Add(
                        encryptedPath,
                        OlAttachmentType.olByValue,
                        Type.Missing,
                        Path.GetFileName(encryptedPath));
                }
                finally
                {
                    // Both files are plaintext/encrypted local artifacts. Remove
                    // them even when encryption or Outlook attachment fails.
                    TryDelete(originalPath);
                    TryDelete(encryptedPath);
                }
            }
        }

        /// <summary>
        /// Decrypts one .enc attachment into the requested output folder.
        /// </summary>
        public static string DecryptAttachment(
            Attachment attachment,
            string recipientPrivateKeyXml,
            string outputFolder)
        {
            if (attachment == null)
                throw new ArgumentNullException(nameof(attachment));

            if (string.IsNullOrWhiteSpace(recipientPrivateKeyXml))
                throw new ArgumentException("A recipient private key is required.", nameof(recipientPrivateKeyXml));

            Directory.CreateDirectory(outputFolder);

            string tempEncryptedPath = Path.Combine(
                Path.GetTempPath(),
                Guid.NewGuid().ToString("N") + ".enc");

            try
            {
                attachment.SaveAsFile(tempEncryptedPath);

                string originalFileName = attachment.FileName != null &&
                    attachment.FileName.EndsWith(".enc", StringComparison.OrdinalIgnoreCase)
                    ? attachment.FileName.Substring(0, attachment.FileName.Length - 4)
                    : attachment.FileName;

                originalFileName = Path.GetFileName(originalFileName ?? string.Empty);
                if (string.IsNullOrWhiteSpace(originalFileName))
                    throw new InvalidOperationException("The encrypted attachment has no usable file name.");

                string outputPath = Path.Combine(outputFolder, originalFileName);

                string encryptedBlob = File.ReadAllText(tempEncryptedPath);
                string myEmail = GetCurrentUserEmail();

                string decryptedBase64 = CryptoHelper.HybridDecryptForRecipient(
                    encryptedBlob,
                    myEmail,
                    recipientPrivateKeyXml);

                byte[] fileBytes = Convert.FromBase64String(decryptedBase64);
                File.WriteAllBytes(outputPath, fileBytes);

                return outputPath;
            }
            finally
            {
                TryDelete(tempEncryptedPath);
            }
        }

        /// <summary>
        /// Decrypts all .enc attachments in a MailItem.
        /// </summary>
        public static void DecryptAllAttachments(
            MailItem mail,
            string recipientPrivateKeyXml,
            string outputFolder)
        {
            if (mail == null || mail.Attachments == null || mail.Attachments.Count == 0)
                return;

            for (int i = mail.Attachments.Count; i >= 1; i--)
            {
                Attachment attachment = mail.Attachments[i];
                if (attachment != null &&
                    attachment.FileName != null &&
                    attachment.FileName.EndsWith(".enc", StringComparison.OrdinalIgnoreCase))
                {
                    DecryptAttachment(
                        attachment,
                        recipientPrivateKeyXml,
                        outputFolder);
                }
            }
        }

        private static void TryDelete(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // Cleanup must not mask the original encryption/decryption error.
                // A production deployment should also use a protected temp folder.
            }
        }

        private static string GetCurrentUserEmail()
        {
            Application outlookApp = new Application();
            AddressEntry addrEntry = outlookApp.Session.CurrentUser.AddressEntry;

            if (addrEntry.Type == "EX")
            {
                ExchangeUser exchangeUser = addrEntry.GetExchangeUser();
                if (exchangeUser != null && !string.IsNullOrWhiteSpace(exchangeUser.PrimarySmtpAddress))
                    return exchangeUser.PrimarySmtpAddress;

                object smtpAddress = addrEntry.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
                return smtpAddress == null ? string.Empty : smtpAddress.ToString();
            }

            return addrEntry.Address ?? string.Empty;
        }
    }
}
