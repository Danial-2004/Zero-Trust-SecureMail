using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace SecureMailAddin
{
    public static class SecureAttachmentHandler
    {
        /// <summary>
        /// Encrypt all attachments in a MailItem before sending.
        /// </summary>
        public static void EncryptAttachments(
            MailItem mail,
             Dictionary<string, string> recipientPublicKeys)
        {
            if (mail == null || mail.Attachments.Count == 0)
                return;

            string tempDir = Path.Combine(Path.GetTempPath(), "SecureMail");
            Directory.CreateDirectory(tempDir);

            // IMPORTANT: iterate backwards (COM safety)
            for (int i = mail.Attachments.Count; i >= 1; i--)
            {
                Attachment attachment = mail.Attachments[i];

                string originalPath = Path.Combine(tempDir, attachment.FileName);
                string encryptedPath = originalPath + ".enc";

                // Save attachment to disk
                attachment.SaveAsFile(originalPath);

                // Encrypt file using hybrid encryption
                byte[] fileBytes = File.ReadAllBytes(originalPath);

                string encryptedBlob =
                    CryptoHelper.HybridEncryptForRecipients(
                        Convert.ToBase64String(fileBytes),
                        recipientPublicKeys);

                File.WriteAllText(encryptedPath, encryptedBlob);

                // Remove original attachment
                attachment.Delete();

                // Attach encrypted file
                mail.Attachments.Add(
                    encryptedPath,
                    OlAttachmentType.olByValue,
                    Type.Missing,
                    Path.GetFileName(encryptedPath));

                // Cleanup plaintext
                File.Delete(originalPath);
            }
        }

        /// <summary>
        /// Decrypt a single encrypted attachment (.enc).
        /// </summary>
        public static string DecryptAttachment(
            Attachment attachment,
            string recipientPrivateKeyXml,
            string outputFolder)
        {
            if (attachment == null)
                throw new ArgumentNullException(nameof(attachment));

            Directory.CreateDirectory(outputFolder);

            string tempEncryptedPath = Path.Combine(
                Path.GetTempPath(),
                Guid.NewGuid().ToString() + ".enc");

            // Save encrypted attachment
            attachment.SaveAsFile(tempEncryptedPath);

            // Remove .enc extension
            string originalFileName = attachment.FileName.EndsWith(".enc", StringComparison.OrdinalIgnoreCase)
                ? attachment.FileName.Substring(0, attachment.FileName.Length - 4)
                : attachment.FileName;

            string outputPath = Path.Combine(outputFolder, originalFileName);

            // Decrypt file
            string encryptedBlob = File.ReadAllText(tempEncryptedPath);

            string myEmail = GetCurrentUserEmail();

            string decryptedBase64 =
                CryptoHelper.HybridDecryptForRecipient(
                    encryptedBlob,
                    GetCurrentUserEmail(),
                    recipientPrivateKeyXml);

            byte[] fileBytes = Convert.FromBase64String(decryptedBase64);

            File.WriteAllBytes(outputPath, fileBytes);

            // Cleanup temp encrypted file
            File.Delete(tempEncryptedPath);

            return outputPath;
        }


        /// <summary>
        /// Decrypt all encrypted attachments in a MailItem.
        /// </summary>
        public static void DecryptAllAttachments(
            MailItem mail,
            string recipientPrivateKeyXml,
            string outputFolder)
        {
            if (mail == null || mail.Attachments.Count == 0)
                return;

            foreach (Attachment attachment in mail.Attachments)
            {
                if (attachment.FileName.EndsWith(".enc", StringComparison.OrdinalIgnoreCase))
                {
                    DecryptAttachment(
                        attachment,
                        recipientPrivateKeyXml,
                        outputFolder);
                }
            }
        }

        /// <summary>
        /// Helper to get the current user's email address.
        /// </summary>
        private static string GetCurrentUserEmail()
        {
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.AddressEntry addrEntry = outlookApp.Session.CurrentUser.AddressEntry;

            if (addrEntry.Type == "EX") // Exchange Account
            {
                return addrEntry.GetExchangeUser()?.PrimarySmtpAddress
                       ?? addrEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString();
            }

            return addrEntry.Address; // Standard POP/IMAP SMTP
        }
    }
}