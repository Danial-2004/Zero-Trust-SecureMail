using Microsoft.Office.Tools.Ribbon;
using SecureMailAddin;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecureOutlookAddin
{
    public partial class SecureMailRibbon
    {
        private void SecureMailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            EncryptButton.Enabled = true;
            DecryptButton.Enabled = true;
        }

        // ================= EXPORT KEY =================
        private void ExportKeyButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string currentUser = GetCurrentUserEmail();

                if (!CryptoHelper.HasKeyPairFor(currentUser))
                    CryptoHelper.GenerateAndSaveRsaKeyPair(currentUser);

                string pubXml = CryptoHelper.GetPublicKeyXml(currentUser);

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Public Key (*.xml)|*.xml",
                    FileName = $"{currentUser}_public.xml"
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, pubXml);
                    MessageBox.Show("Public key exported successfully.", "Secure Mail");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed:\n" + ex.Message, "Secure Mail");
            }
        }

        // ================= IMPORT KEY =================
        private void ImportKeyButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = "Public Key (*.xml)|*.xml"
                };

                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                string xml = File.ReadAllText(ofd.FileName);

                string email = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter recipient email address (SMTP):",
                    "Import Public Key");

                email = NormalizeEmail(email);

                if (string.IsNullOrWhiteSpace(email))
                {
                    MessageBox.Show("Please enter a valid email address.", "Secure Mail");
                    return;
                }

                CryptoHelper.SavePublicKeyFor(email, xml);

                MessageBox.Show(
                    $"Public key imported successfully for:\n\n{email}",
                    "Secure Mail");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed:\n" + ex.Message, "Secure Mail");
            }
        }

        // ================= ENCRYPT =================
        private void EncryptButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector == null || !(inspector.CurrentItem is Outlook.MailItem mail))
                {
                    MessageBox.Show("Please open a compose email window first.", "Secure Mail");
                    return;
                }

                // ✅ SUBJECT VALIDATION (FIXES NULL REFERENCE BUG)
                if (string.IsNullOrWhiteSpace(mail.Subject))
                {
                    MessageBox.Show(
                        "Please add a subject to the email before encrypting.",
                        "Secure Mail");
                    return;
                }

                string senderEmail =
                    mail.SendUsingAccount?.SmtpAddress ??
                    Globals.ThisAddIn.Application.Session.CurrentUser.Address;

                senderEmail = NormalizeEmail(senderEmail);

                if (!CryptoHelper.HasKeyPairFor(senderEmail))
                    CryptoHelper.GenerateAndSaveRsaKeyPair(senderEmail);

                var missingKeys = new StringBuilder();
                string firstRecipient = null;

                foreach (Outlook.Recipient r in mail.Recipients)
                {
                    r.Resolve();

                    string smtp = null;

                    if (r.AddressEntry != null && r.AddressEntry.Type == "EX")
                    {
                        var exch = r.AddressEntry.GetExchangeUser();
                        smtp = exch?.PrimarySmtpAddress;
                    }
                    else
                    {
                        smtp = r.Address;
                    }

                    smtp = NormalizeEmail(smtp);
                    if (string.IsNullOrWhiteSpace(smtp))
                        continue;

                    if (firstRecipient == null)
                        firstRecipient = smtp;

                    try
                    {
                        CryptoHelper.GetPublicKeyXml(smtp);
                    }
                    catch
                    {
                        missingKeys.AppendLine("- " + smtp);
                    }
                }

                if (missingKeys.Length > 0)
                {
                    MessageBox.Show(
                        "Missing public keys for the following recipients:\n\n" +
                        missingKeys +
                        "\nPlease import the required public keys before encrypting.",
                        "Secure Mail");
                    return;
                }

                Dictionary<string, string> recipientKeys = new Dictionary<string, string>();

                foreach (Outlook.Recipient r in mail.Recipients)
                {
                    string smtp = GetSmtpAddress(r);

                    string pubKey = CryptoHelper.GetPublicKeyXml(smtp);

                    recipientKeys.Add(smtp, pubKey);
                }
                string senderPrivateKey = CryptoHelper.LoadPrivateKeyXml(senderEmail);
                string senderPublicKey = CryptoHelper.GetPublicKeyXml(senderEmail);
                recipientKeys[senderEmail] = senderPublicKey;

                mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;

                mail.Body = CryptoHelper.HybridEncryptForRecipients(
                mail.Body ?? "",
                recipientKeys,
                senderEmail,        // ✅ CORRECT
                senderPublicKey);
                if (!mail.Subject.StartsWith("[ENCRYPTED]", StringComparison.OrdinalIgnoreCase))
                    mail.Subject = "[ENCRYPTED] " + mail.Subject;

                SecureAttachmentHandler.EncryptAttachments(mail, recipientKeys);

                mail.Save();

                MessageBox.Show(
                    "Email encrypted successfully.\n\nAll recipient keys were validated.",
                    "Secure Mail");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Encryption failed:\n" + ex.Message, "Secure Mail");
            }
        }

        // ================= DECRYPT =================
        private void DecryptButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.MailItem mailItem = null;

                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector?.CurrentItem is Outlook.MailItem)
                    mailItem = (Outlook.MailItem)inspector.CurrentItem;

                if (mailItem == null)
                {
                    var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (explorer?.Selection.Count > 0 && explorer.Selection[1] is Outlook.MailItem)
                        mailItem = (Outlook.MailItem)explorer.Selection[1];
                }

                if (mailItem == null)
                {
                    MessageBox.Show("Please open or select an encrypted email first.", "Secure Mail");
                    return;
                }

                if (!mailItem.Subject.StartsWith("[ENCRYPTED]", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("This email is not marked as encrypted.", "Secure Mail");
                    return;
                }

                string currentUser = GetCurrentUserEmail().Trim().ToLower();

                if (!CryptoHelper.HasKeyPairFor(currentUser))
                {
                    MessageBox.Show("No private key found for your account.\n\nPlease generate or import your key pair first.", "Secure Mail");
                    return;
                }

                string privateKeyXml = CryptoHelper.LoadPrivateKeyXml(currentUser);

                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                string bodyText = mailItem.Body;

                // Extract secure email block safely
                int startIndex = bodyText.IndexOf("---BEGIN SECURE EMAIL---");
                int endIndex = bodyText.IndexOf("---END SECURE EMAIL---");

                if (startIndex < 0 || endIndex < 0)
                {
                    MessageBox.Show("Invalid secure email format.", "Secure Mail");
                    return;
                }

                string secureBlock = bodyText.Substring(startIndex, endIndex - startIndex + "---END SECURE EMAIL---".Length);

                // Remove only non-printable characters outside DATA block
                // HybridDecryptForRecipient() will handle Base64 inside DATA safely
                

                // -------------------- DECRYPTION --------------------
                string decryptedText = CryptoHelper.HybridDecryptForRecipient(
                    secureBlock,
                    currentUser,
                    privateKeyXml);

                // Update email body
                mailItem.Body = decryptedText + Environment.NewLine + "\n[Decrypted successfully]";
                mailItem.Save();

                // -------------------- DECRYPT ATTACHMENTS --------------------
                string outputFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "DecryptedAttachments");

                SecureAttachmentHandler.DecryptAllAttachments(mailItem, privateKeyXml, outputFolder);

                MessageBox.Show(
                    "Email decrypted successfully.\n\n" +
                    "Attachments were decrypted and saved to:\n\n" +
                    outputFolder +
                    "\n\nEncrypted attachments remain unchanged for security reasons.",
                    "Secure Mail");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Decryption failed:\n" + ex.Message, "Secure Mail");
            }
        }
        // ================= HELPERS =================
        private string GetCurrentUserEmail()
        {
            var session = Globals.ThisAddIn.Application.Session;
            var addrEntry = session.CurrentUser.AddressEntry;

            if (addrEntry.Type == "EX")
            {
                var exchUser = addrEntry.GetExchangeUser();
                return NormalizeEmail(exchUser?.PrimarySmtpAddress);
            }

            return NormalizeEmail(addrEntry.Address);
        }

        public static string GetSmtpAddress(Outlook.Recipient recipient)
        {
            if (recipient == null)
                return string.Empty;

            try
            {
                // First try Exchange API
                if (recipient.AddressEntry != null)
                {
                    if (recipient.AddressEntry.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                        recipient.AddressEntry.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        var exchUser = recipient.AddressEntry.GetExchangeUser();
                        if (exchUser != null && !string.IsNullOrEmpty(exchUser.PrimarySmtpAddress))
                            return exchUser.PrimarySmtpAddress;
                    }
                }

                // Second try MAPI property
                var smtp = recipient.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E");

                if (smtp != null)
                    return smtp.ToString();
            }
            catch { }

            // Final fallback
            return recipient.Address;
        }

        private static string NormalizeEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return null;

            email = email.Trim().ToLowerInvariant();

            // Remove common prefixes
            if (email.StartsWith("smtp:"))
                email = email.Substring(5);

            if (email.StartsWith("mailto:"))
                email = email.Substring(7);

            // Remove Outlook suffixes
            if (email.Contains(";"))
                email = email.Split(';')[0];

            // Handle Exchange Distinguished Names
            if (email.Contains("/o="))
            {
                int idx = email.LastIndexOf("cn=");
                if (idx != -1)
                {
                    email = email.Substring(idx + 3);
                }
            }
            return email.Trim();
        }
    }
}
