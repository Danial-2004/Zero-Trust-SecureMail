using SecureMailAddin.ZeroTrust;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using SecureMailAddin.Keycloak;
using System.Collections.Generic;

namespace SecureMailAddin.MailHandler
{
    public class MailHandlerModule
    {
        private readonly MailClassificationEngine _classificationEngine;

        private readonly ApiIntegrationService _apiService;

           public MailHandlerModule(string companyDomain)
        {
            _classificationEngine = new MailClassificationEngine(companyDomain);
            _apiService = new ApiIntegrationService();
        }

        public MailAnalysisResult AnalyzeMail(Outlook.MailItem mail)
        {
            return _classificationEngine.Analyze(mail);
        }

        public bool HandleBeforeSend(Outlook.MailItem mail)
        {
            try
            {
                if (mail == null)
                    return true;

                Logger.LogMailBeforeSendStarted(mail?.Subject);

                if (!IsMailEncrypted(mail))
                {
                    Logger.LogMailEncryptionCheckFailed(mail?.Subject);

                    System.Windows.Forms.MessageBox.Show(
                        "All emails must be encrypted before sending.\n\nPlease encrypt the email before sending.",
                        "Secure Mail Policy",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    return false;
                }

                MailAnalysisResult result = AnalyzeMail(mail);
                ApplyClassification(mail, result);

                Logger.LogMailClassificationResult(
                    mail?.Subject,
                    result.SuggestedTag,
                    result.SensitivityLevel.ToString(),
                    result.TotalScore);

                Logger.LogMailZeroTrustCheckStarted(mail?.Subject, result.SuggestedTag);

                bool allowed = EnforceZeroTrustBeforeSend(mail?.Subject, result.SuggestedTag);
                if (!allowed)
                {
                    return false;
                }

                if (!ValidateRecipients(mail, result))
                {
                    return false;
                }

                try
                {
                    var workflow = new SecureMailAddin.Keycloak.Module5WorkflowExample();

                    var classificationResult = new SecureMailAddin.Keycloak.ClassificationResult
                    {
                        Score = result.TotalScore,
                        Label = string.IsNullOrWhiteSpace(result.SuggestedTag)
                            ? "PUBLIC"
                            : result.SuggestedTag.ToUpperInvariant(),
                        RequiresEncryption = result.RequiresEncryption
                    };

                    workflow.HandleBeforeSendAsync(mail, classificationResult)
                            .GetAwaiter()
                            .GetResult();
                }
                catch
                {
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to process the email before sending.\n\n" + ex.Message,
                    "Secure Mail",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

                return false;
            }
        }

        public bool EnforceZeroTrustBeforeSend(string subject, string suggestedTag)
        {
            if (string.IsNullOrWhiteSpace(suggestedTag))
                return true;

            ProtectedActionType action;

            if (suggestedTag.Equals("CONFIDENTIAL", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendConfidentialMail;
            }
            else if (suggestedTag.Equals("RESTRICTED", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendRestrictedMail;
            }
            else
            {
                Logger.LogMailAllowedByZeroTrust(subject, suggestedTag);
                return true;
            }

            var zeroTrust = new ZeroTrustHandler();
            var decision = zeroTrust.AuthorizeAsync(action).GetAwaiter().GetResult();

            if (!decision.IsAllowed)
            {
                Logger.LogMailBlockedByZeroTrust(subject, suggestedTag, decision.Reason);

                System.Windows.Forms.MessageBox.Show(
                    "Zero Trust policy denied this action.\n\n" + decision.Reason,
                    "Access Denied",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

                return false;
            }

            Logger.LogMailAllowedByZeroTrust(subject, suggestedTag);
            return true;
        }

        private bool IsMailEncrypted(Outlook.MailItem mail)
        {
            if (mail == null)
                return false;

            string subject = mail.Subject ?? string.Empty;

            return Regex.IsMatch(subject, @"^\s*\[ENCRYPTED\]\s*", RegexOptions.IgnoreCase);
        }

        private void ApplyClassification(Outlook.MailItem mail, MailAnalysisResult result)
        {
            Outlook.UserProperties props = mail.UserProperties;

            SetOrCreateUserProperty(props, "SMClass", result.SuggestedTag);
            SetOrCreateUserProperty(props, "SMLevel", result.SensitivityLevel.ToString());
            SetOrCreateUserProperty(props, "SMReason", result.Reason ?? "");
            SetOrCreateUserProperty(props, "SMHasExternalRecipients", result.HasExternalRecipients.ToString());
            SetOrCreateUserProperty(props, "SMHasAttachments", result.HasAttachments.ToString());
            SetOrCreateUserProperty(props, "SMContainsSensitiveAttachmentData", result.ContainsSensitiveAttachmentData.ToString());
            SetOrCreateUserProperty(props, "SMRestrictedKeywordHits", result.RestrictedKeywordHits.ToString());
            SetOrCreateUserProperty(props, "SMConfidentialKeywordHits", result.ConfidentialKeywordHits.ToString());
            SetOrCreateUserProperty(props, "SMTotalScore", result.TotalScore.ToString());

            ApplyOutlookSensitivity(mail, result);
            ApplyCategory(mail, result);
            ApplySubjectClassification(mail, result);
        }

        private void SetOrCreateUserProperty(Outlook.UserProperties props, string name, string value)
        {
            Outlook.UserProperty prop = null;

            try
            {
                prop = props.Find(name);
            }
            catch
            {
            }

            if (prop == null)
                prop = props.Add(name, Outlook.OlUserPropertyType.olText, true);

            prop.Value = value ?? string.Empty;
        }

        private void ApplyOutlookSensitivity(Outlook.MailItem mail, MailAnalysisResult result)
        {
            switch (result.SensitivityLevel)
            {
                case MailSensitivityLevel.Restricted:
                case MailSensitivityLevel.Confidential:
                    mail.Sensitivity = Outlook.OlSensitivity.olConfidential;
                    break;

                case MailSensitivityLevel.Internal:
                    mail.Sensitivity = Outlook.OlSensitivity.olPrivate;
                    break;

                default:
                    mail.Sensitivity = Outlook.OlSensitivity.olNormal;
                    break;
            }
        }

        private void ApplyCategory(Outlook.MailItem mail, MailAnalysisResult result)
        {
            if (mail == null)
                return;

            var categories = (mail.Categories ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(c => c.Trim())
                .Where(c =>
                    !c.Equals("PUBLIC", StringComparison.OrdinalIgnoreCase) &&
                    !c.Equals("INTERNAL", StringComparison.OrdinalIgnoreCase) &&
                    !c.Equals("CONFIDENTIAL", StringComparison.OrdinalIgnoreCase) &&
                    !c.Equals("RESTRICTED", StringComparison.OrdinalIgnoreCase))
                .ToList();

            categories.Add(result.SuggestedTag.ToUpperInvariant());

            mail.Categories = string.Join(", ", categories);
        }

        private static string RemoveSecurityTags(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
                return string.Empty;

            while (true)
            {
                string updated = Regex.Replace(
                    subject,
                    @"^\s*\[(ENCRYPTED|PUBLIC|INTERNAL|CONFIDENTIAL|RESTRICTED)\]\s*",
                    "",
                    RegexOptions.IgnoreCase);

                if (updated == subject)
                    break;

                subject = updated;
            }

            return subject.Trim();
        }

        private void ApplySubjectClassification(Outlook.MailItem mail, MailAnalysisResult result)
        {
            if (mail == null)
                return;

            // Determine if the mail was already encrypted
            bool encrypted = IsMailEncrypted(mail);

            // Remove every existing security tag
            string cleanSubject = RemoveSecurityTags(mail.Subject);

            // Determine classification
            string classification = string.IsNullOrWhiteSpace(result.SuggestedTag)
                ? "PUBLIC"
                : result.SuggestedTag.ToUpperInvariant();

            // Build subject from scratch
            if (encrypted)
            {
                mail.Subject = $"[ENCRYPTED] [{classification}] {cleanSubject}";
            }
            else
            {
                mail.Subject = $"[{classification}] {cleanSubject}";
            }
        }

        private int GetClearanceLevel(string clearance)
        {
            switch ((clearance ?? "").ToUpperInvariant())
            {
                case "PUBLIC":
                    return 0;

                case "INTERNAL":
                    return 1;

                case "CONFIDENTIAL":
                    return 2;

                case "RESTRICTED":
                    return 3;

                default:
                    return -1;
            }
        }

        private bool ValidateRecipients(
    Outlook.MailItem mail,
    MailAnalysisResult result)
        {
            int mailLevel = GetClearanceLevel(result.SuggestedTag);

            // PUBLIC mail can go to anyone
            if (mailLevel <= 0)
                return true;

            foreach (Outlook.Recipient recipient in mail.Recipients)
            {
                string email = recipient.Address;

                string clearance;

                try
                {
                    clearance = _apiService
                        .GetRecipientClearanceAsync(email)
                        .GetAwaiter()
                        .GetResult();
                }
                catch
                {
                    MessageBox.Show(
                        $"Unable to determine clearance for:\n\n{email}",
                        "Secure Mail",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return false;
                }

                int recipientLevel = GetClearanceLevel(clearance);

                if (recipientLevel < mailLevel)
                {
                    MessageBox.Show(
                        $"Recipient '{email}' has {clearance} clearance.\n\n" +
                        $"This email is classified as {result.SuggestedTag}.\n\n" +
                        $"Email cannot be sent.",
                        "Recipient Clearance Violation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return false;
                }
            }

            return true;
        }
    }
}