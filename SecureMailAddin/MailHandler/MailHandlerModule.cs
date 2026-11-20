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

        // Single ZeroTrustHandler for the lifetime of this module so the
        // replay cache is not recreated on every send attempt.
        private readonly ZeroTrustHandler _zeroTrustHandler;

        public MailHandlerModule(string companyDomain)
        {
            _classificationEngine = new MailClassificationEngine(companyDomain);
            _apiService = new ApiIntegrationService();
            _zeroTrustHandler = new ZeroTrustHandler();
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
                {
                    Logger.LogWarning("MAIL_HANDLER", "Mail item is null; send blocked.");
                    return false;
                }

                Logger.LogMailBeforeSendStarted(mail.Subject);

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

                if (result == null)
                {
                    Logger.LogWarning("MAIL_HANDLER", "Classification returned no result; send blocked.");
                    return false;
                }

                ProtectedActionType mappedAction;
                if (!AccessPolicyEngine.TryMapClassificationToAction(result.SuggestedTag, out mappedAction))
                {
                    Logger.LogWarning("MAIL_HANDLER", "Unsupported or missing mail classification; send blocked.");
                    MessageBox.Show(
                        "The mail classification is missing or unsupported. The email cannot be sent.",
                        "Secure Mail Policy",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                ApplyClassification(mail, result);

                Logger.LogMailClassificationResult(
                    mail.Subject,
                    result.SuggestedTag,
                    result.SensitivityLevel.ToString(),
                    result.TotalScore);

                Logger.LogMailZeroTrustCheckStarted(mail?.Subject, result.SuggestedTag);

                // Generate a unique request ID bound to the mail content so
                // that resubmitting the same email is detected as a replay.
                string requestId = GenerateRequestId(mail, result.SuggestedTag);

                bool allowed = EnforceZeroTrustBeforeSend(mail?.Subject, result.SuggestedTag, requestId);
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
                catch (Exception ex)
                {
                    Logger.LogError("MAIL_HANDLER", "Classification/audit workflow failed; send blocked.", ex);

                    MessageBox.Show(
                        "The required classification or audit event could not be recorded. The email cannot be sent.",
                        "Secure Mail Policy",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
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

        public bool EnforceZeroTrustBeforeSend(string subject, string suggestedTag, string requestId = null)
        {
            ProtectedActionType mappedAction;
            if (!AccessPolicyEngine.TryMapClassificationToAction(suggestedTag, out mappedAction))
            {
                Logger.LogWarning("MAIL_HANDLER", "Mail classification is unknown; send blocked.");
                return false;
            }

            try
            {
                // The strict workflow maps the classification tag to the mail
                // action, so PUBLIC/INTERNAL no longer bypass policy and a
                // downgraded action cannot be requested directly.
                // Reuses the module-level handler so the replay cache is stable.
                AccessDecision decision = _zeroTrustHandler
                    .AuthorizeMailSendAsync(suggestedTag, requestId)
                    .GetAwaiter()
                    .GetResult();

                if (decision == null || !decision.IsAllowed)
                {
                    string reason = decision == null
                        ? "Authorization produced no decision; failing closed."
                        : decision.Reason + " PolicyCode=" + decision.PolicyCode;

                    Logger.LogMailBlockedByZeroTrust(subject, suggestedTag, reason);

                    System.Windows.Forms.MessageBox.Show(
                        "Zero Trust policy denied this action.\n\n" + reason,
                        "Access Denied",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    return false;
                }

                Logger.LogMailAllowedByZeroTrust(subject, suggestedTag);
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError("MAIL_HANDLER", "Zero Trust authorization failed; send blocked.", ex);
                return false;
            }
        }

        private bool IsMailEncrypted(Outlook.MailItem mail)
        {
            if (mail == null)
                return false;

            string subject = mail.Subject ?? string.Empty;
            string body = mail.Body ?? string.Empty;

            bool subjectMarker = Regex.IsMatch(
                subject,
                @"^\s*\[ENCRYPTED\]\s*",
                RegexOptions.IgnoreCase);

            bool secureBlock =
                body.IndexOf("---BEGIN SECURE EMAIL---", StringComparison.OrdinalIgnoreCase) >= 0 &&
                body.IndexOf("---END SECURE EMAIL---", StringComparison.OrdinalIgnoreCase) >= 0 &&
                body.IndexOf("RECIPIENT_KEYS", StringComparison.OrdinalIgnoreCase) >= 0 &&
                Regex.IsMatch(body, @"(?m)^\s*HMAC:\s*[A-Za-z0-9+/=]+\s*$", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(body, @"(?m)^\s*IV:\s*[A-Za-z0-9+/=]+\s*$", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(body, @"(?m)^\s*DATA:\s*[A-Za-z0-9+/=]+\s*$", RegexOptions.IgnoreCase);

            return subjectMarker && secureBlock && AllAttachmentsEncrypted(mail);
        }

        private static bool AllAttachmentsEncrypted(Outlook.MailItem mail)
        {
            try
            {
                if (mail.Attachments == null || mail.Attachments.Count == 0)
                    return true;

                for (int i = mail.Attachments.Count; i >= 1; i--)
                {
                    Outlook.Attachment attachment = mail.Attachments[i];
                    string name = attachment == null ? null : attachment.FileName;

                    if (string.IsNullOrWhiteSpace(name) ||
                        !name.EndsWith(".enc", StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
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

            if (result == null || string.IsNullOrWhiteSpace(result.SuggestedTag))
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

            categories.Add(result.SuggestedTag.Trim().ToUpperInvariant());

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
            int rank;
            return AccessPolicyEngine.TryGetClearanceRank(clearance, out rank)
                ? rank
                : -1;
        }

        private bool ValidateRecipients(
            Outlook.MailItem mail,
            MailAnalysisResult result)
        {
            if (mail == null || result == null)
                return false;

            int mailLevel = GetClearanceLevel(result.SuggestedTag);

            if (mailLevel < 0)
            {
                Logger.LogWarning("MAIL_HANDLER", "Unknown mail classification; recipient validation failed closed.");
                MessageBox.Show(
                    "The mail classification is unsupported. The email cannot be sent.",
                    "Secure Mail Policy",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            // PUBLIC mail can go to any recipient.
            if (mailLevel == 0)
                return true;

            try
            {
                if (mail.Recipients == null || mail.Recipients.Count == 0)
                {
                    MessageBox.Show(
                        "No recipients were found. The email cannot be sent.",
                        "Secure Mail",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                foreach (Outlook.Recipient recipient in mail.Recipients)
                {
                    if (recipient == null)
                        return false;

                    string email = recipient.Address;
                    if (string.IsNullOrWhiteSpace(email))
                    {
                        Logger.LogWarning("MAIL_HANDLER", "Recipient address is missing; send blocked.");
                        return false;
                    }

                    string clearance;
                    try
                    {
                        clearance = _apiService
                            .GetRecipientClearanceAsync(email)
                            .GetAwaiter()
                            .GetResult();
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError("MAIL_HANDLER", "Recipient clearance lookup failed; send blocked.", ex);
                        MessageBox.Show(
                            $"Unable to determine clearance for:\n\n{email}",
                            "Secure Mail",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return false;
                    }

                    int recipientLevel = GetClearanceLevel(clearance);
                    if (recipientLevel < 0)
                    {
                        Logger.LogWarning("MAIL_HANDLER", "Recipient clearance is unsupported; send blocked.");
                        MessageBox.Show(
                            $"Recipient '{email}' has an unsupported clearance. The email cannot be sent.",
                            "Recipient Clearance Violation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return false;
                    }

                    if (recipientLevel < mailLevel)
                    {
                        MessageBox.Show(
                            $"Recipient '{email}' has {clearance} clearance.\n\n" +
                            $"This email is classified as {result.SuggestedTag}.\n\n" +
                            "Email cannot be sent.",
                            "Recipient Clearance Violation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError("MAIL_HANDLER", "Recipient validation failed; send blocked.", ex);
                return false;
            }
        }

        /// <summary>
        /// Generates a deterministic request ID bound to the mail content and
        /// classification. Resubmitting the same email produces the same ID,
        /// enabling request-level replay detection without blocking legitimate
        /// JWT token reuse.
        /// </summary>
        private static string GenerateRequestId(Outlook.MailItem mail, string classificationTag)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var parts = new List<string>();

                // Subject
                parts.Add(mail?.Subject ?? string.Empty);

                // Recipients (sorted for determinism)
                var recipients = new List<string>();
                try
                {
                    if (mail?.Recipients != null)
                    {
                        foreach (Outlook.Recipient recipient in mail.Recipients)
                        {
                            if (recipient != null)
                            {
                                string addr = recipient.Address ?? recipient.Name ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(addr))
                                    recipients.Add(addr.Trim().ToLowerInvariant());
                            }
                        }
                    }
                }
                catch
                {
                    // Outlook may not be fully available; fall back to subject-only.
                }
                recipients.Sort();
                parts.Add(string.Join("|", recipients));

                // Body hash (first 4 KB, enough to distinguish content)
                string body = mail?.Body ?? string.Empty;
                parts.Add(body.Length > 4096 ? body.Substring(0, 4096) : body);

                // Classification tag
                parts.Add(classificationTag ?? string.Empty);

                string canonical = string.Join("\u001E", parts);
                byte[] hash = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(canonical));
                return Convert.ToBase64String(hash);
            }
        }
    }
}
