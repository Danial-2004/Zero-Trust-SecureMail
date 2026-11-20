using System;
using System.Linq;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecureMailAddin.MailHandler
{
    public class MailClassificationEngine
    {
        private readonly string _companyDomain;

        private readonly string[] _restrictedKeywords =
        {
            "exam paper", "question paper", "final paper", "midterm paper",
            "answer key", "solution key", "exam solution", "paper leak",
            "cnic", "identity card", "passport", "license number",
            "national id", "student id copy",
            "password", "passcode", "otp", "verification code",
            "private key", "secret key", "api key", "access token",
            "authentication token", "login credentials", "admin access",
            "salary", "payroll", "bank account", "account number",
            "iban", "credit card", "debit card", "cvv", "billing details",
            "financial record", "transaction details", "invoice confidential",
            "confidential", "strictly confidential", "highly confidential",
            "restricted access", "classified", "do not share",
            "result sheet", "grade sheet", "final result",
            "marks distribution confidential",
            "agreement confidential", "nda", "non disclosure",
            "legal notice", "contract draft confidential",
            "server credentials", "database dump", "system access",
            "root access", "ssh key", "encryption key",
            "secret", "top secret", "internal leak", "sensitive data"
        };

        private readonly string[] _confidentialKeywords =
        {
            "internal", "internal use", "internal document",
            "for internal circulation", "official use", "office use only",
            "student record", "student data", "transcript",
            "marks", "grading", "evaluation", "assignment review",
            "draft", "proposal draft", "initial version",
            "under review", "pending approval",
            "meeting notes", "minutes of meeting", "discussion points",
            "project details", "project plan", "timeline draft",
            "employee record", "staff data", "attendance",
            "performance review", "appraisal",
            "confidential discussion", "private discussion",
            "restricted sharing", "sensitive",
            "departmental", "faculty use", "admin only",
            "committee review",
            "source code", "debug log", "system design",
            "architecture draft", "internal review", "review",
            "discussion","notes",
        };

        public MailClassificationEngine(string companyDomain)
        {
            _companyDomain = (companyDomain ?? string.Empty).Trim().ToLowerInvariant();
        }

        public MailAnalysisResult Analyze(Outlook.MailItem mail)
        {
            var result = new MailAnalysisResult();

            if (mail == null)
                return result;

            result.Subject = mail.Subject ?? string.Empty;
            result.BodyPreview = GetSafePreview(mail.Body);

            ExtractRecipients(mail, result);
            ExtractAttachments(mail, result);

            string fullText = (mail.Subject ?? "") + " " + (mail.Body ?? "");

            int restrictedHitsInMail = CountMatches(fullText, _restrictedKeywords);
            int confidentialHitsInMail = CountMatches(fullText, _confidentialKeywords);

            int sensitiveAttachmentCount = result.Attachments.Count(a => a.ContainsSensitiveData);
            int attachmentScore = result.Attachments.Sum(a => a.SensitivityScore);

            result.RestrictedKeywordHits = restrictedHitsInMail;
            result.ConfidentialKeywordHits = confidentialHitsInMail;

            result.ContainsSensitiveKeywords = restrictedHitsInMail > 0 || confidentialHitsInMail > 0;
            result.ContainsSensitiveAttachmentData = result.Attachments.Any(a => a.ContainsSensitiveData);
            result.HasExternalRecipients = result.Recipients.Any(r => !r.IsInternal);
            result.HasAttachments = result.Attachments.Any();

            int score = 0;

            // Subject/body scoring
            score += restrictedHitsInMail * 10;
            score += confidentialHitsInMail * 4;

            // Attachment filename-only scoring
            score += attachmentScore;
            score += sensitiveAttachmentCount * 4;

            // Routing / context scoring
            if (result.HasExternalRecipients)
                score += 8;

            if (result.HasAttachments)
                score += 2;

            if ((restrictedHitsInMail > 0 || result.ContainsSensitiveAttachmentData) && result.HasExternalRecipients)
                score += 8;

            result.TotalScore = score;

            if (score >= 20)
            {
                result.SensitivityLevel = MailSensitivityLevel.Restricted;
                result.SuggestedTag = "RESTRICTED";
                result.Reason = "High sensitivity score based on subject, body, attachment filenames, and recipients.";
            }
            else if (score >= 8)
            {
                result.SensitivityLevel = MailSensitivityLevel.Confidential;
                result.SuggestedTag = "CONFIDENTIAL";
                result.Reason = "Moderate sensitivity score based on subject, body, or attachment filenames.";
            }
            else if (result.Recipients.Any() && result.Recipients.All(r => r.IsInternal))
            {
                result.SensitivityLevel = MailSensitivityLevel.Internal;
                result.SuggestedTag = "INTERNAL";
                result.Reason = "No strong sensitive indicators found; communication is internal.";
            }
            else
            {
                result.SensitivityLevel = MailSensitivityLevel.Public;
                result.SuggestedTag = "PUBLIC";
                result.Reason = "No significant sensitivity indicators found.";
            }

            result.RequiresEncryption = false;
            result.RequiresUserConfirmation = false;

            return result;
        }

        private void ExtractRecipients(Outlook.MailItem mail, MailAnalysisResult result)
        {
            try
            {
                foreach (Outlook.Recipient recipient in mail.Recipients)
                {
                    string address = ResolveSmtpAddress(recipient);

                    result.Recipients.Add(new RecipientInfo
                    {
                        DisplayName = recipient.Name ?? "",
                        Address = address,
                        IsInternal = IsInternalAddress(address)
                    });
                }
            }
            catch
            {
            }
        }

        private void ExtractAttachments(Outlook.MailItem mail, MailAnalysisResult result)
        {
            try
            {
                foreach (Outlook.Attachment attachment in mail.Attachments)
                {
                    var info = new AttachmentInfo
                    {
                        FileName = attachment.FileName ?? "",
                        Size = attachment.Size,
                        IsEncrypted = (attachment.FileName ?? "").EndsWith(".enc", StringComparison.OrdinalIgnoreCase),
                        ContentScanned = false,
                        ContainsSensitiveData = false,
                        SensitivityScore = 0
                    };

                    try
                    {
                        string fileNameText = info.FileName ?? string.Empty;

                        int restrictedHitsInName = CountMatches(fileNameText, _restrictedKeywords);
                        int confidentialHitsInName = CountMatches(fileNameText, _confidentialKeywords);

                        if (restrictedHitsInName > 0 || confidentialHitsInName > 0)
                        {
                            info.ContainsSensitiveData = true;
                            info.SensitivityScore += restrictedHitsInName * 8;
                            info.SensitivityScore += confidentialHitsInName * 3;
                        }
                    }
                    catch
                    {
                    }

                    result.Attachments.Add(info);
                }
            }
            catch
            {
            }
        }

        private int CountMatches(string text, string[] keywords)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            int count = 0;

            foreach (string keyword in keywords)
            {
                if (string.IsNullOrWhiteSpace(keyword))
                    continue;

                string pattern;

                if (keyword.Contains(" "))
                {
                    string escapedPhrase = Regex.Escape(keyword).Replace("\\ ", "\\s+");
                    pattern = $@"(?<!\w){escapedPhrase}(?!\w)";
                }
                else
                {
                    pattern = $@"\b{Regex.Escape(keyword)}\b";
                }

                MatchCollection matches = Regex.Matches(
                    text,
                    pattern,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                if (matches.Count > 0)
                    count += matches.Count;
            }

            return count;
        }

        private bool IsInternalAddress(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            int atIndex = email.LastIndexOf('@');
            if (atIndex < 0 || atIndex == email.Length - 1)
                return false;

            string domain = email.Substring(atIndex + 1).Trim().ToLowerInvariant();
            return domain == _companyDomain;
        }

        private string ResolveSmtpAddress(Outlook.Recipient recipient)
        {
            try
            {
                if (recipient == null)
                    return "";

                Outlook.AddressEntry entry = recipient.AddressEntry;
                if (entry == null)
                    return recipient.Address ?? "";

                if (entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry ||
                    entry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    Outlook.ExchangeUser exchUser = entry.GetExchangeUser();
                    if (exchUser != null && !string.IsNullOrWhiteSpace(exchUser.PrimarySmtpAddress))
                        return exchUser.PrimarySmtpAddress;
                }

                return recipient.Address ?? "";
            }
            catch
            {
                return recipient.Address ?? "";
            }
        }

        private string GetSafePreview(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
                return "";

            body = body.Replace("\r", " ").Replace("\n", " ").Trim();
            return body.Length <= 200 ? body : body.Substring(0, 200);
        }
    }
}