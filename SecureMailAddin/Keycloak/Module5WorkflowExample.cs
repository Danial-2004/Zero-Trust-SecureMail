using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecureMailAddin.Keycloak
{
    public class Module5WorkflowExample
    {
        private readonly ApiIntegrationService _apiService;

        public Module5WorkflowExample()
        {
            _apiService = new ApiIntegrationService();
        }

        public async Task HandleBeforeSendAsync(Outlook.MailItem mailItem, ClassificationResult result)
        {
            if (mailItem == null)
                throw new ArgumentNullException(nameof(mailItem));

            if (result == null)
                throw new ArgumentNullException(nameof(result));

            // Optional: fetch latest policy so app stays synced with backend rules
            await _apiService.GetClassificationRulesAsync();

            var classificationRequest = new MailClassificationRequest
            {
                OutlookMessageId = mailItem.EntryID ?? string.Empty,
                Subject = mailItem.Subject ?? string.Empty,
                Recipients = GetRecipients(mailItem),
                Attachments = GetAttachments(mailItem),
                Score = result.Score,
                Classification = result.Label ?? string.Empty,
                EncryptionApplied = result.RequiresEncryption,
                ProcessedAtUtc = DateTime.UtcNow
            };

            await _apiService.SubmitMailClassificationAsync(classificationRequest);

            await _apiService.SendAuditLogAsync(new AuditLogRequest
            {
                EventType = "MAIL_CLASSIFIED",
                Username = Environment.UserName,
                EmailSubject = mailItem.Subject ?? string.Empty,
                Classification = result.Label ?? string.Empty,
                TimestampUtc = DateTime.UtcNow,
                Details = "Mail classified successfully and reported to central server."
            });
        }

        private List<string> GetRecipients(Outlook.MailItem mailItem)
        {
            var recipients = new List<string>();

            if (mailItem.Recipients != null)
            {
                foreach (Outlook.Recipient recipient in mailItem.Recipients)
                {
                    if (recipient != null && !string.IsNullOrWhiteSpace(recipient.Address))
                        recipients.Add(recipient.Address);
                }
            }

            return recipients;
        }

        private List<string> GetAttachments(Outlook.MailItem mailItem)
        {
            var attachments = new List<string>();

            if (mailItem.Attachments != null)
            {
                foreach (Outlook.Attachment attachment in mailItem.Attachments)
                {
                    if (attachment != null && !string.IsNullOrWhiteSpace(attachment.FileName))
                        attachments.Add(attachment.FileName);
                }
            }

            return attachments;
        }
    }
}