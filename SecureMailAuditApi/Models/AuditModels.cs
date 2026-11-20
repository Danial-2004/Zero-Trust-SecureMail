using System;
using System.Collections.Generic;

namespace SecureMailAuditApi.Models
{
    public class AuditLogRequest
    {
        public string EventType { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string EmailSubject { get; set; } = string.Empty;
        public string Classification { get; set; } = string.Empty;
        public DateTime TimestampUtc { get; set; }
        public string Details { get; set; } = string.Empty;
    }

    public class MailClassificationRequest
    {
        public string OutlookMessageId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public List<string> Recipients { get; set; } = new List<string>();
        public List<string> Attachments { get; set; } = new List<string>();
        public int Score { get; set; }
        public string Classification { get; set; } = string.Empty;
        public bool EncryptionApplied { get; set; }
        public DateTime ProcessedAtUtc { get; set; }
    }
}