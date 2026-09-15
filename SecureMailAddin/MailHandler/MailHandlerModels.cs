using System;
using System.Collections.Generic;

namespace SecureMailAddin.MailHandler
{
    public enum MailSensitivityLevel
    {
        Public = 0,
        Internal = 1,
        Confidential = 2,
        Restricted = 3
    }

    public class AttachmentInfo
    {
        public string FileName { get; set; }
        public long Size { get; set; }
        public bool IsEncrypted { get; set; }
        public bool ContentScanned { get; set; }
        public bool ContainsSensitiveData { get; set; }
        public int SensitivityScore { get; set; }
    }

    public class RecipientInfo
    {
        public string DisplayName { get; set; }
        public string Address { get; set; }
        public bool IsInternal { get; set; }
    }

    public class MailAnalysisResult
    {
        public string Subject { get; set; }
        public string BodyPreview { get; set; }

        public List<RecipientInfo> Recipients { get; set; } = new List<RecipientInfo>();
        public List<AttachmentInfo> Attachments { get; set; } = new List<AttachmentInfo>();

        public bool HasAttachments { get; set; }
        public bool HasExternalRecipients { get; set; }
        public bool ContainsSensitiveKeywords { get; set; }
        public bool ContainsSensitiveAttachmentData { get; set; }

        public int RestrictedKeywordHits { get; set; }
        public int ConfidentialKeywordHits { get; set; }
        public int TotalScore { get; set; }

        public MailSensitivityLevel SensitivityLevel { get; set; }
        public string SuggestedTag { get; set; }
        public bool RequiresEncryption { get; set; }
        public bool RequiresUserConfirmation { get; set; }
        public string Reason { get; set; }
    }
}