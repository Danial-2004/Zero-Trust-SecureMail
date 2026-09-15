using System.Collections.Generic;

namespace SecureMailPolicyApi.Models
{
    public class PolicyResponse
    {
        public int RestrictedSubjectBodyScore { get; set; }
        public int RestrictedAttachmentNameScore { get; set; }
        public int ConfidentialSubjectBodyScore { get; set; }
        public int ConfidentialAttachmentNameScore { get; set; }
        public int SensitiveAttachmentScore { get; set; }
        public int ExternalRecipientScore { get; set; }
        public int AnyAttachmentScore { get; set; }
        public int SensitiveExternalComboScore { get; set; }

        public List<string> RestrictedKeywords { get; set; }
        public List<string> ConfidentialKeywords { get; set; }
        public List<string> SensitiveAttachmentPatterns { get; set; }
    }
}