using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace SecureMailAddin.Keycloak
{
    public class KeycloakUserInfo
    {
        [JsonProperty("sub")]
        public string sub { get; set; }

        [JsonProperty("preferred_username")]
        public string preferred_username { get; set; }

        [JsonProperty("email")]
        public string email { get; set; }

        [JsonProperty("name")]
        public string name { get; set; }

        [JsonProperty("department")]
        public string department { get; set; }

        [JsonProperty("clearance")]
        public string clearance { get; set; }

        [JsonProperty("realm_access")]
        public RealmAccess realm_access { get; set; }

        [JsonProperty("resource_access")]
        public Dictionary<string, ClientAccess> resource_access { get; set; }
    }

    public class RealmAccess
    {
        [JsonProperty("roles")]
        public List<string> roles { get; set; }
    }

    public class ClientAccess
    {
        [JsonProperty("roles")]
        public List<string> roles { get; set; }
    }

    public class KeycloakUser
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("username")]
        public string Username { get; set; }

        [JsonProperty("email")]
        public string Email { get; set; }

        [JsonProperty("firstName")]
        public string FirstName { get; set; }

        [JsonProperty("lastName")]
        public string LastName { get; set; }

        [JsonProperty("enabled")]
        public bool Enabled { get; set; }

        [JsonProperty("attributes")]
        public Dictionary<string, List<string>> Attributes { get; set; }
    }

    public class AuthResult
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonProperty("refresh_expires_in")]
        public int RefreshExpiresIn { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("scope")]
        public string Scope { get; set; }

        [JsonIgnore]
        public DateTime ExpiresAt { get; set; }
    }

    public class StoredTokenInfo
    {
        public string AccessToken { get; set; }
        public string RefreshToken { get; set; }
        public DateTime ExpiresAt { get; set; }
    }

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

    public class AuditLogRequest
    {
        public string EventType { get; set; }
        public string Username { get; set; }
        public string EmailSubject { get; set; }
        public string Classification { get; set; }
        public DateTime TimestampUtc { get; set; }
        public string Details { get; set; }
    }

    public class MailClassificationRequest
    {
        public string OutlookMessageId { get; set; }
        public string Subject { get; set; }
        public List<string> Recipients { get; set; }
        public List<string> Attachments { get; set; }
        public int Score { get; set; }
        public string Classification { get; set; }
        public bool EncryptionApplied { get; set; }
        public DateTime ProcessedAtUtc { get; set; }
    }

    public class ClassificationResult
    {
        public int Score { get; set; }
        public string Label { get; set; }
        public bool RequiresEncryption { get; set; }
    }
}