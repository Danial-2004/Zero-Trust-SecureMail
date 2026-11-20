using System;
using System.Collections.Generic;

namespace SecureMailAddin.ZeroTrust
{
    /// <summary>
    /// Explicitly enumerated actions. Unknown values, including integer casts,
    /// are rejected by AccessPolicyEngine (default deny).
    /// </summary>
    public enum ProtectedActionType
    {
        OpenAdminDashboard = 0,
        ApprovePendingUser = 1,
        RejectPendingUser = 2,
        SendConfidentialMail = 3,
        SendRestrictedMail = 4,
        ViewSensitiveLogs = 5,
        ExportAuditData = 6,
        SendInternalMail = 7,
        SendPublicMail = 8,
        Unknown = 999
    }

    public enum PolicyDecision
    {
        BLOCK = 0,
        ALLOW = 1
    }

    public enum JwtValidationStatus
    {
        Missing = 0,
        Malformed = 1,
        Invalid = 2,
        Expired = 3,
        NotYetValid = 4,
        InvalidSignature = 5,
        InvalidIssuer = 6,
        InvalidAudience = 7,
        MissingClaims = 8,
        Replayed = 9,
        ValidationUnavailable = 10,
        Valid = 11
    }

    /// <summary>
    /// Deterministic policy codes. These are codes, not confidence scores.
    /// </summary>
    public static class PolicyCodes
    {
        public const string PolicySatisfied = "POLICY_SATISFIED";
        public const string InvalidRequest = "INVALID_REQUEST";
        public const string AuthenticationRequired = "AUTHENTICATION_REQUIRED";
        public const string AuthenticationClaimsAltered = "AUTHENTICATION_CLAIMS_ALTERED";
        public const string JwtMissing = "JWT_MISSING";
        public const string JwtMalformed = "JWT_MALFORMED";
        public const string JwtInvalid = "JWT_INVALID";
        public const string JwtExpired = "JWT_EXPIRED";
        public const string JwtNotYetValid = "JWT_NOT_YET_VALID";
        public const string JwtSignatureInvalid = "JWT_SIGNATURE_INVALID";
        public const string JwtIssuerInvalid = "JWT_ISSUER_INVALID";
        public const string JwtAudienceInvalid = "JWT_AUDIENCE_INVALID";
        public const string JwtLifetimeInvalid = "JWT_LIFETIME_INVALID";
        public const string JwtAlgorithmNotAllowed = "JWT_ALGORITHM_NOT_ALLOWED";
        public const string JwtReplayDetected = "JWT_REPLAY_DETECTED";
        public const string JwtTokenIdMissing = "JWT_TOKEN_ID_MISSING";
        public const string JwtRequiredClaimMissing = "JWT_REQUIRED_CLAIM_MISSING";
        public const string SubjectMissing = "SUBJECT_MISSING";
        public const string RoleMissing = "ROLE_MISSING";
        public const string RoleUnknown = "ROLE_UNKNOWN";
        public const string RoleInsufficient = "ROLE_INSUFFICIENT";
        public const string ClearanceMissing = "CLEARANCE_MISSING";
        public const string ClearanceUnknown = "CLEARANCE_UNKNOWN";
        public const string ClearanceInsufficient = "CLEARANCE_INSUFFICIENT";
        public const string ActionMissing = "ACTION_MISSING";
        public const string ActionUnknown = "ACTION_UNKNOWN";
        public const string ClassificationUnknown = "CLASSIFICATION_UNKNOWN";
        public const string AuthenticationServiceUnavailable = "AUTH_SERVICE_UNAVAILABLE";
        public const string AuthenticationConfigurationMissing = "AUTH_CONFIGURATION_MISSING";
        public const string AuditServiceUnavailable = "AUDIT_SERVICE_UNAVAILABLE";
        public const string PolicyEngineError = "POLICY_ENGINE_ERROR";
        public const string NotEvaluated = "NOT_EVALUATED";
    }

    /// <summary>
    /// Claims used for authorization. The AuthenticationProof property can only
    /// be set by the token-validation layer inside this assembly. A plain
    /// UserClaimsProfile supplied by a client cannot authorize an ALLOW.
    /// </summary>
    public class UserClaimsProfile
    {
        public string Subject { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public List<string> Roles { get; set; } = new List<string>();
        public string Department { get; set; } = string.Empty;
        public string ClearanceLevel { get; set; } = string.Empty;

        /// <summary>
        /// Issued only after cryptographic JWT validation. The proof binds the
        /// identity fields so roles/clearance cannot be changed after validation.
        /// </summary>
        public TrustedAuthenticationProof AuthenticationProof { get; internal set; }
    }

    /// <summary>
    /// Backward-compatible decision object. IsAllowed remains writable so the
    /// existing VSTO call sites continue to compile.
    /// </summary>
    public class AccessDecision
    {
        public PolicyDecision Decision { get; set; } = PolicyDecision.BLOCK;

        public bool IsAllowed
        {
            get { return Decision == PolicyDecision.ALLOW; }
            set { Decision = value ? PolicyDecision.ALLOW : PolicyDecision.BLOCK; }
        }

        public string Reason { get; set; } = string.Empty;
        public string PolicyCode { get; set; } = PolicyCodes.NotEvaluated;
        public string Action { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string EffectiveRole { get; set; } = string.Empty;
        public string Clearance { get; set; } = string.Empty;
        public string RequestedClassification { get; set; } = string.Empty;
        public DateTime EvaluatedAtUtc { get; set; } = DateTime.UtcNow;

        public static AccessDecision Allow(
            string policyCode,
            string reason,
            UserClaimsProfile user,
            string action,
            string effectiveRole,
            string requestedClassification)
        {
            var result = new AccessDecision
            {
                Decision = PolicyDecision.ALLOW,
                PolicyCode = policyCode,
                Reason = reason,
                Action = action ?? string.Empty,
                RequestedClassification = requestedClassification ?? string.Empty,
                EffectiveRole = effectiveRole ?? string.Empty,
                EvaluatedAtUtc = DateTime.UtcNow
            };

            if (user != null)
            {
                result.Subject = user.Subject ?? string.Empty;
                result.Username = user.Username ?? string.Empty;
                result.Clearance = user.ClearanceLevel ?? string.Empty;
            }

            return result;
        }

        public static AccessDecision Deny(
            string policyCode,
            string reason,
            string action = null,
            UserClaimsProfile user = null,
            string effectiveRole = null,
            string requestedClassification = null)
        {
            var result = new AccessDecision
            {
                Decision = PolicyDecision.BLOCK,
                PolicyCode = string.IsNullOrWhiteSpace(policyCode)
                    ? PolicyCodes.InvalidRequest
                    : policyCode,
                Reason = reason ?? string.Empty,
                Action = action ?? string.Empty,
                RequestedClassification = requestedClassification ?? string.Empty,
                EffectiveRole = effectiveRole ?? string.Empty,
                EvaluatedAtUtc = DateTime.UtcNow
            };

            if (user != null)
            {
                result.Subject = user.Subject ?? string.Empty;
                result.Username = user.Username ?? string.Empty;
                result.Clearance = user.ClearanceLevel ?? string.Empty;
            }

            return result;
        }

        public AccessDecision WithAction(string action)
        {
            Action = action ?? string.Empty;
            return this;
        }
    }
}
