using System;
using System.Collections.Generic;
using System.Linq;

namespace SecureMailAddin.ZeroTrust
{
    /// <summary>
    /// Deterministic, default-deny authorization engine.
    ///
    /// Authentication is NOT represented by a client-supplied Boolean. A request
    /// can only be authorized when it carries a TrustedAuthenticationProof issued
    /// by JwtTokenValidator after signature/issuer/audience/lifetime validation.
    /// The proof binds the subject, roles, clearance and other identity fields.
    /// </summary>
    public class AccessPolicyEngine
    {
        // PUBLIC < INTERNAL < CONFIDENTIAL < RESTRICTED
        private static readonly Dictionary<string, int> ClearanceRanks =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "PUBLIC", 0 },
                { "INTERNAL", 1 },
                { "CONFIDENTIAL", 2 },
                { "RESTRICTED", 3 }
            };

        private static readonly HashSet<string> ValidRoles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "guest", "employee", "admin"
            };

        private sealed class ActionPolicy
        {
            public string RequiredRole { get; private set; }
            public int MinimumClearanceRank { get; private set; }

            public ActionPolicy(string requiredRole, int minimumClearanceRank)
            {
                RequiredRole = requiredRole;
                MinimumClearanceRank = minimumClearanceRank;
            }
        }

        // Explicit allow-list. Anything absent from this table is BLOCK.
        private static readonly Dictionary<ProtectedActionType, ActionPolicy> Policies =
            new Dictionary<ProtectedActionType, ActionPolicy>
            {
                { ProtectedActionType.OpenAdminDashboard, new ActionPolicy("admin", 0) },
                { ProtectedActionType.ApprovePendingUser, new ActionPolicy("admin", 0) },
                { ProtectedActionType.RejectPendingUser, new ActionPolicy("admin", 0) },
                { ProtectedActionType.ViewSensitiveLogs, new ActionPolicy("admin", 0) },
                { ProtectedActionType.ExportAuditData, new ActionPolicy("admin", 0) },

                { ProtectedActionType.SendPublicMail, new ActionPolicy(null, 0) },
                { ProtectedActionType.SendInternalMail, new ActionPolicy(null, 1) },
                { ProtectedActionType.SendConfidentialMail, new ActionPolicy(null, 2) },
                { ProtectedActionType.SendRestrictedMail, new ActionPolicy(null, 3) }
            };

        /// <summary>
        /// Existing production signature. The sensitivityLevel/classification tag
        /// is recorded in the decision but does not independently change this
        /// legacy evaluation path. Use EvaluateMailSend for the Outlook workflow.
        /// </summary>
        public AccessDecision Evaluate(
            UserClaimsProfile user,
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            string actionName;

            try
            {
                actionName = action.ToString();
            }
            catch
            {
                actionName = string.Empty;
            }

            try
            {
                if (user == null)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.InvalidRequest,
                        "Authorization request is null.",
                        actionName,
                        null,
                        null,
                        sensitivityLevel);
                }

                // Step 1: cryptographic authentication proof.
                if (user.AuthenticationProof == null)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.AuthenticationRequired,
                        "No trusted JWT validation proof is attached to the identity.",
                        actionName,
                        user,
                        null,
                        sensitivityLevel);
                }

                if (!user.AuthenticationProof.Matches(user))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.AuthenticationClaimsAltered,
                        "Trusted identity claims were altered after JWT validation.",
                        actionName,
                        user,
                        null,
                        sensitivityLevel);
                }

                if (string.IsNullOrWhiteSpace(user.Subject))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.SubjectMissing,
                        "Authenticated subject is missing.",
                        actionName,
                        user,
                        null,
                        sensitivityLevel);
                }

                // Step 2: trusted identity attributes.
                if (user.Roles == null || user.Roles.Count == 0)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.RoleMissing,
                        "Authenticated role is missing.",
                        actionName,
                        user,
                        null,
                        sensitivityLevel);
                }

                string effectiveRole = GetEffectiveRole(user.Roles);
                if (effectiveRole == null)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.RoleUnknown,
                        "No supported authenticated role was found.",
                        actionName,
                        user,
                        null,
                        sensitivityLevel);
                }

                int userClearanceRank;
                if (string.IsNullOrWhiteSpace(user.ClearanceLevel))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.ClearanceMissing,
                        "Authenticated clearance is missing.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                if (!TryGetClearanceRank(user.ClearanceLevel, out userClearanceRank))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.ClearanceUnknown,
                        "Authenticated clearance is not recognized.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                // Step 3: default-deny action resolution.
                if (!Enum.IsDefined(typeof(ProtectedActionType), action))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.ActionUnknown,
                        "Requested action has no explicitly defined policy.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                ActionPolicy policy;
                if (!Policies.TryGetValue(action, out policy))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.ActionUnknown,
                        "Requested action has no explicitly defined policy.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                // Step 4: role requirement.
                if (policy.RequiredRole != null &&
                    !string.Equals(effectiveRole, policy.RequiredRole, StringComparison.OrdinalIgnoreCase))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.RoleInsufficient,
                        "Administrative role is required for this action.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                // Step 5: clearance hierarchy. Higher may satisfy lower; lower
                // never satisfies higher.
                if (userClearanceRank < policy.MinimumClearanceRank)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.ClearanceInsufficient,
                        "User clearance is below the required classification.",
                        actionName,
                        user,
                        effectiveRole,
                        sensitivityLevel);
                }

                return AccessDecision.Allow(
                    PolicyCodes.PolicySatisfied,
                    "Authentication, role, clearance and action requirements are satisfied.",
                    user,
                    actionName,
                    effectiveRole,
                    sensitivityLevel);
            }
            catch (Exception)
            {
                // No exception is allowed to convert a BLOCK into ALLOW.
                return AccessDecision.Deny(
                    PolicyCodes.PolicyEngineError,
                    "An unexpected authorization error occurred; failing closed.",
                    actionName,
                    user,
                    null,
                    sensitivityLevel);
            }
        }

        /// <summary>
        /// String overload for datasets and external tests. Action names are
        /// exact; malformed names, numeric strings and case variations are denied.
        /// </summary>
        public AccessDecision Evaluate(
            UserClaimsProfile user,
            string actionName,
            string sensitivityLevel = null)
        {
            ProtectedActionType action;

            if (string.IsNullOrWhiteSpace(actionName))
            {
                return AccessDecision.Deny(
                    PolicyCodes.ActionMissing,
                    "Requested action is missing.",
                    string.Empty,
                    user,
                    null,
                    sensitivityLevel);
            }

            if (!TryResolveAction(actionName, out action))
            {
                return AccessDecision.Deny(
                    PolicyCodes.ActionUnknown,
                    "Requested action is not an exact, explicitly defined policy name.",
                    actionName,
                    user,
                    null,
                    sensitivityLevel);
            }

            return Evaluate(user, action, sensitivityLevel);
        }

        /// <summary>
        /// Strict Outlook mail-send path: the independently classified tag is
        /// mapped to an action before authorization, preventing downgrade to a
        /// lower-clearance mail action.
        /// </summary>
        public AccessDecision EvaluateMailSend(
            UserClaimsProfile user,
            string classificationTag)
        {
            ProtectedActionType action;

            if (!TryMapClassificationToAction(classificationTag, out action))
            {
                return AccessDecision.Deny(
                    PolicyCodes.ClassificationUnknown,
                    "Mail classification is missing or unsupported; failing closed.",
                    classificationTag ?? string.Empty,
                    user,
                    null,
                    classificationTag);
            }

            return Evaluate(user, action, classificationTag);
        }

        public static bool TryGetClearanceRank(string clearance, out int rank)
        {
            rank = -1;

            if (string.IsNullOrWhiteSpace(clearance))
                return false;

            // Reject padded/malformed values rather than silently normalizing.
            if (!string.Equals(clearance, clearance.Trim(), StringComparison.Ordinal))
                return false;

            return ClearanceRanks.TryGetValue(clearance, out rank);
        }

        public static bool HasSufficientClearance(string userClearance, string requiredClearance)
        {
            int userRank;
            int requiredRank;

            if (!TryGetClearanceRank(userClearance, out userRank) ||
                !TryGetClearanceRank(requiredClearance, out requiredRank))
            {
                return false;
            }

            return userRank >= requiredRank;
        }

        /// <summary>
        /// Exact, case-sensitive action parsing. This is deliberately stricter
        /// than Enum.TryParse(..., ignoreCase: true).
        /// </summary>
        public static bool TryResolveAction(string actionName, out ProtectedActionType action)
        {
            action = ProtectedActionType.Unknown;

            if (string.IsNullOrWhiteSpace(actionName) ||
                !string.Equals(actionName, actionName.Trim(), StringComparison.Ordinal) ||
                char.IsDigit(actionName[0]))
            {
                return false;
            }

            if (!Enum.TryParse(actionName, false, out action))
                return false;

            return Enum.IsDefined(typeof(ProtectedActionType), action) &&
                   Policies.ContainsKey(action);
        }

        public static bool TryMapClassificationToAction(
            string classificationTag,
            out ProtectedActionType action)
        {
            action = ProtectedActionType.Unknown;

            if (string.IsNullOrWhiteSpace(classificationTag) ||
                !string.Equals(classificationTag, classificationTag.Trim(), StringComparison.Ordinal))
            {
                return false;
            }

            if (string.Equals(classificationTag, "PUBLIC", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendPublicMail;
                return true;
            }

            if (string.Equals(classificationTag, "INTERNAL", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendInternalMail;
                return true;
            }

            if (string.Equals(classificationTag, "CONFIDENTIAL", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendConfidentialMail;
                return true;
            }

            if (string.Equals(classificationTag, "RESTRICTED", StringComparison.OrdinalIgnoreCase))
            {
                action = ProtectedActionType.SendRestrictedMail;
                return true;
            }

            return false;
        }

        public static bool TryGetPolicy(
            ProtectedActionType action,
            out string requiredRole,
            out int minimumClearanceRank)
        {
            ActionPolicy policy;
            if (!Policies.TryGetValue(action, out policy))
            {
                requiredRole = null;
                minimumClearanceRank = -1;
                return false;
            }

            requiredRole = policy.RequiredRole;
            minimumClearanceRank = policy.MinimumClearanceRank;
            return true;
        }

        /// <summary>
        /// Selects the highest supported role. Unknown Keycloak infrastructure
        /// roles are ignored and never grant access; a user with only unknown
        /// roles is denied.
        /// </summary>
        public static string GetEffectiveRole(IEnumerable<string> roles)
        {
            if (roles == null)
                return null;

            List<string> normalized = roles
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Select(r => r.Trim())
                .ToList();

            if (normalized.Any(r => string.Equals(r, "admin", StringComparison.OrdinalIgnoreCase)))
                return "admin";

            if (normalized.Any(r => string.Equals(r, "employee", StringComparison.OrdinalIgnoreCase)))
                return "employee";

            if (normalized.Any(r => string.Equals(r, "guest", StringComparison.OrdinalIgnoreCase)))
                return "guest";

            return null;
        }
    }
}
