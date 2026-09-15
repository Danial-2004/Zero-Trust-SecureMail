using System;
using System.Linq;

namespace SecureMailAddin.ZeroTrust
{
    public class AccessPolicyEngine
    {
        public AccessDecision Evaluate(
            UserClaimsProfile user,
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            if (user == null)
            {
                return new AccessDecision
                {
                    IsAllowed = false,
                    Reason = "User claims are missing."
                };
            }

            switch (action)
            {
                case ProtectedActionType.OpenAdminDashboard:
                case ProtectedActionType.ApprovePendingUser:
                case ProtectedActionType.RejectPendingUser:
                case ProtectedActionType.ViewSensitiveLogs:
                case ProtectedActionType.ExportAuditData:
                    return RequireRole(user, "admin");

                case ProtectedActionType.SendConfidentialMail:
                    return RequireClearance(user, "CONFIDENTIAL");

                case ProtectedActionType.SendRestrictedMail:
                    return RequireClearance(user, "RESTRICTED");

                default:
                    return new AccessDecision
                    {
                        IsAllowed = true,
                        Reason = "No restriction rule matched."
                    };
            }
        }

        private AccessDecision RequireRole(UserClaimsProfile user, string requiredRole)
        {
            bool ok = user.Roles != null &&
                      user.Roles.Any(r =>
                          string.Equals(r, requiredRole, StringComparison.OrdinalIgnoreCase));

            return new AccessDecision
            {
                IsAllowed = ok,
                Reason = ok
                    ? "Role requirement satisfied."
                    : $"Required role '{requiredRole}' is missing."
            };
        }

        private AccessDecision RequireClearance(UserClaimsProfile user, string requiredClearance)
        {
            bool ok = string.Equals(
                user.ClearanceLevel?.Trim(),
                requiredClearance,
                StringComparison.OrdinalIgnoreCase);

            return new AccessDecision
            {
                IsAllowed = ok,
                Reason = ok
                    ? "Clearance requirement satisfied."
                    : $"Required clearance '{requiredClearance}' is missing."
            };
        }
    }
}