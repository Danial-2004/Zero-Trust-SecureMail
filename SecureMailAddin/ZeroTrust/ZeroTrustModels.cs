using System.Collections.Generic;

namespace SecureMailAddin.ZeroTrust
{
    public enum ProtectedActionType
    {
        OpenAdminDashboard,
        ApprovePendingUser,
        RejectPendingUser,
        SendConfidentialMail,
        SendRestrictedMail,
        ViewSensitiveLogs,
        ExportAuditData,
        SendInternalMail,
        SendPublicMail
    }

    public class UserClaimsProfile
    {
        public string Username { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public List<string> Roles { get; set; } = new List<string>();
        public string Department { get; set; } = string.Empty;
        public string ClearanceLevel { get; set; } = string.Empty;
    }

    public class AccessDecision
    {
        public bool IsAllowed { get; set; }
        public string Reason { get; set; } = string.Empty;
    }
}