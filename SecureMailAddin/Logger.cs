using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace SecureMailAddin
{
    public static class Logger
    {
        private static string logFilePath;
        private static readonly object _lock = new object();

        /// <summary>
        /// Initializes the logger and creates log directory/file.
        /// Call this once at startup (e.g., ThisAddIn_Startup).
        /// </summary>
        public static void Initialize()
        {
            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SecureMailAddin"
            );

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            logFilePath = Path.Combine(folder, "SecureMailAddin.log");

            Log("SYSTEM", "Logger initialized at: " + logFilePath);
        }

        /// <summary>
        /// Basic log method.
        /// </summary>
        public static void Log(string message)
        {
            WriteLog("INFO", "GENERAL", message);
        }

        /// <summary>
        /// Log with custom category.
        /// </summary>
        public static void Log(string category, string message)
        {
            WriteLog("INFO", category, message);
        }

        public static void LogInfo(string category, string message)
        {
            WriteLog("INFO", category, message);
        }

        public static void LogWarning(string category, string message)
        {
            WriteLog("WARN", category, message);
        }

        public static void LogError(string category, string message, Exception ex = null)
        {
            string finalMessage = ex == null
                ? message
                : $"{message} | Exception: {ex.Message}";

            WriteLog("ERROR", category, finalMessage);
        }

        private static void WriteLog(string level, string category, string message)
        {
            if (string.IsNullOrEmpty(logFilePath))
            {
                Initialize();
            }

            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] [{category}] {message}";

            try
            {
                lock (_lock)
                {
                    File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Logger Error: " + ex.Message);
            }

            Debug.WriteLine(logEntry);

#if DEBUG
            // Keep disabled if too noisy. Uncomment only if needed.
            // MessageBox.Show(logEntry, "SecureMail Debug Log", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
        }

        // =========================================================
        // MODULE 6 - ZERO TRUST LOGGING HELPERS
        // =========================================================

        public static void LogZeroTrustAuthorizationStarted(string action, string username = null)
        {
            WriteLog(
                "INFO",
                "ZERO_TRUST",
                $"Authorization started. Action='{action}', Username='{username ?? "unknown"}'");
        }

        public static void LogZeroTrustAuthorizationSucceeded(string action, string username = null, string reason = null)
        {
            WriteLog(
                "INFO",
                "ZERO_TRUST",
                $"Authorization allowed. Action='{action}', Username='{username ?? "unknown"}', Reason='{reason ?? ""}'");
        }

        public static void LogZeroTrustAuthorizationDenied(string action, string username = null, string reason = null)
        {
            WriteLog(
                "WARN",
                "ZERO_TRUST",
                $"Authorization denied. Action='{action}', Username='{username ?? "unknown"}', Reason='{reason ?? ""}'");
        }

        public static void LogZeroTrustAuthorizationFailed(string action, Exception ex, string username = null)
        {
            WriteLog(
                "ERROR",
                "ZERO_TRUST",
                $"Authorization failed. Action='{action}', Username='{username ?? "unknown"}', Exception='{ex.Message}'");
        }

        public static void LogClaimsExtracted(
            string username,
            string email,
            string department,
            string clearance,
            string roles)
        {
            WriteLog(
                "INFO",
                "CLAIMS",
                $"Claims extracted. Username='{username}', Email='{email}', Department='{department}', Clearance='{clearance}', Roles='{roles}'");
        }

        public static void LogPolicyEvaluationStarted(string action)
        {
            WriteLog(
                "INFO",
                "POLICY",
                $"Policy evaluation started. Action='{action}'");
        }

        public static void LogPolicyEvaluationResult(string action, bool isAllowed, string reason)
        {
            WriteLog(
                isAllowed ? "INFO" : "WARN",
                "POLICY",
                $"Policy evaluation completed. Action='{action}', Allowed='{isAllowed}', Reason='{reason}'");
        }

        public static void LogUnauthorizedAccessAttempt(string username, string action, string reason)
        {
            WriteLog(
                "WARN",
                "UNAUTHORIZED_ACCESS",
                $"Unauthorized access attempt. Username='{username ?? "unknown"}', Action='{action}', Reason='{reason}'");
        }

        // =========================================================
        // LOGIN / ADMIN FLOW
        // =========================================================

        public static void LogUserLoginStarted(string username)
        {
            WriteLog(
                "INFO",
                "LOGIN",
                $"User login started. Username='{username}'");
        }

        public static void LogUserLoginSucceeded(string username)
        {
            WriteLog(
                "INFO",
                "LOGIN",
                $"User login successful. Username='{username}'");
        }

        public static void LogUserLoginFailed(string username, string reason)
        {
            WriteLog(
                "WARN",
                "LOGIN",
                $"User login failed. Username='{username}', Reason='{reason}'");
        }

        public static void LogAdminLoginStarted(string username = null)
        {
            WriteLog(
                "INFO",
                "ADMIN_LOGIN",
                $"Admin login started. Username='{username ?? "unknown"}'");
        }

        public static void LogAdminLoginSucceeded(string username = null)
        {
            WriteLog(
                "INFO",
                "ADMIN_LOGIN",
                $"Admin login successful. Username='{username ?? "unknown"}'");
        }

        public static void LogAdminDashboardOpened(string username = null)
        {
            WriteLog(
                "INFO",
                "ADMIN_DASHBOARD",
                $"Admin dashboard opened. Username='{username ?? "unknown"}'");
        }

        public static void LogAdminDashboardDenied(string username = null, string reason = null)
        {
            WriteLog(
                "WARN",
                "ADMIN_DASHBOARD",
                $"Admin dashboard access denied. Username='{username ?? "unknown"}', Reason='{reason ?? ""}'");
        }

        public static void LogPendingUserApproved(string adminUsername, string approvedUsername)
        {
            WriteLog(
                "INFO",
                "ADMIN_ACTION",
                $"Pending user approved. Admin='{adminUsername ?? "unknown"}', ApprovedUser='{approvedUsername ?? "unknown"}'");
        }

        public static void LogPendingUserRejected(string adminUsername, string rejectedUsername)
        {
            WriteLog(
                "INFO",
                "ADMIN_ACTION",
                $"Pending user rejected. Admin='{adminUsername ?? "unknown"}', RejectedUser='{rejectedUsername ?? "unknown"}'");
        }

        // =========================================================
        // MAIL FLOW / CLASSIFICATION / ENFORCEMENT
        // =========================================================

        public static void LogMailBeforeSendStarted(string subject)
        {
            WriteLog(
                "INFO",
                "MAIL",
                $"Before-send processing started. Subject='{subject ?? ""}'");
        }

        public static void LogMailEncryptionCheckFailed(string subject)
        {
            WriteLog(
                "WARN",
                "MAIL_ENCRYPTION",
                $"Mail blocked because encryption was missing. Subject='{subject ?? ""}'");
        }

        public static void LogMailClassificationResult(
            string subject,
            string suggestedTag,
            string sensitivity,
            int totalScore)
        {
            WriteLog(
                "INFO",
                "MAIL_CLASSIFICATION",
                $"Mail classified. Subject='{subject ?? ""}', Tag='{suggestedTag ?? ""}', Sensitivity='{sensitivity ?? ""}', Score='{totalScore}'");
        }

        public static void LogMailZeroTrustCheckStarted(string subject, string suggestedTag)
        {
            WriteLog(
                "INFO",
                "MAIL_ZERO_TRUST",
                $"Mail Zero Trust check started. Subject='{subject ?? ""}', SuggestedTag='{suggestedTag ?? ""}'");
        }

        public static void LogMailBlockedByZeroTrust(string subject, string suggestedTag, string reason)
        {
            WriteLog(
                "WARN",
                "MAIL_ZERO_TRUST",
                $"Mail blocked by Zero Trust. Subject='{subject ?? ""}', SuggestedTag='{suggestedTag ?? ""}', Reason='{reason}'");
        }

        public static void LogMailAllowedByZeroTrust(string subject, string suggestedTag)
        {
            WriteLog(
                "INFO",
                "MAIL_ZERO_TRUST",
                $"Mail allowed by Zero Trust. Subject='{subject ?? ""}', SuggestedTag='{suggestedTag ?? ""}'");
        }

        // =========================================================
        // TOKEN / USER INFO / CLAIMS
        // =========================================================

        public static void LogTokenValidationStarted()
        {
            WriteLog("INFO", "TOKEN", "Token validation started.");
        }

        public static void LogTokenValidationSucceeded()
        {
            WriteLog("INFO", "TOKEN", "Token validation succeeded.");
        }

        public static void LogTokenValidationFailed(string reason)
        {
            WriteLog("ERROR", "TOKEN", $"Token validation failed. Reason='{reason}'");
        }

        public static void LogUserInfoFetchStarted()
        {
            WriteLog("INFO", "USERINFO", "User info fetch started.");
        }

        public static void LogUserInfoFetchSucceeded(string username)
        {
            WriteLog("INFO", "USERINFO", $"User info fetch succeeded. Username='{username ?? "unknown"}'");
        }

        public static void LogUserInfoFetchFailed(string reason)
        {
            WriteLog("ERROR", "USERINFO", $"User info fetch failed. Reason='{reason}'");
        }
    }
}