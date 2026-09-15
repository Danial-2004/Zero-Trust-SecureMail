using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SecureMailAddin.Keycloak;

namespace SecureMailAddin.ZeroTrust
{
    public class UnauthorizedAccessLogger
    {
        private readonly ApiIntegrationService _apiService;

        public UnauthorizedAccessLogger()
        {
            _apiService = new ApiIntegrationService();
        }

        public async Task LogDeniedActionAsync(string username, string action, string reason)
        {
            await _apiService.SendAuditLogAsync(new AuditLogRequest
            {
                EventType = "UNAUTHORIZED_ACCESS_ATTEMPT",
                Username = username ?? "unknown",
                EmailSubject = "",
                Classification = "",
                TimestampUtc = DateTime.UtcNow,
                Details = $"Action='{action}', Reason='{reason}'"
            });
        }
    }
}