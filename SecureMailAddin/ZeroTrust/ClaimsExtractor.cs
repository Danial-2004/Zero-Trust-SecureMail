using System;
using System.Linq;
using SecureMailAddin.Keycloak;

namespace SecureMailAddin.ZeroTrust
{
    public class ClaimsExtractor
    {
        public UserClaimsProfile Extract(KeycloakUserInfo userInfo)
        {
            if (userInfo == null)
                throw new ArgumentNullException(nameof(userInfo));

            var profile = new UserClaimsProfile
            {
                Username = userInfo.preferred_username ?? "",
                Email = userInfo.email ?? "",
                Department = userInfo.department ?? "",
                ClearanceLevel = userInfo.clearance ?? ""
            };

            if (userInfo.realm_access?.roles != null)
            {
                profile.Roles = userInfo.realm_access.roles
                    .Where(r => !string.IsNullOrWhiteSpace(r))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            return profile;
        }
    }
}