using System;
using System.Linq;
using SecureMailAddin.Keycloak;

namespace SecureMailAddin.ZeroTrust
{
    public class ClaimsExtractor
    {
        /// <summary>
        /// Converts cryptographically validated JWT claims into an authorization
        /// profile. The proof binds every field copied here.
        /// </summary>
        public UserClaimsProfile Extract(ValidatedTokenIdentity identity)
        {
            if (identity == null)
                throw new ArgumentNullException(nameof(identity));

            if (identity.Proof == null)
                throw new InvalidOperationException("Validated identity does not contain a trusted authentication proof.");

            return new UserClaimsProfile
            {
                Subject = identity.Subject ?? string.Empty,
                Username = identity.Username ?? string.Empty,
                Email = identity.Email ?? string.Empty,
                Department = identity.Department ?? string.Empty,
                ClearanceLevel = identity.ClearanceLevel ?? string.Empty,
                Roles = identity.Roles == null
                    ? new System.Collections.Generic.List<string>()
                    : identity.Roles.ToList(),
                AuthenticationProof = identity.Proof
            };
        }

        /// <summary>
        /// Legacy method retained for compatibility with other code. It does NOT
        /// attach a trusted authentication proof, so AccessPolicyEngine will
        /// deny any profile produced by this overload.
        /// </summary>
        [Obsolete("User-info responses are not trusted authentication evidence. Use Extract(ValidatedTokenIdentity) after JWT validation.")]
        public UserClaimsProfile Extract(KeycloakUserInfo userInfo)
        {
            if (userInfo == null)
                throw new ArgumentNullException(nameof(userInfo));

            var profile = new UserClaimsProfile
            {
                Subject = userInfo.sub ?? string.Empty,
                Username = userInfo.preferred_username ?? string.Empty,
                Email = userInfo.email ?? string.Empty,
                Department = userInfo.department ?? string.Empty,
                ClearanceLevel = userInfo.clearance ?? string.Empty
            };

            if (userInfo.realm_access != null && userInfo.realm_access.roles != null)
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
