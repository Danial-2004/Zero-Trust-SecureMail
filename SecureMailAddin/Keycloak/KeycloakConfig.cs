using System;

namespace SecureMailAddin.Keycloak
{
    /// <summary>
    /// Environment-first configuration. No client secret is committed to source.
    /// Existing local KeycloakConfig.cs files can override these property names;
    /// JwtValidationOptions also discovers them through reflection.
    /// </summary>
    public static class KeycloakConfig
    {
        public static string BaseUrl
        {
            get { return Env("SECUREMAIL_KEYCLOAK_BASE_URL", "http://localhost:8080"); }
        }

        public static string Realm
        {
            get { return Env("SECUREMAIL_KEYCLOAK_REALM", "SecureMailRealm"); }
        }

        public static string ClientId
        {
            get { return Env("SECUREMAIL_KEYCLOAK_CLIENT_ID", "securemail-client"); }
        }

        public static string ClientSecret
        {
            get { return Env("SECUREMAIL_KEYCLOAK_CLIENT_SECRET", string.Empty); }
        }

        public static string Issuer
        {
            get
            {
                string configured = Env("SECUREMAIL_JWT_ISSUER", null);
                if (!string.IsNullOrWhiteSpace(configured))
                    return configured.Trim();

                return BaseUrl.TrimEnd('/') + "/realms/" + Realm;
            }
        }

        public static string JwksUrl
        {
            get
            {
                string configured = Env("SECUREMAIL_KEYCLOAK_JWKS_URL", null);
                if (!string.IsNullOrWhiteSpace(configured))
                    return configured.Trim();

                return Issuer.TrimEnd('/') + "/protocol/openid-connect/certs";
            }
        }

        public static string TokenEndpoint
        {
            get { return Issuer.TrimEnd('/') + "/protocol/openid-connect/token"; }
        }

        public static string UserInfoEndpoint
        {
            get { return Issuer.TrimEnd('/') + "/protocol/openid-connect/userinfo"; }
        }

        public static string AdminUsersEndpoint
        {
            get { return BaseUrl.TrimEnd('/') + "/admin/realms/" + Realm + "/users"; }
        }

        public static string PolicyApiBaseUrl
        {
            get { return Env("SECUREMAIL_POLICY_API_BASE_URL", "https://localhost:7132"); }
        }

        public static string AuditApiBaseUrl
        {
            get { return Env("SECUREMAIL_AUDIT_API_BASE_URL", "https://localhost:7280"); }
        }

        private static string Env(string name, string fallback)
        {
            string value = Environment.GetEnvironmentVariable(name);
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }
    }
}
