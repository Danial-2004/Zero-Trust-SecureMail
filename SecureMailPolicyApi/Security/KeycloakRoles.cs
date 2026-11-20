using System;
using System.Security.Claims;
using System.Text.Json;

namespace SecureMailPolicyApi.Security
{
    public static class KeycloakRoles
    {
        public static bool HasRealmRole(ClaimsPrincipal user, string requiredRole)
        {
            if (user == null || string.IsNullOrWhiteSpace(requiredRole))
                return false;

            foreach (Claim claim in user.Claims)
            {
                if (claim == null || string.IsNullOrWhiteSpace(claim.Value))
                    continue;

                bool directRoleClaim =
                    claim.Type == "roles" ||
                    claim.Type == "role" ||
                    claim.Type == ClaimTypes.Role;

                if (directRoleClaim)
                {
                    if (string.Equals(claim.Value, requiredRole, StringComparison.OrdinalIgnoreCase) ||
                        JsonArrayContains(claim.Value, requiredRole))
                    {
                        return true;
                    }
                }

                if (claim.Type == "realm_access" && RealmAccessContains(claim.Value, requiredRole))
                    return true;
            }

            return false;
        }

        private static bool RealmAccessContains(string json, string requiredRole)
        {
            try
            {
                using (JsonDocument document = JsonDocument.Parse(json))
                {
                    JsonElement roles;
                    if (document.RootElement.ValueKind != JsonValueKind.Object ||
                        !document.RootElement.TryGetProperty("roles", out roles) ||
                        roles.ValueKind != JsonValueKind.Array)
                    {
                        return false;
                    }

                    foreach (JsonElement role in roles.EnumerateArray())
                    {
                        if (role.ValueKind == JsonValueKind.String &&
                            string.Equals(role.GetString(), requiredRole, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (JsonException)
            {
                return false;
            }

            return false;
        }

        private static bool JsonArrayContains(string json, string requiredRole)
        {
            string trimmed = json.TrimStart();
            if (!trimmed.StartsWith("[", StringComparison.Ordinal))
                return false;

            try
            {
                using (JsonDocument document = JsonDocument.Parse(json))
                {
                    if (document.RootElement.ValueKind != JsonValueKind.Array)
                        return false;

                    foreach (JsonElement role in document.RootElement.EnumerateArray())
                    {
                        if (role.ValueKind == JsonValueKind.String &&
                            string.Equals(role.GetString(), requiredRole, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (JsonException)
            {
                return false;
            }

            return false;
        }
    }
}
