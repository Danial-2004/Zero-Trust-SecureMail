using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SecureMailAddin.ZeroTrust
{
    /// <summary>
    /// Performs real JWT validation before authorization: signature, allowed
    /// algorithm, issuer, audience, lifetime, required claims and optional jti
    /// replay detection. Any error or unavailable JWKS service fails closed.
    /// </summary>
    public sealed class JwtTokenValidator : IJwtTokenValidator
    {
        private const string ValidatorId = "SecureMailAddin.ZeroTrust.JwtTokenValidator.v1";
        private readonly JwtValidationOptions _options;

        public JwtTokenValidator(JwtValidationOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            _options = options;
        }

        public TokenValidationOutcome Validate(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Missing,
                    PolicyCodes.JwtMissing,
                    "Authentication token is missing.");
            }

            try
            {
                string[] parts = token.Split('.');
                if (parts.Length != 3 || parts.Any(p => string.IsNullOrEmpty(p)))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.Malformed,
                        PolicyCodes.JwtMalformed,
                        "Authentication token is not a compact JWS.");
                }

                var handler = new JwtSecurityTokenHandler { MapInboundClaims = false };
                if (!handler.CanReadToken(token))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.Malformed,
                        PolicyCodes.JwtMalformed,
                        "Authentication token cannot be read.");
                }

                JwtSecurityToken unvalidated = handler.ReadJwtToken(token);
                string algorithm = unvalidated.Header == null ? null : unvalidated.Header.Alg;

                if (string.IsNullOrWhiteSpace(algorithm) ||
                    !_options.AllowedAlgorithms.Contains(algorithm, StringComparer.OrdinalIgnoreCase))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.Invalid,
                        PolicyCodes.JwtAlgorithmNotAllowed,
                        "JWT signing algorithm is not allowed.");
                }

                IList<SecurityKey> keys = _options.SigningKeyProvider == null
                    ? null
                    : _options.SigningKeyProvider.GetSigningKeys();

                if (keys == null || keys.Count == 0)
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.ValidationUnavailable,
                        PolicyCodes.AuthenticationServiceUnavailable,
                        "Trusted signing keys are unavailable.");
                }

                var parameters = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidIssuer = _options.Issuer,
                    ValidateAudience = true,
                    ValidAudience = _options.Audience,
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKeys = keys,
                    ValidateLifetime = true,
                    RequireExpirationTime = true,
                    RequireSignedTokens = true,
                    ClockSkew = _options.ClockSkew,
                    ValidAlgorithms = _options.AllowedAlgorithms,
                    NameClaimType = "preferred_username",
                    RoleClaimType = "roles"
                };

                SecurityToken validatedToken;
                ClaimsPrincipal principal = handler.ValidateToken(token, parameters, out validatedToken);

                if (principal == null)
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.Invalid,
                        PolicyCodes.JwtInvalid,
                        "JWT validation did not produce a principal.");
                }

                JwtSecurityToken jwt = validatedToken as JwtSecurityToken;
                if (jwt == null || jwt.Payload == null)
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.Invalid,
                        PolicyCodes.JwtInvalid,
                        "Validated token is not a JWT.");
                }

                string subject = FirstNonEmpty(
                    ReadPayloadString(jwt.Payload, "sub"),
                    jwt.Subject);

                string username = FirstNonEmpty(
                    ReadPayloadString(jwt.Payload, "preferred_username"),
                    subject);

                string email = ReadPayloadString(jwt.Payload, "email") ?? string.Empty;
                string department = ReadPayloadString(jwt.Payload, "department") ?? string.Empty;
                string clearance = ReadPayloadString(jwt.Payload, "clearance") ?? string.Empty;
                string tokenId = FirstNonEmpty(
                    ReadPayloadString(jwt.Payload, "jti"),
                    jwt.Id);

                IList<string> roles = ExtractRoles(jwt.Payload, _options.ClientId);
                string audience = jwt.Audiences == null
                    ? string.Empty
                    : jwt.Audiences.FirstOrDefault() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(subject))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.MissingClaims,
                        PolicyCodes.SubjectMissing,
                        "Validated token is missing the subject claim.");
                }

                if (roles == null || roles.Count == 0)
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.MissingClaims,
                        PolicyCodes.RoleMissing,
                        "Validated token is missing role claims.");
                }

                if (string.IsNullOrWhiteSpace(clearance))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.MissingClaims,
                        PolicyCodes.ClearanceMissing,
                        "Validated token is missing the trusted clearance claim.");
                }

                // jti identifies the JWT, not an individual request. A valid
                // OAuth bearer access token is reusable until it expires, so
                // the jti is bound into the authentication proof but is NOT
                // used to enforce one-time token use. Request-level replay
                // protection is handled separately (RequestReplayCache).
                if (_options.RequireJti && string.IsNullOrWhiteSpace(tokenId))
                {
                    return TokenValidationOutcome.Failure(
                        JwtValidationStatus.MissingClaims,
                        PolicyCodes.JwtTokenIdMissing,
                        "Validated token is missing the jti claim.");
                }

                var unsignedIdentity = new ValidatedTokenIdentity(
                    null,
                    subject,
                    username,
                    email,
                    department,
                    clearance,
                    roles,
                    jwt.Issuer ?? string.Empty,
                    audience,
                    tokenId ?? string.Empty,
                    jwt.ValidTo.ToUniversalTime());

                TrustedAuthenticationProof proof =
                    TrustedAuthenticationProof.Issue(unsignedIdentity, ValidatorId);

                var identity = new ValidatedTokenIdentity(
                    proof,
                    subject,
                    username,
                    email,
                    department,
                    clearance,
                    roles,
                    jwt.Issuer ?? string.Empty,
                    audience,
                    tokenId ?? string.Empty,
                    jwt.ValidTo.ToUniversalTime());

                return TokenValidationOutcome.Success(identity);
            }
            catch (Exception ex)
            {
                return MapException(ex);
            }
        }

        private static TokenValidationOutcome MapException(Exception ex)
        {
            string typeName = ex.GetType().Name;

            if (typeName.IndexOf("Expired", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Expired,
                    PolicyCodes.JwtExpired,
                    "Authentication token has expired.");
            }

            if (typeName.IndexOf("NotYetValid", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.NotYetValid,
                    PolicyCodes.JwtNotYetValid,
                    "Authentication token is not yet valid.");
            }

            if (typeName.IndexOf("Signature", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("SigningKey", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.InvalidSignature,
                    PolicyCodes.JwtSignatureInvalid,
                    "JWT signature or signing key validation failed.");
            }

            if (typeName.IndexOf("Issuer", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.InvalidIssuer,
                    PolicyCodes.JwtIssuerInvalid,
                    "JWT issuer is not trusted.");
            }

            if (typeName.IndexOf("Audience", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.InvalidAudience,
                    PolicyCodes.JwtAudienceInvalid,
                    "JWT audience is not intended for SecureMail.");
            }

            if (typeName.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("NoExpiration", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Expired,
                    PolicyCodes.JwtLifetimeInvalid,
                    "JWT lifetime is invalid or missing.");
            }

            if (typeName.IndexOf("Replay", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Replayed,
                    PolicyCodes.JwtReplayDetected,
                    "JWT replay was detected.");
            }

            if (typeName.IndexOf("Algorithm", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Invalid,
                    PolicyCodes.JwtAlgorithmNotAllowed,
                    "JWT signing algorithm is not allowed.");
            }

            if (typeName.IndexOf("SecurityToken", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("Format", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("Argument", StringComparison.OrdinalIgnoreCase) >= 0 ||
                typeName.IndexOf("Json", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.Invalid,
                    PolicyCodes.JwtInvalid,
                    "Authentication token is invalid or malformed.");
            }

            return TokenValidationOutcome.Failure(
                JwtValidationStatus.ValidationUnavailable,
                PolicyCodes.AuthenticationServiceUnavailable,
                "Authentication validation service is unavailable; failing closed.");
        }

        private static IList<string> ExtractRoles(JwtPayload payload, string clientId)
        {
            var roles = new List<string>();

            object realmAccess;
            if (payload.TryGetValue("realm_access", out realmAccess))
                AddRoles(ToJToken(realmAccess), roles);

            object directRoles;
            if (payload.TryGetValue("roles", out directRoles))
                AddRoles(ToJToken(directRoles), roles);

            object resourceAccess;
            if (payload.TryGetValue("resource_access", out resourceAccess))
            {
                JObject resourceObject = ToJToken(resourceAccess) as JObject;
                if (resourceObject != null)
                {
                    foreach (JProperty property in resourceObject.Properties())
                    {
                        if (!string.IsNullOrWhiteSpace(clientId) &&
                            !string.Equals(property.Name, clientId, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        JObject clientObject = property.Value as JObject;
                        if (clientObject != null)
                            AddRoles(clientObject["roles"], roles);
                    }
                }
            }

            return roles
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void AddRoles(JToken token, List<string> roles)
        {
            if (token == null)
                return;

            JObject obj = token as JObject;
            if (obj != null)
            {
                AddRoles(obj["roles"], roles);
                return;
            }

            JArray array = token as JArray;
            if (array != null)
            {
                foreach (JToken item in array)
                {
                    if (item != null && item.Type == JTokenType.String)
                        roles.Add(item.ToString());
                }

                return;
            }

            if (token.Type == JTokenType.String)
                roles.Add(token.ToString());
        }

        private static JToken ToJToken(object value)
        {
            if (value == null)
                return null;

            JToken existing = value as JToken;
            if (existing != null)
                return existing;

            // IdentityModel 8 returns nested JWT claim values as System.Text.Json
            // JsonElement instances. Convert them before role/claim parsing.
            if (value is JsonElement)
            {
                JsonElement element = (JsonElement)value;
                try
                {
                    return JToken.Parse(element.GetRawText());
                }
                catch
                {
                    return null;
                }
            }

            string text = value as string;
            if (text != null)
            {
                string trimmed = text.TrimStart();
                if (trimmed.StartsWith("{", StringComparison.Ordinal) ||
                    trimmed.StartsWith("[", StringComparison.Ordinal))
                {
                    try { return JToken.Parse(text); }
                    catch { return new JValue(text); }
                }

                return new JValue(text);
            }

            try { return JToken.FromObject(value); }
            catch { return new JValue(JsonConvert.SerializeObject(value)); }
        }

        private static string ReadPayloadString(JwtPayload payload, string key)
        {
            object value;
            if (payload == null || !payload.TryGetValue(key, out value) || value == null)
                return null;

            JToken token = ToJToken(value);
            if (token == null)
                return null;

            JArray array = token as JArray;
            if (array != null)
            {
                foreach (JToken item in array)
                {
                    if (item != null && item.Type == JTokenType.String)
                        return item.ToString();
                }

                return null;
            }

            if (token.Type == JTokenType.String)
                return token.ToString();

            return token.ToString(Formatting.None);
        }

        private static string FirstNonEmpty(params string[] values)
        {
            if (values == null)
                return null;

            for (int i = 0; i < values.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(values[i]))
                    return values[i].Trim();
            }

            return null;
        }
    }
}
