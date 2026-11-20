using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using Microsoft.IdentityModel.Tokens;

namespace SecureMailAddin.ZeroTrust
{
    public interface ISigningKeyProvider
    {
        IList<SecurityKey> GetSigningKeys();
    }

    public sealed class StaticSigningKeyProvider : ISigningKeyProvider
    {
        private readonly ReadOnlyCollection<SecurityKey> _keys;

        public StaticSigningKeyProvider(IEnumerable<SecurityKey> keys)
        {
            if (keys == null)
                throw new ArgumentNullException(nameof(keys));

            _keys = new ReadOnlyCollection<SecurityKey>(keys.ToList());
        }

        public IList<SecurityKey> GetSigningKeys()
        {
            return _keys;
        }
    }

    /// <summary>
    /// Fetches Keycloak realm signing keys from the JWKS endpoint. Fails closed:
    /// no expired cached keys are used when the validation service is unavailable.
    /// </summary>
    public sealed class JwksSigningKeyProvider : ISigningKeyProvider
    {
        private readonly string _jwksUrl;
        private readonly HttpClient _http;
        private readonly TimeSpan _cacheLifetime;
        private readonly object _sync = new object();
        private IList<SecurityKey> _cachedKeys;
        private DateTime _cachedUntilUtc;

        public JwksSigningKeyProvider(string jwksUrl, TimeSpan? cacheLifetime = null)
        {
            if (string.IsNullOrWhiteSpace(jwksUrl))
                throw new ArgumentException("JWKS URL is required.", nameof(jwksUrl));

            _jwksUrl = jwksUrl;
            _cacheLifetime = cacheLifetime ?? TimeSpan.FromMinutes(15);
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        }

        public IList<SecurityKey> GetSigningKeys()
        {
            lock (_sync)
            {
                DateTime now = DateTime.UtcNow;

                if (_cachedKeys != null && _cachedKeys.Count > 0 && now < _cachedUntilUtc)
                    return _cachedKeys;

                string json = _http.GetStringAsync(_jwksUrl).GetAwaiter().GetResult();
                var set = new JsonWebKeySet(json);
                var keys = set.GetSigningKeys().ToList();

                if (keys.Count == 0)
                    throw new InvalidOperationException("The Keycloak JWKS endpoint returned no signing keys.");

                _cachedKeys = new ReadOnlyCollection<SecurityKey>(keys);
                _cachedUntilUtc = now.Add(_cacheLifetime);
                return _cachedKeys;
            }
        }
    }

    public sealed class JwtValidationOptions
    {
        private static readonly string[] DefaultAlgorithms =
        {
            "RS256", "RS384", "RS512", "PS256", "PS384", "PS512"
        };

        private JwtValidationOptions(
            string issuer,
            string audience,
            string clientId,
            ISigningKeyProvider signingKeyProvider,
            IEnumerable<SecurityKey> staticSigningKeys,
            TimeSpan clockSkew,
            IEnumerable<string> allowedAlgorithms,
            bool requireJti)
        {
            Issuer = issuer;
            Audience = audience;
            ClientId = clientId ?? audience;
            SigningKeyProvider = signingKeyProvider ??
                (staticSigningKeys == null
                    ? null
                    : new StaticSigningKeyProvider(staticSigningKeys));
            ClockSkew = clockSkew;
            AllowedAlgorithms = new ReadOnlyCollection<string>(
                (allowedAlgorithms ?? DefaultAlgorithms).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
            RequireJti = requireJti;
        }

        public string Issuer { get; private set; }
        public string Audience { get; private set; }
        public string ClientId { get; private set; }
        public ISigningKeyProvider SigningKeyProvider { get; private set; }
        public TimeSpan ClockSkew { get; private set; }
        public IList<string> AllowedAlgorithms { get; private set; }
        public bool RequireJti { get; private set; }

        internal static JwtValidationOptions ForUnitTests(
            string issuer,
            string audience,
            IEnumerable<SecurityKey> keys,
            bool requireJti = false)
        {
            return new JwtValidationOptions(
                issuer,
                audience,
                audience,
                null,
                keys,
                TimeSpan.FromSeconds(5),
                DefaultAlgorithms,
                requireJti);
        }

        internal static JwtValidationOptions ForUnitTestsWithKeyProvider(
            string issuer,
            string audience,
            ISigningKeyProvider keyProvider,
            bool requireJti = false)
        {
            return new JwtValidationOptions(
                issuer,
                audience,
                audience,
                keyProvider,
                null,
                TimeSpan.FromSeconds(5),
                DefaultAlgorithms,
                requireJti,
                replayCache);
        }

        public static bool TryFromEnvironment(
            out JwtValidationOptions options,
            out string configurationError)
        {
            options = null;
            configurationError = null;

            string baseUrl = FirstNonEmpty(
                Env("SECUREMAIL_KEYCLOAK_BASE_URL"),
                Config("BaseUrl"),
                "http://localhost:8080");

            string realm = FirstNonEmpty(
                Env("SECUREMAIL_KEYCLOAK_REALM"),
                Config("Realm"),
                "SecureMailRealm");

            string issuer = FirstNonEmpty(
                Env("SECUREMAIL_JWT_ISSUER"),
                Config("Issuer"),
                baseUrl.TrimEnd('/') + "/realms/" + realm);

            string clientId = FirstNonEmpty(
                Env("SECUREMAIL_KEYCLOAK_CLIENT_ID"),
                Env("SECUREMAIL_JWT_CLIENT_ID"),
                Config("ClientId"));

            string audience = FirstNonEmpty(
                Env("SECUREMAIL_JWT_AUDIENCE"),
                clientId);

            if (string.IsNullOrWhiteSpace(issuer) || string.IsNullOrWhiteSpace(audience))
            {
                configurationError =
                    "JWT issuer/audience is not configured. Set SECUREMAIL_JWT_ISSUER and " +
                    "SECUREMAIL_JWT_AUDIENCE (or SECUREMAIL_KEYCLOAK_CLIENT_ID).";
                return false;
            }

            string jwksUrl = FirstNonEmpty(
                Env("SECUREMAIL_KEYCLOAK_JWKS_URL"),
                Env("SECUREMAIL_JWT_JWKS_URL"),
                Config("JwksUrl"),
                issuer.TrimEnd('/') + "/protocol/openid-connect/certs");

            int skewSeconds;
            if (!int.TryParse(Env("SECUREMAIL_JWT_CLOCK_SKEW_SECONDS"), out skewSeconds))
                skewSeconds = 30;

            if (skewSeconds < 0)
                skewSeconds = 0;
            if (skewSeconds > 120)
                skewSeconds = 120;

            bool requireJti = true;
            string requireJtiText = Env("SECUREMAIL_JWT_REQUIRE_JTI");
            if (string.Equals(requireJtiText, "false", StringComparison.OrdinalIgnoreCase) ||
                requireJtiText == "0")
            {
                requireJti = false;
            }

            try
            {
                options = new JwtValidationOptions(
                    issuer,
                    audience,
                    clientId ?? audience,
                    new JwksSigningKeyProvider(jwksUrl),
                    null,
                    TimeSpan.FromSeconds(skewSeconds),
                    DefaultAlgorithms,
                    requireJti);

                return true;
            }
            catch (Exception ex)
            {
                configurationError = "JWT validation configuration is invalid: " + ex.Message;
                return false;
            }
        }

        private static string Env(string name)
        {
            return Environment.GetEnvironmentVariable(name);
        }

        private static string Config(string propertyName)
        {
            try
            {
                Type type =
                    Type.GetType("SecureMailAddin.Keycloak.KeycloakConfig, SecureMailAddin") ??
                    Type.GetType("SecureMailAddin.Keycloak.KeycloakConfig");

                if (type == null)
                    return null;

                PropertyInfo property = type.GetProperty(
                    propertyName,
                    BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);

                return property == null ? null : property.GetValue(null, null) as string;
            }
            catch
            {
                return null;
            }
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
