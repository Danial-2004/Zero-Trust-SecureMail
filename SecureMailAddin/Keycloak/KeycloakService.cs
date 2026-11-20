using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SecureMailAddin.Keycloak
{
    /// <summary>
    /// Keycloak token/user service. This handles authentication transport and
    /// admin user lookup. It is NOT the authorization decision maker, and it is
    /// NOT a JWT validator: ZeroTrustHandler performs local cryptographic JWT
    /// validation before AccessPolicyEngine evaluates any request.
    /// </summary>
    public class KeycloakService
    {
        private static readonly HttpClient SharedHttpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(20)
        };

        private readonly HttpClient _http;

        public KeycloakService()
            : this(SharedHttpClient)
        {
        }

        public KeycloakService(HttpClient httpClient)
        {
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        }

        public async Task<AuthResult> LoginAsync(string username, string password)
        {
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
                throw new ArgumentException("Username and password are required.");

            var fields = new Dictionary<string, string>
            {
                { "grant_type", "password" },
                { "client_id", KeycloakConfig.ClientId },
                { "username", username.Trim() },
                { "password", password },
                { "scope", "openid profile email" }
            };

            AddClientSecret(fields);

            AuthResult result = await RequestTokenAsync(fields).ConfigureAwait(false);
            SaveToken(result);
            return result;
        }

        public async Task<string> GetValidAccessTokenAsync()
        {
            StoredTokenInfo stored = TokenStore.Load();

            if (stored == null || string.IsNullOrWhiteSpace(stored.AccessToken))
                return null;

            // Refresh slightly before expiry so a token cannot expire in flight.
            if (stored.ExpiresAt > DateTime.UtcNow.AddSeconds(30))
                return stored.AccessToken;

            if (string.IsNullOrWhiteSpace(stored.RefreshToken))
                return null;

            var fields = new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", KeycloakConfig.ClientId },
                { "refresh_token", stored.RefreshToken }
            };

            AddClientSecret(fields);

            AuthResult refreshed = await RequestTokenAsync(fields).ConfigureAwait(false);
            SaveToken(refreshed);
            return refreshed.AccessToken;
        }

        public async Task<KeycloakUserInfo> GetUserInfoWithUserTokenAsync(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
                return null;

            using (var request = new HttpRequestMessage(HttpMethod.Get, KeycloakConfig.UserInfoEndpoint))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using (HttpResponseMessage response = await _http.SendAsync(request).ConfigureAwait(false))
                {
                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "Keycloak user-info request failed: " + (int)response.StatusCode);
                    }

                    return JsonConvert.DeserializeObject<KeycloakUserInfo>(json);
                }
            }
        }

        public async Task<bool> UserExistsByEmailAsync(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                throw new ArgumentException("Email is required.", nameof(email));

            string url = KeycloakConfig.AdminUsersEndpoint +
                         "?email=" + Uri.EscapeDataString(email.Trim()) +
                         "&exact=true";

            string json = await AdminGetAsync(url).ConfigureAwait(false);
            var users = JsonConvert.DeserializeObject<List<KeycloakUser>>(json) ?? new List<KeycloakUser>();

            return users.Any(u => string.Equals(
                u.Email,
                email.Trim(),
                StringComparison.OrdinalIgnoreCase));
        }

        public async Task<string> GetUserClearanceByEmailAsync(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                throw new ArgumentException("Email is required.", nameof(email));

            string url = KeycloakConfig.AdminUsersEndpoint +
                         "?email=" + Uri.EscapeDataString(email.Trim()) +
                         "&exact=true";

            string json = await AdminGetAsync(url).ConfigureAwait(false);
            var users = JsonConvert.DeserializeObject<List<KeycloakUser>>(json) ?? new List<KeycloakUser>();

            KeycloakUser user = users.FirstOrDefault(u => string.Equals(
                u.Email,
                email.Trim(),
                StringComparison.OrdinalIgnoreCase));

            if (user == null || user.Attributes == null)
            {
                throw new KeyNotFoundException(
                    "No trusted clearance attribute was found for the recipient.");
            }

            List<string> values;
            if (!user.Attributes.TryGetValue("clearance", out values) ||
                values == null ||
                values.Count == 0 ||
                string.IsNullOrWhiteSpace(values[0]))
            {
                throw new KeyNotFoundException(
                    "The recipient has no trusted clearance attribute.");
            }

            return values[0].Trim();
        }

        public async Task RegisterUserAsync(
            string username,
            string email,
            string password,
            string firstName,
            string lastName,
            bool emailVerified)
        {
            if (string.IsNullOrWhiteSpace(username) ||
                string.IsNullOrWhiteSpace(email) ||
                string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException("Username, email and password are required.");
            }

            var representation = new
            {
                username = username.Trim(),
                email = email.Trim(),
                firstName = firstName ?? string.Empty,
                lastName = lastName ?? string.Empty,
                enabled = true,
                emailVerified = emailVerified,
                credentials = new[]
                {
                    new
                    {
                        type = "password",
                        value = password,
                        temporary = false
                    }
                },
                attributes = new Dictionary<string, List<string>>
                {
                    { "clearance", new List<string> { "PUBLIC" } },
                    { "department", new List<string> { string.Empty } }
                }
            };

            string token = await RequireCurrentTokenAsync().ConfigureAwait(false);
            string json = JsonConvert.SerializeObject(representation);

            using (var request = new HttpRequestMessage(HttpMethod.Post, KeycloakConfig.AdminUsersEndpoint))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                using (HttpResponseMessage response = await _http.SendAsync(request).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "Keycloak user registration failed: " + (int)response.StatusCode + " " + body);
                    }
                }
            }
        }

        private async Task<string> AdminGetAsync(string url)
        {
            string token = await RequireCurrentTokenAsync().ConfigureAwait(false);

            using (var request = new HttpRequestMessage(HttpMethod.Get, url))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using (HttpResponseMessage response = await _http.SendAsync(request).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "Keycloak admin request failed: " + (int)response.StatusCode + " " + body);
                    }

                    return body;
                }
            }
        }

        private async Task<string> RequireCurrentTokenAsync()
        {
            string token = await GetValidAccessTokenAsync().ConfigureAwait(false);

            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException(
                    "A valid access token is required for Keycloak administration; failing closed.");
            }

            return token;
        }

        private async Task<AuthResult> RequestTokenAsync(Dictionary<string, string> fields)
        {
            using (var content = new FormUrlEncodedContent(fields))
            using (HttpResponseMessage response = await _http
                .PostAsync(KeycloakConfig.TokenEndpoint, content)
                .ConfigureAwait(false))
            {
                string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(
                        "Keycloak token request failed: " + (int)response.StatusCode);
                }

                AuthResult result = JsonConvert.DeserializeObject<AuthResult>(json);

                if (result == null || string.IsNullOrWhiteSpace(result.AccessToken))
                    throw new InvalidOperationException("Keycloak returned no access token.");

                int lifetimeSeconds = result.ExpiresIn > 0 ? result.ExpiresIn : 300;
                result.ExpiresAt = DateTime.UtcNow.AddSeconds(lifetimeSeconds);
                return result;
            }
        }

        private static void AddClientSecret(Dictionary<string, string> fields)
        {
            string secret = KeycloakConfig.ClientSecret;
            if (!string.IsNullOrWhiteSpace(secret))
                fields["client_secret"] = secret;
        }

        private static void SaveToken(AuthResult auth)
        {
            TokenStore.Save(new StoredTokenInfo
            {
                AccessToken = auth.AccessToken,
                RefreshToken = auth.RefreshToken,
                ExpiresAt = auth.ExpiresAt
            });
        }
    }
}
