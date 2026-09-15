using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SecureMailAddin.Keycloak
{
    public class ApiIntegrationService
    {
        private readonly HttpClient _http;
        private readonly KeycloakService _keycloakService;

        public ApiIntegrationService()
        {
            _http = new HttpClient();
            _keycloakService = new KeycloakService();
        }

        private async Task<HttpRequestMessage> CreateAuthorizedRequestAsync(HttpMethod method, string url, object body = null)
        {
            string token = await _keycloakService.GetValidAccessTokenAsync();

            var request = new HttpRequestMessage(method, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            if (body != null)
            {
                string json = JsonConvert.SerializeObject(body);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
            }

            return request;
        }

        private async Task<HttpResponseMessage> SendWithRetryAsync(Func<Task<HttpRequestMessage>> requestFactory)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                HttpRequestMessage request = null;
                try
                {
                    request = await requestFactory();
                    var response = await _http.SendAsync(request);

                    if ((int)response.StatusCode >= 500)
                    {
                        await Task.Delay(1000 * attempt);
                        continue;
                    }

                    return response;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    await Task.Delay(1000 * attempt);
                }
                finally
                {
                    request?.Dispose();
                }
            }

            if (lastException != null)
                throw lastException;

            throw new Exception("Request failed after retry attempts.");
        }

        public async Task<PolicyResponse> GetClassificationRulesAsync()
        {
            try
            {
                var response = await SendWithRetryAsync(() =>
                    CreateAuthorizedRequestAsync(
                        HttpMethod.Get,
                        $"{KeycloakConfig.PolicyApiBaseUrl}/api/policy/current"));

                string json = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var policy = JsonConvert.DeserializeObject<PolicyResponse>(json);

                    if (policy != null)
                        PolicyCacheService.Save(policy);

                    return policy;
                }

                var cachedPolicy = PolicyCacheService.Load();
                if (cachedPolicy != null)
                    return cachedPolicy;

                throw new Exception("Failed to fetch classification rules: " + response.StatusCode + " - " + json);
            }
            catch
            {
                var cachedPolicy = PolicyCacheService.Load();
                if (cachedPolicy != null)
                    return cachedPolicy;

                throw;
            }
        }

        public async Task SendAuditLogAsync(AuditLogRequest requestModel)
        {
            var response = await SendWithRetryAsync(() =>
                CreateAuthorizedRequestAsync(
                    HttpMethod.Post,
                    $"{KeycloakConfig.AuditApiBaseUrl}/api/audit/log",
                    requestModel));

            string json = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Failed to send audit log: " + response.StatusCode + " - " + json);
        }

        public async Task SubmitMailClassificationAsync(MailClassificationRequest requestModel)
        {
            var response = await SendWithRetryAsync(() =>
                CreateAuthorizedRequestAsync(
                    HttpMethod.Post,
                    $"{KeycloakConfig.AuditApiBaseUrl}/api/audit/mail-event",
                    requestModel));

            string json = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Failed to submit classification event: " + response.StatusCode + " - " + json);
        }

        public async Task<string> GetRecipientClearanceAsync(string email)
        {
            return await _keycloakService.GetUserClearanceByEmailAsync(email);
        }
    }
}