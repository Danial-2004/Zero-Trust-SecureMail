using System;
using System.Threading.Tasks;
using SecureMailAddin.Keycloak;

namespace SecureMailAddin.ZeroTrust
{
    public class ZeroTrustHandler
    {
        private readonly KeycloakService _keycloakService;
        private readonly ClaimsExtractor _claimsExtractor;
        private readonly AccessPolicyEngine _policyEngine;
        private readonly UnauthorizedAccessLogger _logger;

        public ZeroTrustHandler()
        {
            _keycloakService = new KeycloakService();
            _claimsExtractor = new ClaimsExtractor();
            _policyEngine = new AccessPolicyEngine();
            _logger = new UnauthorizedAccessLogger();
        }

        // Existing flow: use currently valid token from Keycloak service
        public async Task<AccessDecision> AuthorizeAsync(
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            try
            {
                string token = await _keycloakService.GetValidAccessTokenAsync();

                if (string.IsNullOrWhiteSpace(token))
                {
                    return new AccessDecision
                    {
                        IsAllowed = false,
                        Reason = "Access token is missing."
                    };
                }

                return await AuthorizeAsync(token, action, sensitivityLevel);
            }
            catch (Exception ex)
            {
                return new AccessDecision
                {
                    IsAllowed = false,
                    Reason = "Authorization failed: " + ex.Message
                };
            }
        }

        // Explicit-token flow: best for admin login/dashboard actions
        public async Task<AccessDecision> AuthorizeAsync(
            string token,
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            try
            {
                Logger.LogZeroTrustAuthorizationStarted(action.ToString());

                if (string.IsNullOrWhiteSpace(token))
                {
                    Logger.LogTokenValidationFailed("Access token is missing.");

                    return new AccessDecision
                    {
                        IsAllowed = false,
                        Reason = "Access token is missing."
                    };
                }

                Logger.LogTokenValidationStarted();

                Logger.LogUserInfoFetchStarted();
                var userInfo = await _keycloakService.GetUserInfoWithUserTokenAsync(token);

                if (userInfo == null)
                {
                    Logger.LogTokenValidationFailed("User info could not be retrieved from token.");

                    return new AccessDecision
                    {
                        IsAllowed = false,
                        Reason = "Token is invalid or user info could not be retrieved."
                    };
                }

                Logger.LogTokenValidationSucceeded();
                Logger.LogUserInfoFetchSucceeded(userInfo.preferred_username);

                var claims = _claimsExtractor.Extract(userInfo);

                Logger.LogClaimsExtracted(
                    claims.Username,
                    claims.Email,
                    claims.Department,
                    claims.ClearanceLevel,
                    claims.Roles == null ? "" : string.Join(",", claims.Roles));

                Logger.LogPolicyEvaluationStarted(action.ToString());

                var decision = _policyEngine.Evaluate(claims, action, sensitivityLevel);

                Logger.LogPolicyEvaluationResult(
                    action.ToString(),
                    decision.IsAllowed,
                    decision.Reason);

                if (!decision.IsAllowed)
                {
                    Logger.LogUnauthorizedAccessAttempt(
                        claims.Username,
                        action.ToString(),
                        decision.Reason);

                    await _logger.LogDeniedActionAsync(
                        claims.Username,
                        action.ToString(),
                        decision.Reason);
                }
                else
                {
                    Logger.LogZeroTrustAuthorizationSucceeded(
                        action.ToString(),
                        claims.Username,
                        decision.Reason);
                }

                return decision;
            }
            catch (Exception ex)
            {
                Logger.LogZeroTrustAuthorizationFailed(action.ToString(), ex);

                return new AccessDecision
                {
                    IsAllowed = false,
                    Reason = "Authorization failed: " + ex.Message
                };
            }
        }
    }
}