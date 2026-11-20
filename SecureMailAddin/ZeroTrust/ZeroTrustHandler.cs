using System;
using System.Threading.Tasks;
using SecureMailAddin.Keycloak;

namespace SecureMailAddin.ZeroTrust
{
    public class ZeroTrustHandler
    {
        private readonly KeycloakService _keycloakService;
        private readonly ZeroTrustAuthorizer _authorizer;
        private readonly UnauthorizedAccessLogger _logger;

        // Shared, stable authorizer so the JWT replay cache is not reset on
        // every operation. JWTs remain reusable until expiry.
        private static readonly ZeroTrustAuthorizer _sharedAuthorizer =
            ZeroTrustAuthorizer.CreateDefault();

        // Shared, stable request-replay cache so duplicate request IDs are
        // detected across operations within the same process. Entries expire
        // at TTL so the cache cannot grow indefinitely.
        private static readonly RequestReplayCache _requestReplayCache =
            new RequestReplayCache();

        // Bounded TTL for request-replay entries (5 minutes).
        private static readonly TimeSpan RequestReplayTtl = TimeSpan.FromMinutes(5);

        public ZeroTrustHandler()
        {
            _keycloakService = new KeycloakService();
            _authorizer = _sharedAuthorizer;
            _logger = new UnauthorizedAccessLogger();
        }

        internal ZeroTrustHandler(
            KeycloakService keycloakService,
            ZeroTrustAuthorizer authorizer,
            UnauthorizedAccessLogger logger)
        {
            _keycloakService = keycloakService;
            _authorizer = authorizer ?? ZeroTrustAuthorizer.CreateDefault();
            _logger = logger ?? new UnauthorizedAccessLogger();
        }

        /// <summary>
        /// Existing flow: obtain the current token, then validate it locally
        /// before any authorization decision.
        /// </summary>
        public async Task<AccessDecision> AuthorizeAsync(
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            try
            {
                string token = await _keycloakService.GetValidAccessTokenAsync();

                if (string.IsNullOrWhiteSpace(token))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.JwtMissing,
                        "Access token is missing.",
                        action.ToString());
                }

                return await AuthorizeAsync(token, action, sensitivityLevel);
            }
            catch (Exception ex)
            {
                Logger.LogZeroTrustAuthorizationFailed(action.ToString(), ex);

                return AccessDecision.Deny(
                    PolicyCodes.AuthenticationServiceUnavailable,
                    "Authentication service is unavailable; failing closed.",
                    action.ToString());
            }
        }

        /// <summary>
        /// Explicit-token flow for admin login/dashboard actions. The token is
        /// cryptographically validated; the user-info endpoint is not treated as
        /// authentication evidence.
        /// </summary>
        public async Task<AccessDecision> AuthorizeAsync(
            string token,
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            string actionName = action.ToString();

            try
            {
                Logger.LogZeroTrustAuthorizationStarted(actionName);
                Logger.LogTokenValidationStarted();

                AccessDecision decision =
                    _authorizer.Authorize(token, action, sensitivityLevel);

                LogValidationAndPolicy(decision, actionName);
                await FinalizeDecisionAsync(decision);
                return decision;
            }
            catch (Exception ex)
            {
                Logger.LogZeroTrustAuthorizationFailed(actionName, ex);

                return AccessDecision.Deny(
                    PolicyCodes.PolicyEngineError,
                    "Authorization failed; failing closed.",
                    actionName,
                    null,
                    null,
                    sensitivityLevel);
            }
        }

        /// <summary>
        /// Strict mail-send flow: the classification tag is mapped to an action
        /// by the engine, preventing a downgrade to a lower-clearance action.
        /// JWTs remain reusable until expiry. Request-level replay protection
        /// is applied via the optional requestId: duplicate request IDs within
        /// the TTL window are rejected.
        /// </summary>
        public async Task<AccessDecision> AuthorizeMailSendAsync(string classificationTag, string requestId = null)
        {
            try
            {
                string token = await _keycloakService.GetValidAccessTokenAsync();

                if (string.IsNullOrWhiteSpace(token))
                {
                    return AccessDecision.Deny(
                        PolicyCodes.JwtMissing,
                        "Access token is missing.",
                        classificationTag ?? string.Empty);
                }

                return await AuthorizeMailSendAsync(token, classificationTag, requestId);
            }
            catch (Exception ex)
            {
                Logger.LogZeroTrustAuthorizationFailed("MailSend", ex);

                return AccessDecision.Deny(
                    PolicyCodes.AuthenticationServiceUnavailable,
                    "Authentication service is unavailable; failing closed.",
                    classificationTag ?? string.Empty);
            }
        }

        public async Task<AccessDecision> AuthorizeMailSendAsync(
            string token,
            string classificationTag,
            string requestId = null)
        {
            string actionName = "MailSend:" + (classificationTag ?? "UNKNOWN");

            try
            {
                Logger.LogZeroTrustAuthorizationStarted(actionName);
                Logger.LogTokenValidationStarted();

                // Layer 1: cryptographic JWT validation. The token remains
                // reusable until expiry; jti is not used for one-time use.
                AccessDecision decision =
                    _authorizer.AuthorizeMailSend(token, classificationTag);

                if (!decision.IsAllowed)
                {
                    LogValidationAndPolicy(decision, actionName);
                    await FinalizeDecisionAsync(decision);
                    return decision;
                }

                // Layer 2: request-level replay protection. Bind the request
                // ID to the authenticated identity and operation. Duplicate
                // request IDs within the TTL window are rejected.
                string effectiveRequestId = requestId ?? Guid.NewGuid().ToString("N");
                DateTime requestExpiry = DateTime.UtcNow.Add(RequestReplayTtl);

                if (!_requestReplayCache.TryRegisterRequest(effectiveRequestId, requestExpiry))
                {
                    decision = AccessDecision.Deny(
                        PolicyCodes.JwtReplayDetected,
                        "Duplicate request detected; request replay blocked.",
                        actionName,
                        null,
                        null,
                        classificationTag);
                    LogValidationAndPolicy(decision, actionName);
                    await FinalizeDecisionAsync(decision);
                    return decision;
                }

                LogValidationAndPolicy(decision, actionName);
                await FinalizeDecisionAsync(decision);
                return decision;
            }
            catch (Exception ex)
            {
                Logger.LogZeroTrustAuthorizationFailed(actionName, ex);

                return AccessDecision.Deny(
                    PolicyCodes.PolicyEngineError,
                    "Authorization failed; failing closed.",
                    actionName,
                    null,
                    null,
                    classificationTag);
            }
        }

        private void LogValidationAndPolicy(AccessDecision decision, string actionName)
        {
            if (decision == null)
                return;

            if (IsAuthenticationFailure(decision.PolicyCode))
                Logger.LogTokenValidationFailed(decision.Reason);
            else
                Logger.LogTokenValidationSucceeded();

            Logger.LogClaimsExtracted(
                string.IsNullOrWhiteSpace(decision.Username) ? "unknown" : decision.Username,
                string.Empty,
                string.Empty,
                decision.Clearance ?? string.Empty,
                decision.EffectiveRole ?? string.Empty);

            Logger.LogPolicyEvaluationStarted(actionName);
            Logger.LogPolicyEvaluationResult(
                actionName,
                decision.IsAllowed,
                decision.Reason);
        }

        private async Task FinalizeDecisionAsync(AccessDecision decision)
        {
            if (decision == null)
                return;

            if (!decision.IsAllowed)
            {
                Logger.LogUnauthorizedAccessAttempt(
                    string.IsNullOrWhiteSpace(decision.Username) ? "unknown" : decision.Username,
                    decision.Action,
                    decision.Reason);

                try
                {
                    await _logger.LogDeniedActionAsync(
                        string.IsNullOrWhiteSpace(decision.Username) ? "unknown" : decision.Username,
                        decision.Action,
                        decision.Reason + " PolicyCode=" + decision.PolicyCode);
                }
                catch (Exception ex)
                {
                    // Audit failure never changes a BLOCK into ALLOW.
                    Logger.LogError("AUDIT", "Audit service unavailable.", ex);
                    decision.PolicyCode = PolicyCodes.AuditServiceUnavailable;
                    decision.Reason += " Audit service unavailable.";
                }
            }
            else
            {
                Logger.LogZeroTrustAuthorizationSucceeded(
                    decision.Action,
                    decision.Username,
                    decision.Reason);
            }
        }

        private static bool IsAuthenticationFailure(string policyCode)
        {
            string code = policyCode ?? string.Empty;

            return code.StartsWith("JWT_", StringComparison.OrdinalIgnoreCase) ||
                   code.StartsWith("AUTH_", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(code, PolicyCodes.SubjectMissing, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(code, PolicyCodes.RoleMissing, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(code, PolicyCodes.ClearanceMissing, StringComparison.OrdinalIgnoreCase);
        }
    }
}
