using System;

namespace SecureMailAddin.ZeroTrust
{
    /// <summary>
    /// Token-to-decision pipeline used by the add-in and by the isolated C#
    /// evaluation harness. It never treats a user-info response or a Boolean as
    /// authentication evidence.
    /// </summary>
    public sealed class ZeroTrustAuthorizer
    {
        private readonly IJwtTokenValidator _validator;
        private readonly ClaimsExtractor _claimsExtractor;
        private readonly AccessPolicyEngine _policyEngine;

        public ZeroTrustAuthorizer()
            : this(CreateDefaultValidator())
        {
        }

        public ZeroTrustAuthorizer(IJwtTokenValidator validator)
        {
            if (validator == null)
                throw new ArgumentNullException(nameof(validator));

            _validator = validator;
            _claimsExtractor = new ClaimsExtractor();
            _policyEngine = new AccessPolicyEngine();
        }

        public static ZeroTrustAuthorizer CreateDefault()
        {
            return new ZeroTrustAuthorizer(CreateDefaultValidator());
        }

        private static IJwtTokenValidator CreateDefaultValidator()
        {
            JwtValidationOptions options;
            string configurationError;

            if (!JwtValidationOptions.TryFromEnvironment(out options, out configurationError))
                return new ConfigurationFailureValidator(configurationError);

            return new JwtTokenValidator(options);
        }

        public AccessDecision Authorize(
            string token,
            ProtectedActionType action,
            string sensitivityLevel = null)
        {
            return AuthorizeValidated(token, action.ToString(), sensitivityLevel,
                (profile, actionName, classification) =>
                    _policyEngine.Evaluate(profile, action, classification));
        }

        public AccessDecision Authorize(
            string token,
            string actionName,
            string sensitivityLevel = null)
        {
            return AuthorizeValidated(token, actionName, sensitivityLevel,
                (profile, resolvedActionName, classification) =>
                    _policyEngine.Evaluate(profile, resolvedActionName, classification));
        }

        public AccessDecision AuthorizeMailSend(
            string token,
            string classificationTag)
        {
            return AuthorizeValidated(token, classificationTag, classificationTag,
                (profile, actionName, classification) =>
                    _policyEngine.EvaluateMailSend(profile, classification));
        }

        private AccessDecision AuthorizeValidated(
            string token,
            string actionName,
            string sensitivityLevel,
            Func<UserClaimsProfile, string, string, AccessDecision> evaluate)
        {
            try
            {
                TokenValidationOutcome outcome = _validator.Validate(token);
                if (outcome == null)
                {
                    return AccessDecision.Deny(
                        PolicyCodes.AuthenticationServiceUnavailable,
                        "Token validator returned no result; failing closed.",
                        actionName,
                        null,
                        null,
                        sensitivityLevel);
                }

                if (!outcome.Succeeded || outcome.Identity == null)
                {
                    return AccessDecision.Deny(
                        outcome.PolicyCode,
                        outcome.Reason,
                        actionName,
                        null,
                        null,
                        sensitivityLevel);
                }

                UserClaimsProfile profile = _claimsExtractor.Extract(outcome.Identity);
                return evaluate(profile, actionName, sensitivityLevel);
            }
            catch (Exception)
            {
                return AccessDecision.Deny(
                    PolicyCodes.PolicyEngineError,
                    "An unexpected authorization error occurred; failing closed.",
                    actionName,
                    null,
                    null,
                    sensitivityLevel);
            }
        }

        private sealed class ConfigurationFailureValidator : IJwtTokenValidator
        {
            private readonly string _reason;

            public ConfigurationFailureValidator(string reason)
            {
                _reason = string.IsNullOrWhiteSpace(reason)
                    ? "JWT validation is not configured."
                    : reason;
            }

            public TokenValidationOutcome Validate(string token)
            {
                return TokenValidationOutcome.Failure(
                    JwtValidationStatus.ValidationUnavailable,
                    PolicyCodes.AuthenticationConfigurationMissing,
                    _reason);
            }
        }
    }
}
