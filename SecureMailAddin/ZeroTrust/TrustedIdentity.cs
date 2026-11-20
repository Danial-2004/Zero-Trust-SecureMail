using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SecureMailAddin.ZeroTrust
{
    /// <summary>
    /// A non-forgeable, in-process proof that cryptographic token validation
    /// succeeded. It has no public constructor and contains no client-writable
    /// Boolean. It also binds the trusted claims with a SHA-256 digest.
    /// </summary>
    public sealed class TrustedAuthenticationProof
    {
        private TrustedAuthenticationProof(
            string validatorId,
            DateTime validatedAtUtc,
            string tokenId,
            string subject,
            string claimsDigest)
        {
            ValidationId = Guid.NewGuid().ToString("N");
            ValidatorId = validatorId ?? string.Empty;
            ValidatedAtUtc = validatedAtUtc;
            TokenId = tokenId ?? string.Empty;
            Subject = subject ?? string.Empty;
            ClaimsDigest = claimsDigest ?? string.Empty;
        }

        public string ValidationId { get; private set; }
        public string ValidatorId { get; private set; }
        public DateTime ValidatedAtUtc { get; private set; }
        public string TokenId { get; private set; }
        public string Subject { get; private set; }
        public string ClaimsDigest { get; private set; }

        internal static TrustedAuthenticationProof Issue(
            ValidatedTokenIdentity identity,
            string validatorId)
        {
            if (identity == null)
                throw new ArgumentNullException(nameof(identity));

            string digest = ComputeClaimsDigest(
                identity.Subject,
                identity.Username,
                identity.Email,
                identity.Department,
                identity.ClearanceLevel,
                identity.Roles);

            return new TrustedAuthenticationProof(
                validatorId,
                DateTime.UtcNow,
                identity.TokenId,
                identity.Subject,
                digest);
        }

        /// <summary>
        /// Available only to code compiled into this assembly (including the
        /// linked-source evaluation projects). Production callers outside the
        /// assembly cannot fabricate a proof.
        /// </summary>
        internal static TrustedAuthenticationProof ForUnitTests(UserClaimsProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            string digest = ComputeClaimsDigest(
                profile.Subject,
                profile.Username,
                profile.Email,
                profile.Department,
                profile.ClearanceLevel,
                profile.Roles);

            return new TrustedAuthenticationProof(
                "unit-test",
                DateTime.UtcNow,
                "unit-test-token",
                profile.Subject,
                digest);
        }

        internal static string ComputeClaimsDigest(
            string subject,
            string username,
            string email,
            string department,
            string clearance,
            IEnumerable<string> roles)
        {
            string canonicalRoles = string.Join(
                "\u001F",
                (roles ?? new string[0]).Select(r => r ?? string.Empty).ToArray());

            string canonical =
                (subject ?? string.Empty) + "\u001E" +
                (username ?? string.Empty) + "\u001E" +
                (email ?? string.Empty) + "\u001E" +
                (department ?? string.Empty) + "\u001E" +
                (clearance ?? string.Empty) + "\u001E" +
                canonicalRoles;

            using (var sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(canonical);
                return Convert.ToBase64String(sha.ComputeHash(bytes));
            }
        }

        internal bool Matches(UserClaimsProfile profile)
        {
            if (profile == null)
                return false;

            string actual = ComputeClaimsDigest(
                profile.Subject,
                profile.Username,
                profile.Email,
                profile.Department,
                profile.ClearanceLevel,
                profile.Roles);

            return FixedTimeEquals(ClaimsDigest, actual) &&
                   string.Equals(Subject, profile.Subject ?? string.Empty, StringComparison.Ordinal);
        }

        internal static bool FixedTimeEquals(string expected, string actual)
        {
            if (expected == null || actual == null)
                return false;

            byte[] a = Encoding.UTF8.GetBytes(expected);
            byte[] b = Encoding.UTF8.GetBytes(actual);

            if (a.Length != b.Length)
                return false;

            int difference = 0;
            for (int i = 0; i < a.Length; i++)
                difference |= a[i] ^ b[i];

            return difference == 0;
        }
    }

    /// <summary>
    /// Identity produced only after successful JWT validation.
    /// </summary>
    public sealed class ValidatedTokenIdentity
    {
        internal ValidatedTokenIdentity(
            TrustedAuthenticationProof proof,
            string subject,
            string username,
            string email,
            string department,
            string clearanceLevel,
            IEnumerable<string> roles,
            string issuer,
            string audience,
            string tokenId,
            DateTime expiresUtc)
        {
            Proof = proof;
            Subject = subject ?? string.Empty;
            Username = username ?? string.Empty;
            Email = email ?? string.Empty;
            Department = department ?? string.Empty;
            ClearanceLevel = clearanceLevel ?? string.Empty;
            Roles = new ReadOnlyCollection<string>(
                (roles ?? new string[0]).Where(r => !string.IsNullOrWhiteSpace(r)).ToList());
            Issuer = issuer ?? string.Empty;
            Audience = audience ?? string.Empty;
            TokenId = tokenId ?? string.Empty;
            ExpiresUtc = expiresUtc;
        }

        public TrustedAuthenticationProof Proof { get; private set; }
        public string Subject { get; private set; }
        public string Username { get; private set; }
        public string Email { get; private set; }
        public string Department { get; private set; }
        public string ClearanceLevel { get; private set; }
        public IList<string> Roles { get; private set; }
        public string Issuer { get; private set; }
        public string Audience { get; private set; }
        public string TokenId { get; private set; }
        public DateTime ExpiresUtc { get; private set; }
    }

    public sealed class TokenValidationOutcome
    {
        private TokenValidationOutcome(
            JwtValidationStatus status,
            string policyCode,
            string reason,
            ValidatedTokenIdentity identity)
        {
            Status = status;
            PolicyCode = policyCode ?? PolicyCodes.JwtInvalid;
            Reason = reason ?? string.Empty;
            Identity = identity;
        }

        public JwtValidationStatus Status { get; private set; }
        public string PolicyCode { get; private set; }
        public string Reason { get; private set; }
        public ValidatedTokenIdentity Identity { get; private set; }

        public bool Succeeded
        {
            get { return Status == JwtValidationStatus.Valid && Identity != null && Identity.Proof != null; }
        }

        internal static TokenValidationOutcome Success(ValidatedTokenIdentity identity)
        {
            return new TokenValidationOutcome(
                JwtValidationStatus.Valid,
                PolicyCodes.PolicySatisfied,
                "JWT signature, issuer, audience, lifetime and required claims were validated.",
                identity);
        }

        public static TokenValidationOutcome Failure(
            JwtValidationStatus status,
            string policyCode,
            string reason)
        {
            return new TokenValidationOutcome(status, policyCode, reason, null);
        }
    }

    public interface IJwtTokenValidator
    {
        TokenValidationOutcome Validate(string token);
    }

    public interface IReplayCache
    {
        bool TryRegister(string tokenId, DateTime expiresUtc);
    }

    /// <summary>
    /// Process-local replay cache. Use a shared persistent cache in production
    /// if multiple add-in/API processes must reject the same jti.
    /// </summary>
    public sealed class InMemoryReplayCache : IReplayCache
    {
        private readonly ConcurrentDictionary<string, DateTime> _entries =
            new ConcurrentDictionary<string, DateTime>(StringComparer.Ordinal);

        public bool TryRegister(string tokenId, DateTime expiresUtc)
        {
            DateTime now = DateTime.UtcNow;

            if (string.IsNullOrWhiteSpace(tokenId) || expiresUtc <= now)
                return false;

            PurgeExpired(now);

            DateTime existing;
            if (_entries.TryGetValue(tokenId, out existing))
            {
                if (existing > now)
                    return false;

                DateTime ignored;
                _entries.TryRemove(tokenId, out ignored);
            }

            return _entries.TryAdd(tokenId, expiresUtc);
        }

        private void PurgeExpired(DateTime now)
        {
            foreach (var entry in _entries.ToArray())
            {
                if (entry.Value <= now)
                {
                    DateTime ignored;
                    _entries.TryRemove(entry.Key, out ignored);
                }
            }
        }
    }

    /// <summary>
    /// Thread-safe, bounded, TTL-expiring cache for request-level replay
    /// protection. Stores unique request identifiers (not JWT jti values).
    /// An OAuth bearer token is reusable until expiry; individual protected
    /// operations are identified by a unique request ID. Duplicate request
    /// IDs within the TTL window are rejected. Entries expire at the
    /// configured TTL so the cache cannot grow indefinitely.
    /// </summary>
    public sealed class RequestReplayCache
    {
        private readonly ConcurrentDictionary<string, DateTime> _entries =
            new ConcurrentDictionary<string, DateTime>(StringComparer.Ordinal);
        private readonly int _maxEntries;

        public RequestReplayCache(int maxEntries = 100000)
        {
            _maxEntries = maxEntries > 0 ? maxEntries : 100000;
        }

        /// <summary>
        /// Registers a request ID. Returns true if the ID is new (first
        /// submission), false if it is a duplicate within the TTL window.
        /// Fail-closed: when the cache is at capacity, new IDs are rejected.
        /// </summary>
        public bool TryRegisterRequest(string requestId, DateTime expiresUtc)
        {
            DateTime now = DateTime.UtcNow;

            if (string.IsNullOrWhiteSpace(requestId) || expiresUtc <= now)
                return false;

            PurgeExpired(now);

            DateTime existing;
            if (_entries.TryGetValue(requestId, out existing))
            {
                if (existing > now)
                    return false; // Duplicate request within TTL

                DateTime ignored;
                _entries.TryRemove(requestId, out ignored);
            }

            // Fail-closed when at capacity: reject rather than silently
            // evicting, which could allow a replay to slip through.
            if (_entries.Count >= _maxEntries && !_entries.ContainsKey(requestId))
                return false;

            return _entries.TryAdd(requestId, expiresUtc);
        }

        private void PurgeExpired(DateTime now)
        {
            foreach (var entry in _entries.ToArray())
            {
                if (entry.Value <= now)
                {
                    DateTime ignored;
                    _entries.TryRemove(entry.Key, out ignored);
                }
            }
        }
    }
}
