using System;
using System.Collections.Generic;
using System.Linq;

namespace SecureMailAddin
{
    public static class PendingUsersRequests
    {
        private static readonly object _lock = new object();
        private static List<PendingUserRequest> _pendingUsers = new List<PendingUserRequest>();

        // Returns a copy of the pending users
        public static List<PendingUserRequest> PendingUsers
        {
            get
            {
                lock (_lock)
                {
                    return new List<PendingUserRequest>(_pendingUsers);
                }
            }
        }

        // Add a new request safely; returns false if duplicate
        public static bool AddPendingUser(PendingUserRequest request)
        {
            lock (_lock)
            {
                // Check for existing username OR email
                if (_pendingUsers.Any(u =>
                    u.Username.Equals(request.Username, StringComparison.OrdinalIgnoreCase) ||
                    u.Email.Equals(request.Email, StringComparison.OrdinalIgnoreCase)))
                {
                    return false; // duplicate exists, do not add
                }

                _pendingUsers.Add(request);
                return true;
            }
        }

        // Remove a request by Id
        public static void RemovePendingUser(string id)
        {
            lock (_lock)
            {
                _pendingUsers.RemoveAll(u => u.Id == id);
            }
        }

        // Optional: check if email or username already exists in pending requests
        public static bool Exists(string username, string email)
        {
            lock (_lock)
            {
                return _pendingUsers.Any(u =>
                    u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) ||
                    u.Email.Equals(email, StringComparison.OrdinalIgnoreCase));
            }
        }
    }

    public class PendingUserRequest
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Username { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public bool EmailVerified { get; set; }
    }
}
