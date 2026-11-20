using SecureMailAddin.MailHandler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SecureMailAddin.ZeroTrust
{
    public class RecipientClearanceService
    {
        public MailSensitivityLevel GetRecipientClearance(string email)
        {
            // TODO:
            // Replace this with your API / Keycloak / Database lookup.

            var clearances = new Dictionary<string, MailSensitivityLevel>
            {
                { "alice@company.com", MailSensitivityLevel.Restricted },
                { "bob@company.com", MailSensitivityLevel.Confidential },
                { "john@company.com", MailSensitivityLevel.Internal }
            };

            if (clearances.ContainsKey(email.ToLower()))
                return clearances[email.ToLower()];

            return MailSensitivityLevel.Public;
        }
    }
}
