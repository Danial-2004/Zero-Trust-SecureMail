using System;
using System.IO;
using Newtonsoft.Json;
using SecureMailPolicyApi.Models;

namespace SecureMailPolicyApi.Services
{
    public class FilePolicyService
    {
        private readonly string _policyFile;
        private readonly object _sync = new object();

        public FilePolicyService()
        {
            string dataFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            Directory.CreateDirectory(dataFolder);

            _policyFile = Path.Combine(dataFolder, "policy.json");
        }

        public PolicyResponse GetPolicy()
        {
            lock (_sync)
            {
                if (!File.Exists(_policyFile))
                    throw new FileNotFoundException("Policy file not found.", _policyFile);

                string json = File.ReadAllText(_policyFile);

                if (string.IsNullOrWhiteSpace(json))
                    throw new InvalidOperationException("Policy file is empty.");

                var policy = JsonConvert.DeserializeObject<PolicyResponse>(json);

                if (policy == null)
                    throw new InvalidOperationException("Failed to parse policy file.");

                return policy;
            }
        }

        public void SavePolicy(PolicyResponse policy)
        {
            if (policy == null)
                throw new ArgumentNullException(nameof(policy));

            lock (_sync)
            {
                string json = JsonConvert.SerializeObject(policy, Formatting.Indented);
                AtomicWrite(_policyFile, json);
            }
        }

        private static void AtomicWrite(string destinationPath, string content)
        {
            string? folder = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(folder))
                Directory.CreateDirectory(folder);

            string temporaryPath = Path.Combine(
                folder ?? string.Empty,
                Path.GetFileName(destinationPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                File.WriteAllText(temporaryPath, content);
                File.Move(temporaryPath, destinationPath, true);
            }
            finally
            {
                if (File.Exists(temporaryPath))
                    File.Delete(temporaryPath);
            }
        }
    }
}
