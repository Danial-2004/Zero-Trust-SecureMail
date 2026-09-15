using System;
using System.IO;
using Newtonsoft.Json;
using SecureMailPolicyApi.Models;

namespace SecureMailPolicyApi.Services
{
    public class FilePolicyService
    {
        private readonly string _policyFile;

        public FilePolicyService()
        {
            string dataFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            Directory.CreateDirectory(dataFolder);

            _policyFile = Path.Combine(dataFolder, "policy.json");
        }

        public PolicyResponse GetPolicy()
        {
            if (!File.Exists(_policyFile))
                throw new FileNotFoundException("Policy file not found.", _policyFile);

            string json = File.ReadAllText(_policyFile);

            if (string.IsNullOrWhiteSpace(json))
                throw new Exception("Policy file is empty.");

            var policy = JsonConvert.DeserializeObject<PolicyResponse>(json);

            if (policy == null)
                throw new Exception("Failed to parse policy file.");

            return policy;
        }

        public void SavePolicy(PolicyResponse policy)
        {
            string json = JsonConvert.SerializeObject(policy, Formatting.Indented);
            File.WriteAllText(_policyFile, json);
        }
    }
}