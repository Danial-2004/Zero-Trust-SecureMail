using System;
using System.IO;
using Newtonsoft.Json;

namespace SecureMailAddin.Keycloak
{
    public static class PolicyCacheService
    {
        private static readonly string FolderPath =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SecureMailAddin");

        private static readonly string FilePath =
            Path.Combine(FolderPath, "policy_cache.json");

        public static void Save(PolicyResponse policy)
        {
            Directory.CreateDirectory(FolderPath);
            string json = JsonConvert.SerializeObject(policy, Formatting.Indented);
            File.WriteAllText(FilePath, json);
        }

        public static PolicyResponse Load()
        {
            if (!File.Exists(FilePath))
                return null;

            string json = File.ReadAllText(FilePath);
            return JsonConvert.DeserializeObject<PolicyResponse>(json);
        }
    }
}