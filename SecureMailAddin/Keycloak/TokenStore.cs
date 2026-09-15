using System;
using System.IO;
using Newtonsoft.Json;

namespace SecureMailAddin.Keycloak
{
    public static class TokenStore
    {
        private static readonly string FolderPath =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SecureMailAddin");

        private static readonly string FilePath =
            Path.Combine(FolderPath, "token_store.json");

        public static void Save(StoredTokenInfo token)
        {
            Directory.CreateDirectory(FolderPath);
            string json = JsonConvert.SerializeObject(token, Formatting.Indented);
            File.WriteAllText(FilePath, json);
        }

        public static StoredTokenInfo Load()
        {
            if (!File.Exists(FilePath))
                return null;

            string json = File.ReadAllText(FilePath);
            return JsonConvert.DeserializeObject<StoredTokenInfo>(json);
        }

        public static void Clear()
        {
            if (File.Exists(FilePath))
                File.Delete(FilePath);
        }
    }
}