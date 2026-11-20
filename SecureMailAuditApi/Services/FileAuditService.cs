using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using SecureMailAuditApi.Models;

namespace SecureMailAuditApi.Services
{
    public class FileAuditService
    {
        private readonly string _dataFolder;
        private readonly string _auditLogFile;
        private readonly string _mailEventFile;
        private readonly object _sync = new object();

        public FileAuditService()
        {
            _dataFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            _auditLogFile = Path.Combine(_dataFolder, "audit_logs.json");
            _mailEventFile = Path.Combine(_dataFolder, "mail_events.json");

            Directory.CreateDirectory(_dataFolder);
        }

        public void SaveAuditLog(AuditLogRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            lock (_sync)
            {
                var logs = ReadList<AuditLogRequest>(_auditLogFile);
                logs.Add(request);
                WriteList(_auditLogFile, logs);
            }
        }

        public void SaveMailEvent(MailClassificationRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            lock (_sync)
            {
                var events = ReadList<MailClassificationRequest>(_mailEventFile);
                events.Add(request);
                WriteList(_mailEventFile, events);
            }
        }

        public List<AuditLogRequest> GetAuditLogs()
        {
            lock (_sync)
            {
                return ReadList<AuditLogRequest>(_auditLogFile);
            }
        }

        public List<MailClassificationRequest> GetMailEvents()
        {
            lock (_sync)
            {
                return ReadList<MailClassificationRequest>(_mailEventFile);
            }
        }

        private List<T> ReadList<T>(string filePath)
        {
            if (!File.Exists(filePath))
                return new List<T>();

            string json = File.ReadAllText(filePath);

            if (string.IsNullOrWhiteSpace(json))
                return new List<T>();

            return JsonConvert.DeserializeObject<List<T>>(json) ?? new List<T>();
        }

        private static void WriteList<T>(string filePath, List<T> data)
        {
            string json = JsonConvert.SerializeObject(data, Formatting.Indented);
            AtomicWrite(filePath, json);
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
