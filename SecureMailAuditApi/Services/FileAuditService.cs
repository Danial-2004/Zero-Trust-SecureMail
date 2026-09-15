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

        public FileAuditService()
        {
            _dataFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            _auditLogFile = Path.Combine(_dataFolder, "audit_logs.json");
            _mailEventFile = Path.Combine(_dataFolder, "mail_events.json");

            Directory.CreateDirectory(_dataFolder);
        }

        public void SaveAuditLog(AuditLogRequest request)
        {
            var logs = ReadList<AuditLogRequest>(_auditLogFile);
            logs.Add(request);
            WriteList(_auditLogFile, logs);
        }

        public void SaveMailEvent(MailClassificationRequest request)
        {
            var events = ReadList<MailClassificationRequest>(_mailEventFile);
            events.Add(request);
            WriteList(_mailEventFile, events);
        }

        public List<AuditLogRequest> GetAuditLogs()
        {
            return ReadList<AuditLogRequest>(_auditLogFile);
        }

        public List<MailClassificationRequest> GetMailEvents()
        {
            return ReadList<MailClassificationRequest>(_mailEventFile);
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

        private void WriteList<T>(string filePath, List<T> data)
        {
            string json = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }
    }
}