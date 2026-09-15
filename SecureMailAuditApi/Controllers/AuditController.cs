using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SecureMailAuditApi.Models;
using SecureMailAuditApi.Services;

namespace SecureMailAuditApi.Controllers
{
    [ApiController]
    [Route("api/audit")]
    public class AuditController : ControllerBase
    {
        private readonly FileAuditService _fileAuditService;

        public AuditController(FileAuditService fileAuditService)
        {
            _fileAuditService = fileAuditService;
        }

        [HttpPost("log")]
        [Authorize]
        public IActionResult Log([FromBody] AuditLogRequest request)
        {
            if (request == null)
                return BadRequest("Request body is missing.");

            _fileAuditService.SaveAuditLog(request);

            return Ok(new
            {
                message = "Audit log stored successfully."
            });
        }

        [HttpPost("mail-event")]
        [Authorize]
        public IActionResult MailEvent([FromBody] MailClassificationRequest request)
        {
            if (request == null)
                return BadRequest("Request body is missing.");

            _fileAuditService.SaveMailEvent(request);

            return Ok(new
            {
                message = "Mail classification event stored successfully."
            });
        }

        [HttpGet("logs")]
        [Authorize]
        public IActionResult GetLogs()
        {
            var logs = _fileAuditService.GetAuditLogs();
            return Ok(logs);
        }

        [HttpGet("mail-events")]
        [Authorize]
        public IActionResult GetMailEvents()
        {
            var events = _fileAuditService.GetMailEvents();
            return Ok(events);
        }
    }
}