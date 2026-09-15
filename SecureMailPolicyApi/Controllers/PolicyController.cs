using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SecureMailPolicyApi.Models;
using SecureMailPolicyApi.Services;

namespace SecureMailPolicyApi.Controllers
{
    [ApiController]
    [Route("api/policy")]
    public class PolicyController : ControllerBase
    {
        private readonly FilePolicyService _filePolicyService;

        public PolicyController(FilePolicyService filePolicyService)
        {
            _filePolicyService = filePolicyService;
        }

        [HttpGet("current")]
        [Authorize]
        public IActionResult GetCurrent()
        {
            var policy = _filePolicyService.GetPolicy();
            return Ok(policy);
        }

        [HttpPost("current")]
        [Authorize]
        public IActionResult UpdateCurrent([FromBody] PolicyResponse policy)
        {
            if (policy == null)
                return BadRequest("Policy body is missing.");

            _filePolicyService.SavePolicy(policy);
            return Ok(new { message = "Policy updated successfully." });
        }
    }
}