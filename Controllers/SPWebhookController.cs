using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Microsoft.AspNetCore.WebUtilities;
using Newtonsoft.Json;
using SPO.Webhook.Models;
using System.Text;

namespace SPO.Webhook.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SPWebhookController : ControllerBase
    {
        private readonly ILogger<SPWebhookController> _logger;
        private readonly IConfiguration _config;

        public SPWebhookController(IConfiguration config, ILogger<SPWebhookController> logger)
        {
            _logger = logger;
            _config = config;
        }


        [HttpPost("handlerequest")]
        public async Task<IActionResult> HandleRequest()
        {
            string validationToken = string.Empty;
            string webhookClientState = _config.GetValue<string>("webhookclientstate", string.Empty)!;

            if (Request.Headers.TryGetValue("ClientState", out StringValues clientStateHeader))
            {
                string clientStateHeaderValue = clientStateHeader.FirstOrDefault() ?? string.Empty;

                if (!string.IsNullOrEmpty(clientStateHeaderValue) && clientStateHeaderValue.Equals(webhookClientState))
                {
                    _logger.LogInformation("Received client state: {ClientState}", clientStateHeaderValue);

                    var queryParams = QueryHelpers.ParseQuery(Request.QueryString.Value ?? string.Empty);

                    if (queryParams.TryGetValue("validationtoken", out var validationTokenValue))
                    {
                        validationToken = validationTokenValue.ToString();
                        _logger.LogInformation("Received validation token: {ValidationToken}", validationToken);
                        return Ok(validationToken); // Must echo back the token
                    }
                    else
                    {
                        try
                        {
                            if (Request.ContentLength > 0)
                            {
                                Request.EnableBuffering(); // Add this before reading the body
                                Request.Body.Position = 0;
                                using var reader = new StreamReader(Request.Body, Encoding.UTF8, leaveOpen: true);
                                var requestContent = await reader.ReadToEndAsync();

                                if (!string.IsNullOrEmpty(requestContent))
                                {
                                    try
                                    {
                                        var objNotification = JsonConvert.DeserializeObject<SPWebhookContent>(requestContent);
                                        var notification = objNotification?.Value?.FirstOrDefault();

                                        if (notification != null)
                                        {
                                            _ = Task.Run(() =>
                                            {
                                                _logger.LogInformation("Resource: {Resource}", notification.Resource);
                                                _logger.LogInformation("SubscriptionId: {SubscriptionId}", notification.SubscriptionId);
                                                _logger.LogInformation("TenantId: {TenantId}", notification.TenantId);
                                                _logger.LogInformation("SiteUrl: {SiteUrl}", notification.SiteUrl);
                                                _logger.LogInformation("WebId: {WebId}", notification.WebId);
                                                _logger.LogInformation("ExpirationDateTime: {Expiration}", notification.ExpirationDateTime);
                                            });

                                            return Ok(); // Webhook processed successfully
                                        }
                                    }
                                    catch (JsonException ex)
                                    {
                                        _logger.LogError(ex, "JSON deserialization error");
                                        return BadRequest("Invalid JSON");
                                    }
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            _logger.LogInformation(ex, $"{ex.Message}");
                        }
                    }
                }
                else
                {
                    _logger.LogWarning("ClientState mismatch or missing");
                    return Forbid();
                }
            }

            _logger.LogWarning("Missing ClientState header");
            return BadRequest("Missing ClientState header");
        }
    }
}
