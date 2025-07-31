namespace SPO.Webhook.Models
{
    public class SPWebhookNotification
    {
        public string SubscriptionId { get; set; }=string.Empty;
        public string ClientState { get; set; } = string.Empty;
        public string ExpirationDateTime { get; set; } = string.Empty;
        public string Resource { get; set; } = string.Empty;
        public string TenantId { get; set; } = string.Empty;
        public string SiteUrl { get; set; } = string.Empty;
        public string WebId { get; set; } = string.Empty;
    }
}
