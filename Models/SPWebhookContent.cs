namespace SPO.Webhook.Models
{
    public class SPWebhookContent
    {
        public List<SPWebhookNotification> Value { get; set; } = new List<SPWebhookNotification>();
    }
}
