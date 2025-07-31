using Microsoft.AspNetCore.Identity.Data;
using Microsoft.SharePoint.Client;
using PnP.Framework;

namespace SPO.Webhook.CSOM
{
    public class GetChanges
    {
        string clientId = "48a8902f-4815-4e08-8a99-271788392924";
        string tenantId = "32f5eede-83de-456a-bb71-e145353d895a";
        string redirectUrl = "http://localhost";
        string siteUrl = "https://nagp365.sharepoint.com/sites/developer";

        public async Task Login()
        {
            // Create an instance of the AuthenticationManager type
            var authManager = AuthenticationManager.CreateWithInteractiveLogin(clientId, redirectUrl, tenantId);
            // Get a reference to the ClientContext of CSOM
            using (var context = await authManager.GetContextAsync(siteUrl))
            {
                // Read web properties
                var web = context.Web;
                context.Load(web, w => w.Id, w => w.Title);
                await context.ExecuteQueryRetryAsync();

                Console.WriteLine($"{web.Id} - {web.Title}");

                // Retrieve a list by title together with selected properties
                var documents = web.GetListByTitle("RER", l => l.Id, l => l.Title);

                Console.WriteLine($"{documents.Id} - {documents.Title}");

                // Retrieve the top 10 items from the list
                var query = CamlQuery.CreateAllItemsQuery(10);
                var items = documents.GetItems(query);
                context.Load(items);
                await context.ExecuteQueryRetryAsync();

                // Browse through all the items
                foreach (var i in items)
                {
                    Console.WriteLine($"{i.Id} - {i["Title"]}");
                }
            }
        }
    }
}
