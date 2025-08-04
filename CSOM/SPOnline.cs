using Microsoft.AspNetCore.Identity.Data;
using Microsoft.SharePoint.Client;
using PnP.Core.Model.SharePoint;
using PnP.Framework;
using System.Security;
using System.Security.Cryptography.X509Certificates;

namespace SPO.Webhook.CSOM
{
    public class SPOnline
    {
        string clientId = "48a8902f-4815-4e08-8a99-271788392924";
        string tenantId = "32f5eede-83de-456a-bb71-e145353d895a";
        string redirectUrl = "http://localhost";
        string siteUrl = "https://nagp365.sharepoint.com/sites/developer";

        public async Task GetChanges()
        {
            // CreateWithInteractiveLogin - Login using browser
            //var authManager = AuthenticationManager.CreateWithInteractiveLogin(clientId, redirectUrl, tenantId);

            //CreateWithCredentials - Will not work if MFA is enabled with account
            //var authManager = AuthenticationManager.CreateWithCredentials(clientId, "ateet@nagp365.onmicrosoft.com", ConvertToSecureString(""), redirectUrl);

            var basePath = AppContext.BaseDirectory;
            var certPath = Path.Combine(basePath, "Certificate", "SharePoint-Connect-From-AD-App.pfx");
            var authManager = AuthenticationManager.CreateWithCertificate(clientId, certPath, "", tenantId);

            using (var context = await authManager.GetContextAsync(siteUrl))
            {
                // Read web properties
                var web = context.Web;
                context.Load(web, w => w.Id, w => w.Title);
                await context.ExecuteQueryRetryAsync();

                Console.WriteLine($"{web.Id} - {web.Title}");

                // Retrieve a list by title together with selected properties
                var rerList = web.GetListByTitle("RER", l => l.Id, l => l.Title);
                context.Load(rerList);
                await context.ExecuteQueryRetryAsync();
                Console.WriteLine($"{rerList.Id} - {rerList.Title}");

                //ChangeQuery helps the method gets the changes by setting appropriate property as true  
                ChangeQuery cq = new ChangeQuery();
                cq.Item = true;
                cq.Add = true;
                cq.Update = true;
                cq.DeleteObject = true;

                ChangeCollection changes = rerList.GetChanges(cq);
                // Store the change token received with the last change
                context.Load(changes);
                context.ExecuteQuery();
                var lastChangetoken = changes.Last().ChangeToken;
                Console.WriteLine(changes.Count.ToString());
                foreach (var change in changes)
                {
                    //Console.WriteLine($"Change item id: {change.ChangeToken.}");
                    if (change is IChangeItem changeItem)
                    {
                        DateTime changeTime = changeItem.Time;

                        if (changeItem.IsPropertyAvailable<IChangeItem>(p => p.ListId))
                        {
                            // do something with the returned list id
                        }

                        if (changeItem.ChangeType == PnP.Core.Model.SharePoint.ChangeType.Add)
                        {
                            // do something "add" specific
                        }
                    }
                    if (change is Microsoft.SharePoint.Client.ChangeItem spChangeItem)
                    {
                        Console.WriteLine($"Change item id: {spChangeItem.ItemId}");
                    }
                    else if (change is IChangeField changeField)
                    {
                        // process the change
                    }
                }
            }
        }

        public async Task GetSPListItems()
        {
            // Create an instance of the AuthenticationManager type
            var authManager = AuthenticationManager.CreateWithInteractiveLogin(clientId, redirectUrl, tenantId);
            //var authManager = AuthenticationManager.CreateWithCredentials(clientId, "ateet@nagp365.onmicrosoft.com", System.Security.SecureString("Nagarro@04012025"), redirectUrl);
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

        public SecureString ConvertToSecureString(string password)
        {
            if (string.IsNullOrEmpty(password))
                throw new ArgumentNullException(nameof(password));

            SecureString secure = new SecureString();
            foreach (char c in password)
                secure.AppendChar(c);

            secure.MakeReadOnly();
            return secure;
        }

    }
}
