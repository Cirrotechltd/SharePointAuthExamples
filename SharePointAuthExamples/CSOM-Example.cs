using System;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Identity.Client;

namespace SharePointCsomExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Azure AD and SharePoint Configuration
            string tenantId = "f4cbf802-fffb-4961-8633-9a240e5d234a";
            string clientId = "bdc37915-34aa-4ec6-8a86-1ce1363a09b1";
            string certificateThumbprint = "82062f3830d6ae387e878dc0cb31e616c78042f8";
            string siteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";

            try
            {
                // Load the certificate
                X509Certificate2 certificate = GetCertificateFromStore(certificateThumbprint);

                // Get access token
                string accessToken = GetAccessToken(tenantId, clientId, certificate).Result;

                // Initialize the SharePoint Client Context
                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                    };

                    while (true)
                    {
                        // Display menu
                        Console.WriteLine("\nMenu:");
                        Console.WriteLine("1. Get Site Title");
                        Console.WriteLine("2. Create a New List");
                        Console.WriteLine("3. Retrieve List Items");
                        Console.WriteLine("4. Exit");
                        Console.Write("Choose an option: ");

                        string choice = Console.ReadLine();

                        switch (choice)
                        {
                            case "1":
                                GetSiteTitle(context);
                                break;

                            case "2":
                                Console.Write("Enter the name of the new list: ");
                                string listName = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(listName))
                                {
                                    CreateList(context, listName);
                                }
                                else
                                {
                                    Console.WriteLine("List name cannot be empty.");
                                }
                                break;

                            case "3":
                                Console.Write("Enter the name of the list to retrieve items from: ");
                                string listToRetrieve = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(listToRetrieve))
                                {
                                    RetrieveListItems(context, listToRetrieve);
                                }
                                else
                                {
                                    Console.WriteLine("List name cannot be empty.");
                                }
                                break;

                            case "4":
                                Console.WriteLine("Exiting...");
                                return;

                            default:
                                Console.WriteLine("Invalid choice. Please enter 1, 2, 3, or 4.");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static X509Certificate2 GetCertificateFromStore(string thumbprint)
        {
            using (X509Store store = new X509Store(StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = store.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, validOnly: false);

                if (certCollection.Count == 0)
                {
                    throw new Exception("Certificate not found.");
                }

                return certCollection[0];
            }
        }

        private static async Task<string> GetAccessToken(string tenantId, string clientId, X509Certificate2 certificate)
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithCertificate(certificate)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            AuthenticationResult result = await app.AcquireTokenForClient(new[] { "https://lennoxfamily.sharepoint.com/.default" }).ExecuteAsync();
            return result.AccessToken;
        }

        private static void GetSiteTitle(ClientContext context)
        {
            Web web = context.Web;
            context.Load(web, w => w.Title);
            context.ExecuteQuery();

            Console.WriteLine($"Site Title: {web.Title}");
        }

        private static void CreateList(ClientContext context, string listName)
        {
            ListCreationInformation creationInfo = new ListCreationInformation
            {
                Title = listName,
                TemplateType = (int)ListTemplateType.GenericList
            };

            List list = context.Web.Lists.Add(creationInfo);
            list.Description = "Created via CSOM";
            list.Update();

            context.ExecuteQuery();

            Console.WriteLine($"List '{listName}' created successfully.");
        }

        private static void RetrieveListItems(ClientContext context, string listName)
        {
            List list = context.Web.Lists.GetByTitle(listName);
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View><RowLimit>10</RowLimit></View>";

            ListItemCollection items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            Console.WriteLine($"Items in list '{listName}':");

            foreach (var item in items)
            {
                Console.WriteLine($"ID: {item.Id}, Title: {item["Title"]}");
            }
        }
    }
}
