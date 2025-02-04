using System;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Identity.Client;
using System.Threading.Tasks;

namespace SharePointCsomExample
{
    class Program
    {
        static void Main(string[] args) //Main / OriginalMain
        {
            // Azure AD and SharePoint Configuration
            string tenantId = "f4cbf802-fffb-4961-8633-9a240e5d234a";
            string clientId = "bdc37915-34aa-4ec6-8a86-1ce1363a09b1";
            string certificateThumbprint = "82062f3830d6ae387e878dc0cb31e616c78042f8"; //82062F3830D6AE387E878DC0CB31E616C78042F8
            string siteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";

            try
            {
                // Load the certificate
                X509Certificate2 certificate = LoadCertificateFromFile("C:\\Users\\iain\\certificate.pfx", "FinePix2004!!");

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
                        Console.WriteLine("4. List Documents in a Library");
                        Console.WriteLine("5. Exit");
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
                                Console.Write("Enter the name of the document library: ");
                                string libraryName = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(libraryName))
                                {
                                    ListDocumentsInLibrary(context, libraryName);
                                }
                                else
                                {
                                    Console.WriteLine("Library name cannot be empty.");
                                }
                                break;

                            case "5":
                                Console.WriteLine("Exiting...");
                                return;

                            default:
                                Console.WriteLine("Invalid choice. Please enter 1, 2, 3, 4, or 5.");
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

        private static X509Certificate2 LoadCertificateFromFile(string certificatePath, string certificatePassword)
        {
            if (string.IsNullOrWhiteSpace(certificatePath) || !System.IO.File.Exists(certificatePath))
            {
                throw new Exception("Certificate file not found.");
            }

            return new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.EphemeralKeySet);
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

        private static void ListDocumentsInLibrary(ClientContext context, string libraryName)
        {
            try
            {
                List documentLibrary = context.Web.Lists.GetByTitle(libraryName);
                CamlQuery query = new CamlQuery
                {
                    ViewXml = "<View><ViewFields><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' /></ViewFields></View>"
                };

                ListItemCollection items = documentLibrary.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                Console.WriteLine($"Documents in library '{libraryName}':");
                foreach (var item in items)
                {
                    string fileName = item["FileLeafRef"]?.ToString();
                    string fileUrl = item["FileRef"]?.ToString();
                    Console.WriteLine($"Name: {fileName}, URL: {fileUrl}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing documents: {ex.Message}");
            }
        }
    }
}
