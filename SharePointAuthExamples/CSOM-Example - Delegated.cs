using System;
using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using System.Threading.Tasks;

namespace SharePointDelegatedExample
{
    class Program
    {
        private static string tenantId = "f4cbf802-fffb-4961-8633-9a240e5d234a"; // Replace with your Tenant ID
        private static string clientId = "90d31291-77f6-4b09-8888-00d012e72dec"; // Replace with your Client ID
        private static string siteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";
        private static string[] scopes = { "https://lennoxfamily.sharepoint.com/.default" };

        static async Task Main(string[] args)
        {
            try
            {
                string accessToken = await GetAccessToken();

                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (sender, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                    };

                    while (true)
                    {
                        Console.WriteLine("\nMenu:");
                        Console.WriteLine("1. Get Site Title");
                        Console.WriteLine("2. Create a New List");
                        Console.WriteLine("3. Retrieve List Items");
                        Console.WriteLine("4. Create a List Item");
                        Console.WriteLine("5. Update a List Item");
                        Console.WriteLine("6. List Documents in a Library");
                        Console.WriteLine("7. Exit");
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
                                    CreateList(context, listName);
                                else
                                    Console.WriteLine("List name cannot be empty.");
                                break;

                            case "3":
                                Console.Write("Enter the name of the list to retrieve items from: ");
                                string listToRetrieve = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(listToRetrieve))
                                    RetrieveListItems(context, listToRetrieve);
                                else
                                    Console.WriteLine("List name cannot be empty.");
                                break;

                            case "4":
                                Console.Write("Enter the name of the list to add an item to: ");
                                string targetList = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(targetList))
                                {
                                    Console.Write("Enter the title for the new item: ");
                                    string itemTitle = Console.ReadLine();
                                    CreateListItem(context, targetList, itemTitle);
                                }
                                else
                                {
                                    Console.WriteLine("List name cannot be empty.");
                                }
                                break;

                            case "5":
                                Console.Write("Enter the name of the list to update an item in: ");
                                string updateList = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(updateList))
                                {
                                    Console.Write("Enter the ID of the item to update: ");
                                    if (int.TryParse(Console.ReadLine(), out int itemId))
                                    {
                                        Console.Write("Enter the new title for the item: ");
                                        string updatedTitle = Console.ReadLine();
                                        UpdateListItem(context, updateList, itemId, updatedTitle);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Invalid item ID.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("List name cannot be empty.");
                                }
                                break;

                            case "6":
                                Console.Write("Enter the name of the document library: ");
                                string libraryName = Console.ReadLine()?.Trim();
                                if (!string.IsNullOrWhiteSpace(libraryName))
                                    ListDocumentsInLibrary(context, libraryName);
                                else
                                    Console.WriteLine("Library name cannot be empty.");
                                break;

                            case "7":
                                Console.WriteLine("Exiting...");
                                return;

                            default:
                                Console.WriteLine("Invalid choice. Please enter a number between 1 and 7.");
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

        private static async Task<string> GetAccessToken()
        {
            var app = PublicClientApplicationBuilder
                .Create(clientId)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .WithRedirectUri("http://localhost")
                .Build();

            AuthenticationResult result = await app.AcquireTokenInteractive(scopes).ExecuteAsync();
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
            try
            {
                List list = context.Web.Lists.GetByTitle(listName);
                CamlQuery query = new CamlQuery
                {
                    ViewXml = "<View><RowLimit>10</RowLimit></View>"
                };

                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                Console.WriteLine($"Items in list '{listName}':");
                foreach (var item in items)
                {
                    Console.WriteLine($"ID: {item.Id}, Title: {item["Title"]}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving list items: {ex.Message}");
            }
        }

        private static void CreateListItem(ClientContext context, string listName, string title)
        {
            try
            {
                List list = context.Web.Lists.GetByTitle(listName);
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem["Title"] = title;
                newItem.Update();
                context.ExecuteQuery();
                Console.WriteLine($"Item with title '{title}' created in list '{listName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating list item: {ex.Message}");
            }
        }

        private static void UpdateListItem(ClientContext context, string listName, int itemId, string newTitle)
        {
            try
            {
                List list = context.Web.Lists.GetByTitle(listName);
                ListItem item = list.GetItemById(itemId);
                item["Title"] = newTitle;
                item.Update();
                context.ExecuteQuery();
                Console.WriteLine($"Item with ID '{itemId}' in list '{listName}' updated to title '{newTitle}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating list item: {ex.Message}");
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
