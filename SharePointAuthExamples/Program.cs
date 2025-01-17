using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace SharePointApiExample
{
    class Program
    {
        private static async Task Main(string[] args)
        {
            // Azure AD and SharePoint Configuration
            string tenantId = "f4cbf802-fffb-4961-8633-9a240e5d234a";
            string clientId = "bdc37915-34aa-4ec6-8a86-1ce1363a09b1";
            string certificateThumbprint = "82062f3830d6ae387e878dc0cb31e616c78042f8";
            string sharepointSiteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";

            try
            {
                // Load Certificate
                //X509Certificate2 certificate = GetCertificateFromStore(certificateThumbprint);
                X509Certificate2 certificate = LoadCertificateFromFile("C:\\Users\\iain\\certificate.pfx", "Password");

                //X509Certificate2 certificate = GetCertificateFromStore(certificateThumbprint);

                // Get Access Token
                string[] scopes = { "https://lennoxfamily.sharepoint.com/.default" };
                string accessToken = await GetAccessTokenAsync(tenantId, clientId, certificate, scopes);

                while (true)
                {
                    // Display menu
                    Console.WriteLine("\nMenu:");
                    Console.WriteLine("1. Make default API call (_api/web)");
                    Console.WriteLine("2. Make a custom REST API call");
                    Console.WriteLine("3. Create a new list");
                    Console.WriteLine("4. Get all items from a list");
                    Console.WriteLine("5. Exit");
                    Console.Write("Choose an option: ");

                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            await CallSharePointApi(sharepointSiteUrl, "_api/web", accessToken);
                            break;

                        case "2":
                            string customEndpoint = GetCustomApiEndpoint();
                            await CallSharePointApi(sharepointSiteUrl, customEndpoint, accessToken);
                            break;

                        case "3":
                            Console.Write("Enter the name of the new list: ");
                            string listName = Console.ReadLine()?.Trim();
                            if (!string.IsNullOrWhiteSpace(listName))
                            {
                                await CreateList(sharepointSiteUrl, listName, accessToken);
                            }
                            else
                            {
                                Console.WriteLine("List name cannot be empty.");
                            }
                            break;

                        case "4":
                            Console.Write("Enter the name of the list to retrieve items from: ");
                            string listToRetrieve = Console.ReadLine()?.Trim();
                            if (!string.IsNullOrWhiteSpace(listToRetrieve))
                            {
                                await GetListItems(sharepointSiteUrl, listToRetrieve, accessToken);
                            }
                            else
                            {
                                Console.WriteLine("List name cannot be empty.");
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

        private static async Task<string> GetAccessTokenAsync(string tenantId, string clientId, X509Certificate2 certificate, string[] scopes)
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithCertificate(certificate)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }

        private static async Task CallSharePointApi(string siteUrl, string endpoint, string accessToken)
        {
            using (HttpClient client = new HttpClient())
            {
                // Add Authorization Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Set the Accept header to request JSON
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // Construct API URL
                string apiUrl = $"{siteUrl}/{endpoint.TrimStart('/')}";

                Console.WriteLine($"\nCalling SharePoint API: {apiUrl}");

                // Make the API Call
                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    PrettyPrintJson(content);
                }
                else
                {
                    Console.WriteLine($"\nError: {response.StatusCode} - {response.ReasonPhrase}");
                }
            }
        }

        private static async Task CreateList(string siteUrl, string listName, string accessToken)
        {
            using (HttpClient client = new HttpClient())
            {
                // Add Authorization Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // Construct API URL
                string apiUrl = $"{siteUrl}/_api/web/lists";

                // Payload for creating a new list
                var payload = new
                {
                    Title = listName,
                    BaseTemplate = 100, // Template for a generic custom list
                    AllowContentTypes = true,
                    ContentTypesEnabled = true
                };

                // Serialize payload to JSON
                string jsonPayload = JsonSerializer.Serialize(payload);

                // Create HttpContent and set the Content-Type header
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                // Make the POST request
                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"\nList '{listName}' created successfully!");
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"\nError creating list: {response.StatusCode} - {response.ReasonPhrase}");
                    Console.WriteLine($"Details: {error}");
                }
            }
        }

        private static async Task GetListItems(string siteUrl, string listName, string accessToken)
        {
            using (HttpClient client = new HttpClient())
            {
                // Add Authorization Header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Set the Accept header to request JSON
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // Construct API URL to get list items
                string apiUrl = $"{siteUrl}/_api/web/lists/getbytitle('{listName}')/items";

                Console.WriteLine($"\nCalling SharePoint API to retrieve items from list: {listName}");

                // Make the GET request
                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();

                    // Attempt to parse the response as JSON
                    try
                    {
                        var parsedJson = JsonSerializer.Deserialize<object>(content, new JsonSerializerOptions { WriteIndented = true });
                        Console.WriteLine("\nList Items:");
                        Console.WriteLine(JsonSerializer.Serialize(parsedJson, new JsonSerializerOptions { WriteIndented = true }));
                    }
                    catch (JsonException)
                    {
                        Console.WriteLine("\nResponse is not valid JSON. Raw output:");
                        Console.WriteLine(content);
                    }
                }
                else
                {
                    Console.WriteLine($"\nError retrieving items: {response.StatusCode} - {response.ReasonPhrase}");
                    string error = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Details: {error}");
                }
            }
        }

        private static string GetCustomApiEndpoint()
        {
            while (true)
            {
                Console.Write("Enter the custom API endpoint (e.g., '_api/lists'): ");
                string customEndpoint = Console.ReadLine()?.Trim();

                if (string.IsNullOrWhiteSpace(customEndpoint))
                {
                    Console.WriteLine("The endpoint cannot be empty. Please try again.");
                }
                else if (!customEndpoint.StartsWith("_api/", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("The endpoint must start with '_api/'. Please try again.");
                }
                else
                {
                    return customEndpoint;
                }
            }
        }

        private static void PrettyPrintJson(string json)
        {
            try
            {
                var parsedJson = JsonSerializer.Deserialize<object>(json);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string prettyJson = JsonSerializer.Serialize(parsedJson, options);
                Console.WriteLine("\nAPI Response:");
                Console.WriteLine(prettyJson);
            }
            catch (JsonException)
            {
                Console.WriteLine("\nResponse is not valid JSON. Raw output:");
                Console.WriteLine(json);
            }
        }
    }
}
