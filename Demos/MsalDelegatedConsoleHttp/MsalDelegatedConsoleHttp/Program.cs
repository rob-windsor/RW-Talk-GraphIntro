using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using System.Text;

namespace MsalDelegatedConsoleHttp
{
    internal class Program
    {
        private static string tenantName = "robwindsortest980";

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private static async Task<string> GetAccessToken()
        {
            var clientId = "5ee80709-aba3-4ab9-9ccf-be69992f16ce";
            var authority = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/";
            var azureApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithRedirectUri("http://localhost")
                .Build();

            var scopes = new string[] { "User.Read", "Mail.ReadBasic", "Files.ReadWrite" };
            var authResult = await azureApp.AcquireTokenInteractive(scopes).ExecuteAsync();

            return authResult.AccessToken;
        }

        private async static Task MakeCallsToMicrosoftGraph()
        {
            var token = await GetAccessToken();
            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = $"https://graph.microsoft.com/v1.0/me";
            using (var response = await client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                var jsonObject = JsonObject.Parse(json);

                Console.WriteLine("Profile:");
                Console.WriteLine($"Your name is {jsonObject?["displayName"]}. Your title is {jsonObject?["jobTitle"]}");
                Console.WriteLine();
                Console.WriteLine();
            }

            url = "https://graph.microsoft.com/v1.0/me/messages";
            using (var response = await client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                var jsonObject = JsonObject.Parse(json);
                var jsonArray = jsonObject?["value"] as JsonArray;

                Console.WriteLine("Mail:");
                if (jsonArray != null)
                {
                    foreach (var item in jsonArray)
                    {
                        Console.WriteLine($"{item?["subject"]} from {item?["sender"]?["emailAddress"]?["address"]}");
                    }
                    Console.WriteLine();
                    Console.WriteLine();
                }
            }

            url = "https://graph.microsoft.com/v1.0/me/drive/root:/GraphDemo.txt:/content";
            var fileContent = "Updated from the MsalDelegatedConsoleHttp console application.";
            using (var uploadContent = new StringContent(fileContent, Encoding.UTF8, "text/plain"))
            {
                using (var uploadResponse = await client.PutAsync(url, uploadContent))
                {
                    uploadResponse.EnsureSuccessStatusCode();
                    Console.WriteLine("GraphDemo.txt uploaded/updated in OneDrive.");
                }
            }

            url = "https://graph.microsoft.com/v1.0/me/drive/root/children";
            using (var response = await client.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                var jsonObject = JsonObject.Parse(json);
                var jsonArray = jsonObject?["value"] as JsonArray;

                Console.WriteLine("Files:");
                if (jsonArray != null)
                {
                    foreach (var node in jsonArray)
                    {
                        Console.WriteLine($"{node?["name"]} last modified on {node?["lastModifiedDateTime"]?.GetValue<DateTime>().ToString("D")}");
                    }
                }
            }
        }

    }
}
