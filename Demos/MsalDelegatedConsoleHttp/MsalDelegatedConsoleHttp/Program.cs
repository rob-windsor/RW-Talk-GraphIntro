using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;

namespace MsalDelegatedConsoleHttp
{
    internal class Program
    {
        private static string tenantName = "";

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private static async Task<string> GetAccessToken()
        {
            var clientId = "";
            var authority = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/";
            var azureApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithRedirectUri("http://localhost")
                .Build();

            var scopes = new string[] { "User.Read", "Mail.ReadBasic", "Files.Read" };
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

            url = $"https://graph.microsoft.com/v1.0/me/messages";
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

            url = $"https://graph.microsoft.com/v1.0/me/drive/root/children";
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
                        Console.WriteLine($"{node?["name"]} created on {node?["createdDateTime"]?.GetValue<DateTime>().ToString("D")}");
                    }
                }
            }
        }

    }
}
