using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace MsalAppOnlyConsoleSdk
{
    internal class Program
    {
        private static string tenantName = "";
        private static string userName = "";
        private static string token = null;

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private async static Task<string> GetAccessToken()
        {
            if (token != null)
            {
                return token;
            }

            var clientId = "";
            var clientSecret = "";
            var authority = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/";
            var azureApp = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithClientSecret(clientSecret)
                .Build();

            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var authResult = await azureApp.AcquireTokenForClient(scopes).ExecuteAsync();
            return authResult.AccessToken;
        }

        public class TokenProvider : IAccessTokenProvider
        {
            public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
                CancellationToken cancellationToken = default)
            {
                var token = await GetAccessToken();
                return token;
            }

            public AllowedHostsValidator AllowedHostsValidator { get; }
        }

        private async static Task MakeCallsToMicrosoftGraph()
        {
            var provider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
            var client = new GraphServiceClient(provider);
            var userEmail = $"{userName}@{tenantName}.onmicrosoft.com";

            var profile = await client.Users[userEmail].GetAsync();
            Console.WriteLine("Profile:");
            Console.WriteLine($"The user's name is {profile?.DisplayName}. The user's title is {profile?.JobTitle}");
            Console.WriteLine();
            Console.WriteLine();

            var messages = await client.Users[userEmail].Messages.GetAsync();
            Console.WriteLine("Mail:");
            if (messages != null && messages.Value != null)
            {
                foreach (var message in messages.Value)
                {
                    Console.WriteLine($"{message.Subject} from {message?.Sender?.EmailAddress?.Address}");
                }
                Console.WriteLine();
                Console.WriteLine();
            }

            var drive = await client.Users[userEmail].Drive.GetAsync();
            if (drive != null)
            {
                var files = await client.Drives[drive.Id].Items["root"].Children.GetAsync();
                Console.WriteLine("Files:");
                if (files != null && files.Value != null)
                {
                    foreach (var file in files.Value)
                    {
                        Console.WriteLine($"{file.Name} created on {file.CreatedDateTime.Value.ToString("D")}");
                    }
                }
            }
        }

    }
}
