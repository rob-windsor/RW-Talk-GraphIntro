using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;

namespace MsalDelegatedConsoleSdk
{
    internal class Program
    {
        private static string tenantName = "robwindsortest980";
        private static string token = null;

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private static async Task<string> GetAccessToken()
        {
            if (token != null)
            {
                return token;
            }

            var clientId = "43ec3caf-b22b-4fe3-84fd-340bd5b384cb";
            var authority = $"https://login.microsoftonline.com/{tenantName}.onmicrosoft.com/";
            var azureApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithRedirectUri("http://localhost")
                .Build();

            var scopes = new string[] { "User.Read", "Mail.ReadBasic", "Files.Read" };
            var authResult = await azureApp.AcquireTokenInteractive(scopes).ExecuteAsync();

            token = authResult.AccessToken;
            return token;
        }

        public class TokenProvider : IAccessTokenProvider
        {
            public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
                CancellationToken cancellationToken = default)
            {
                var token =  await GetAccessToken();
                return token;
            }

            public AllowedHostsValidator AllowedHostsValidator { get; }
        }

        private async static Task MakeCallsToMicrosoftGraph()
        {
            var provider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
            var client = new GraphServiceClient(provider);

            var profile = await client.Me.GetAsync();
            Console.WriteLine("Profile:");
            Console.WriteLine($"Your name is {profile?.DisplayName}. Your title is {profile?.JobTitle}");
            Console.WriteLine();
            Console.WriteLine();

            var messages = await client.Me.Messages.GetAsync();
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

            var drive = await client.Me.Drive.GetAsync();
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
