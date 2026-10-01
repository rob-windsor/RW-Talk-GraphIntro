using Azure.Identity;
using Microsoft.Graph;

namespace ConsoleAppOnlySdk
{
    internal class Program
    {
        private static string tenantName = "robwindsortest980";
        private static string userName = "meganb";

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private async static Task MakeCallsToMicrosoftGraph()
        {
            var credential = new ClientSecretCredential(
                tenantId: $"{tenantName}.onmicrosoft.com",
                clientId: "a74dea2b-0ea9-42cc-8c54-e28b0b17a54b",
                clientSecret: "fde8Q~L_c1pyXvPQWEUACCUXL.RCBGlcuHqy8c06"
            );
            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var client = new GraphServiceClient(credential, scopes);
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
                    Console.WriteLine($"{message?.Subject} from {message?.Sender?.EmailAddress?.Address}");
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
                        Console.WriteLine($"{file?.Name} created on {file?.CreatedDateTime?.ToString("D")}");
                    }
                }
            }
        }

    }
}
