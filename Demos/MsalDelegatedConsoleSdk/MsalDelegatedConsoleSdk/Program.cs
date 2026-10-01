using Azure.Identity;
using Microsoft.Graph;
using System.Text;

namespace MsalDelegatedConsoleSdk
{
    internal class Program
    {
        private static string tenantName = "";

        static void Main(string[] args)
        {
            MakeCallsToMicrosoftGraph().Wait();
        }

        private async static Task MakeCallsToMicrosoftGraph()
        {
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = $"{tenantName}.onmicrosoft.com",
                ClientId = "",
                RedirectUri = new Uri("http://localhost")
            };

            var credential = new InteractiveBrowserCredential(options);
            var scopes = new string[] { "User.Read", "Mail.ReadBasic", "Files.ReadWrite" };
            var client = new GraphServiceClient(credential, scopes);

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
                    Console.WriteLine($"{message?.Subject} from {message?.Sender?.EmailAddress?.Address}");
                }
                Console.WriteLine();
                Console.WriteLine();
            }


            var drive = await client.Me.Drive.GetAsync();
            if (drive?.Id != null)
            {
                var fileContent = "Updated from the MsalDelegatedConsoleSdk console application.";
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(fileContent));

                var newFile = await client
                    .Drives[drive.Id]
                    .Items["root"]
                    .ItemWithPath("GraphDemo.txt")
                    .Content
                    .PutAsync(stream);

                var files = await client
                    .Drives[drive.Id]
                    .Items["root"]
                    .Children
                    .GetAsync();

                Console.WriteLine("Files:");
                if (files != null && files.Value != null)
                {
                    foreach (var file in files.Value)
                    {
                        Console.WriteLine($"{file?.Name} last modified on {file?.LastModifiedDateTime?.ToString("D")}");
                    }
                }
            }
        }

    }
}
