using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace ReadEmail
{
    class Program
    {
        static string tenantId = "Your tenantId ";
        static string clientId = "Your clientid  ";
        static string clientSecret = "Your client secret ";
        static string userId = "User Object Id or GUID ";
        static async System.Threading.Tasks.Task Main(string[] args)
        {


            await SendMailAsync();

        }
        public static async System.Threading.Tasks.Task SendMailAsync()
        {
            var message = new Message
            {
                Subject = "Subjec Text of Email",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Content Message"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "Email Address"
                        }
                    }
                },
                CcRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "Email Address"
                        }
                    }
                }
            };
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
            IConfidentialClientApplication confidentialClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}/v2.0"))
                .Build();

            // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
            var authResult = await confidentialClient
                    .AcquireTokenForClient(scopes)
                    .ExecuteAsync().ConfigureAwait(false);

            var token = authResult.AccessToken;
            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );
            await graphServiceClient.Users[userId]
                  .SendMail(message, false)
                  .Request()
                  .PostAsync();
        }

    }
}
