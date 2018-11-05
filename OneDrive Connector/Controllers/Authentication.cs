using System;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;


namespace OneDrive_Connector.Controllers
{
    class Authentication
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = env.ClientId;

        // The Group.Read.All permission is an admin-only scope, so authorization will fail if you 
        // want to sign in with a non-admin account. Remove that permission and comment out the group operations in 
        // the UserMode() method if you want to run this sample with a non-admin account.
        public static string[] Scopes = { "User.Read", "Mail.Send", "Files.ReadWrite" };

        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);
        public static string UserToken = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            // Create Microsoft Graph client.
            try
            {
                graphClient = new GraphServiceClient(
                    "https://graph.microsoft.com/v1.0",
                    new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            var token = await GetTokenForUserAsync();
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        }));
                return graphClient;
            }

            catch (Exception ex)
            {
                Debug.WriteLine("Could not create a graph client: " + ex.Message);
            }

            return graphClient;
        }

        
        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult;
            try
            {
                var CacheRequest = (await IdentityClientApp.GetAccountsAsync());
                var First = CacheRequest.First();
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, First);
                UserToken = authResult.AccessToken;
            }

            catch (Exception)
            {
                if (UserToken == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);

                    UserToken = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }

            return UserToken;
        }

    }
}
