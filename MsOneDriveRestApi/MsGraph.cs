using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsOneDriveRestApi
{
    public class MsGraph
    {
        //Nuget:
        //Microsoft.Identity.Client

        //For ClientId
        //Go to https://apps.dev.microsoft.com/#/appList
        //Create new Converged applications
        //After create > Click "Add Platform" and create "Native Application"
        //Copy "Application Id" and past here
        static string clientId = "App id";

        public static DateTimeOffset Expiration;
        private static string TokenForUser = null;
        private static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenAsync()
        {
            //Scopes > https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
            AuthenticationResult authResult;
            var scopes = new string[]
                    {
                        "https://graph.microsoft.com/User.Read",
                        "https://graph.microsoft.com/Files.ReadWrite",
                        "https://graph.microsoft.com/Files.ReadWrite.AppFolder"
                };

            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(scopes, IdentityClientApp.Users.First());
                TokenForUser = authResult.AccessToken;
            }
            catch (Exception)
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenAsync(scopes);

                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }

            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            foreach (var user in IdentityClientApp.Users)
            {
                IdentityClientApp.Remove(user);
            }

            TokenForUser = null;
        }
    }
}
