using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Oauth2.v2;
using Google.Apis.Oauth2.v2.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MailService.OAuthGoogle
{
    public static class OAuth20
    {
        public async static Task<KeyValuePair<string, string>> Login(string userEmail = null)
        {
            string yourAppClientID = "466636461440-nob44m3k4c9jr1n3fdtfr4s0vogoh3p4.apps.googleusercontent.com";
            string yourAppClientSecret = "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";

            string gmailScope = "https://mail.google.com/";

            string[] scopes;

            if (userEmail != null)
            {
                scopes = new string[1];
            }
            else
            {
                scopes = new string[2];
            }
            scopes[0] = gmailScope;

            if (userEmail == null)
            {
                scopes[1] = "https://www.googleapis.com/auth/userinfo.email";
            }

            Console.WriteLine("Requesting authorization");

            UserCredential credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = yourAppClientID,
                    ClientSecret = yourAppClientSecret,
                },
                 scopes,
                 "user",
                 CancellationToken.None);

            Console.WriteLine("Authorization granted or not required (if the saved access token already available)");

            if (credential.Token.IsExpired(credential.Flow.Clock))
            {
                Console.WriteLine("The access token has expired, refreshing it");

                if (await credential.RefreshTokenAsync(CancellationToken.None))
                {
                    Console.WriteLine("The access token is now refreshed");
                }
                else
                {
                    Console.WriteLine("The access token has expired but we can't refresh it :(");
                    return new KeyValuePair<string, string>();
                }
            }
            else
            {
                Console.WriteLine("The access token is OK, continue");
            }

            if (userEmail == null)
            {
                Console.WriteLine("Requesting the e-mail address of the user from Google");

                Oauth2Service oauthService = new Oauth2Service(new BaseClientService.Initializer() { HttpClientInitializer = credential });
                Userinfo userInfo = await oauthService.Userinfo.Get().ExecuteAsync();
                userEmail = userInfo.Email;
                Console.WriteLine("E-mail address is " + userEmail);
            }

            return new KeyValuePair<string, string>(userEmail, credential.Token.AccessToken);

            //bool revokeToken = false;
            //if (revokeToken)
            //{
            //    credential.RevokeTokenAsync(CancellationToken.None).Wait();
            //}
        }

        public static async Task<KeyValuePair<string, string>> RefreshToken() => await Login();

        public static async Task Logout()
        {
            FileDataStore ds = new FileDataStore(GoogleWebAuthorizationBroker.Folder);
            await ds.DeleteAsync<TokenResponse>("user");
        }
    }
}
