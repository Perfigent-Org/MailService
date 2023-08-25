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
        public async static Task<KeyValuePair<string, string>> Login(string clientId, string clientSecret)
        {
            string[] scopes = { "https://mail.google.com/", "https://www.googleapis.com/auth/userinfo.email" };

            Console.WriteLine("Requesting authorization");

            UserCredential credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = clientId,
                    ClientSecret = clientSecret,
                },
                 scopes,
                 clientId,
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

            return await GetToken(credential);
            //bool revokeToken = false;
            //if (revokeToken)
            //{
            //    credential.RevokeTokenAsync(CancellationToken.None).Wait();
            //}
        }

        private static async Task<KeyValuePair<string, string>> GetToken(UserCredential credential)
        {
            Console.WriteLine("Requesting the e-mail address of the user from Google");

            Oauth2Service oauthService = new Oauth2Service(new BaseClientService.Initializer() { HttpClientInitializer = credential });
            Userinfo userInfo = await oauthService.Userinfo.Get().ExecuteAsync();
            Console.WriteLine("E-mail address is " + userInfo.Email);

            return new KeyValuePair<string, string>(userInfo.Email, credential.Token.AccessToken);
        }

        public static async Task<KeyValuePair<string, string>> RefreshToken(string clientId, string clientSecret) => await Login(clientId, clientSecret);

        public static async Task Logout(string clientId)
        {
            FileDataStore ds = new FileDataStore(GoogleWebAuthorizationBroker.Folder);
            await ds.DeleteAsync<TokenResponse>(clientId);
        }
    }
}
