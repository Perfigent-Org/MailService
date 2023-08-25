using Microsoft.Identity.Client;

namespace MailService.OAuthOffice365
{
    public class Client
    {
        private static readonly string Tenant = "common";
        private static readonly string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp;

        public static IPublicClientApplication PublicClientApp(string clientId)
        {
            if (_clientApp == null)
            {
                _clientApp = PublicClientApplicationBuilder.Create(clientId)
                    .WithAuthority($"{Instance}{Tenant}")
                    .WithDefaultRedirectUri()
                    .Build();
                TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
            }

            return _clientApp;
        }
    }
}
