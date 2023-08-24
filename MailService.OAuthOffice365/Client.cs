using Microsoft.Identity.Client;

namespace MailService.OAuthOffice365
{
    public class Client
    {
        static Client()
        {
            _clientApp = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority($"{Instance}{Tenant}")
                .WithDefaultRedirectUri()
                .Build();
            TokenCacheHelper.EnableSerialization(_clientApp.UserTokenCache);
        }

        private static readonly string ClientId = "17b9156a-feec-4c72-ab42-acc6c6d5590a";
        private static readonly string Tenant = "common";
        private static readonly string Instance = "https://login.microsoftonline.com/";
        private static IPublicClientApplication _clientApp;

        public static IPublicClientApplication PublicClientApp { get { return _clientApp; } }
    }
}
