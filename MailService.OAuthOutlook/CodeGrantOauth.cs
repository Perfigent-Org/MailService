using Newtonsoft.Json;
using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailService.OAuthOutlook
{
    internal class CodeGrantOauth
    {
        private string _accessToken = null;
        private string _refreshToken = null;
        private string _authorizationCode = null;
        private int _expiration;
        private string _error = null;

        // Production OAuth server endpoints.
        private string AuthorizationUri = "https://login.live.com/oauth20_authorize.srf";  // Authorization code endpoint.
        private static string RedirectUri = "https://login.live.com/oauth20_desktop.srf";  // Callback endpoint. Used only with Native apps. In 'useLocalServer' mode, we use LoopbackCallback instead.
        private string RefreshUri = "https://login.live.com/oauth20_token.srf";  // Get tokens endpoint.
        private string LogoutUri = "https://login.live.com/oauth20_logout.srf";

        // Scopes is currently set to support SMTP/IMAP, getting email address of the user, and working in offline mode.
        // (i.e. that the app can refresh access tokens without the user's intervention).
        // If you change scopes, you need to reqeust the refresh token again! For instance, delete token.json file for that.
        private const string Scopes = "wl.imap,wl.emails,wl.offline_access";

        // codeQueryString is the query string for the authorizationUri. To force user log in, include the &prompt=login parameter.
        private const string CodeQueryString = "?client_id={0}&scope=" + Scopes + "&response_type=code&redirect_uri={1}";

        // accessBody is the request body used with the refreshUri to get the access token using the authorization code.
        private const string AccessBody = "client_id={0}&code={1}&grant_type=authorization_code&redirect_uri={2}";

        // refreshBody is the request body used with the refreshUri to get the access token using a refresh token.
        private const string RefreshBody = "client_id={0}&grant_type=refresh_token&redirect_uri={1}&refresh_token={2}";

        private const string LogoutBody = "?client_id={0}&redirect_uri={1}";

        private string _clientId = null;
        private string _uri = null;
        private string _logoutUri = null;

        public string AccessToken { get { return this._accessToken; } }
        public string RefreshToken { get { return this._refreshToken; } }
        public int Expiration { get { return this._expiration; } }
        public string Error { get { return this._error; } }

        public CodeGrantOauth(string clientId)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("The client ID is missing.");
            }

            this._clientId = clientId;
            this._uri = string.Format(this.AuthorizationUri + CodeQueryString, this._clientId, RedirectUri);
            this._logoutUri = string.Format(this.LogoutUri + LogoutBody, this._clientId, RedirectUri);
        }

        public async Task<string> GetAccessTokenAsync()
        {
            //ProcessStartInfo psi = new ProcessStartInfo
            //{
            //    FileName = _uri,
            //    UseShellExecute = true

            //};
            //Process.Start(psi);

            // Here we're requesting the user to copy/paste the code themselves (not UX-friendly but easy to implement). With a desktop app or Forms functionality (.NET Core 3.0+),
            // it should be possible to intercept the authorization code right from the browser. See "WinForms\.NET 4.5 OAuth\C#\OAuthConsoleApps\MicrosoftLogin.cs"
            // or https://docs.microsoft.com/en-us/advertising/shopping-content/code-example-authentication-oauth to get the idea.
            // Or, simply use the 'useLocalServer' mode of this sample.
            //Console.Write("Please enter the authorization code: ");
            //string authorizationCode = Console.ReadLine().Trim();

            string authorizationCode = await StartTaskAsSTAThread(() => RunWebBrowserFormAndGetCode(_uri));

            if (string.IsNullOrEmpty(authorizationCode))
            {
                this._error = "No authorization code provided";
                return null;
            }

            _authorizationCode = authorizationCode;
            var accessTokenRequestBody = string.Format(AccessBody, this._clientId, this._authorizationCode, WebUtility.UrlEncode(RedirectUri));
            AccessTokenResponse tokensFromServer = await GetTokensAsync(this.RefreshUri, accessTokenRequestBody);
            this._accessToken = tokensFromServer.AccessToken;
            this._refreshToken = tokensFromServer.RefreshToken;
            this._expiration = tokensFromServer.Expiration;

            return this._accessToken;
        }

        // Get the access token by using the refresh token.
        public async Task<string> RefreshAccessTokenAsync(string refreshToken)
        {
            string refreshTokenRequestBody = string.Format(RefreshBody, this._clientId, WebUtility.UrlEncode(RedirectUri), refreshToken);
            AccessTokenResponse tokensFromServer = await GetTokensAsync(this.RefreshUri, refreshTokenRequestBody);
            this._accessToken = tokensFromServer.AccessToken;
            this._refreshToken = tokensFromServer.RefreshToken;
            this._expiration = tokensFromServer.Expiration;
            return this._accessToken;
        }

        public async Task<bool> Logout()
        {
            return await StartTaskAsSTAThread(() => RunWebBrowserFormAndLogout(_logoutUri));
        }

        // Called by GetAccessToken and RefreshAccessToken to get an access token.
        private static async Task<AccessTokenResponse> GetTokensAsync(string uri, string body)
        {
            AccessTokenResponse tokensFromServer = null;
            var request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "POST";
            request.Accept = "application/json";
            request.ContentType = "application/x-www-form-urlencoded";

            request.ContentLength = body.Length;

            using (Stream requestStream = await request.GetRequestStreamAsync())
            {
                StreamWriter writer = new StreamWriter(requestStream);
                await writer.WriteAsync(body);
                writer.Close();
            }

            HttpWebResponse response = null;
            try
            {
                try
                {
                    response = (HttpWebResponse)await request.GetResponseAsync();
                }
                catch (WebException wex)
                {
                    if (wex.Status == WebExceptionStatus.ProtocolError)
                    {
                        await ExceptionUtils.ThrowWithBodyAsync(wex);
                    }
                    else
                    {
                        throw;
                    }
                }

                using (Stream responseStream = response.GetResponseStream())
                {
                    var reader = new StreamReader(responseStream);
                    string json = await reader.ReadToEndAsync();
                    reader.Close();
                    tokensFromServer = JsonConvert.DeserializeObject(json, typeof(AccessTokenResponse)) as AccessTokenResponse;
                }
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }

            return tokensFromServer;
        }

        private static string RunWebBrowserFormAndGetCode(string url)
        {
            string code = null;

            Form webBrowserForm = new Form();
            webBrowserForm.WindowState = FormWindowState.Maximized;
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.Dock = DockStyle.Fill;
            webBrowser.Url = new Uri(url);

            webBrowserForm.Controls.Add(webBrowser);

            WebBrowserDocumentCompletedEventHandler documentCompletedHandler = (s, e) =>
            {
                string[] parts = webBrowser.Url.Query.Split(new char[] { '?', '&' });
                foreach (string part in parts)
                {
                    if (part.StartsWith("code="))
                    {
                        code = part.Split('=')[1];
                        webBrowserForm.Close();
                    }
                    else if (part.StartsWith("error="))
                    {
                        webBrowserForm.Close();
                    }
                }
            };

            webBrowser.DocumentCompleted += documentCompletedHandler;
            Application.Run(webBrowserForm);
            webBrowser.DocumentCompleted -= documentCompletedHandler;

            return code;
        }

        private static bool RunWebBrowserFormAndLogout(string url)
        {
            bool isLoggedOut = false;

            Form webBrowserForm = new Form();
            webBrowserForm.WindowState = FormWindowState.Maximized;
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.Dock = DockStyle.Fill;
            webBrowser.Url = new Uri(url);

            webBrowserForm.Controls.Add(webBrowser);

            WebBrowserDocumentCompletedEventHandler documentCompletedHandler = (s, e) =>
            {
                isLoggedOut = webBrowser.Url.ToString().Contains(RedirectUri);
                webBrowserForm.Close();
            };

            webBrowser.DocumentCompleted += documentCompletedHandler;
            Application.Run(webBrowserForm);
            webBrowser.DocumentCompleted -= documentCompletedHandler;

            return isLoggedOut;
        }

        // We must run WebBrowser control in a separate STA thread (otherwise, it just wouldn't work in a console app),
        // but it's not required for Forms-based apps (there you'd simply place it on the form and that't all).
        private static Task<T> StartTaskAsSTAThread<T>(Func<T> taskFunc)
        {
            TaskCompletionSource<T> tcs = new TaskCompletionSource<T>();
            Thread thread = new Thread(() =>
            {
                try
                {
                    tcs.SetResult(taskFunc());
                }
                catch (Exception e)
                {
                    tcs.SetException(e);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }
        // Loopback version ('useLocalServer' mode). Starts a local web server on localhost, and makes the browser to access the page hosted on this "server".
        // When Microsoft web server sends the browser the redirect URL (and that URL's domain is 'localhost'), the browser loads the page
        // from this local server passing the verification code as a URL parameter. We extract the verification code from it then.

        // Based on Google.Apis.Auth.OAuth2.LocalServerCodeReceiver.
        // https://github.com/googleapis/google-api-dotnet-client/blob/master/Src/Support/Google.Apis.Auth/OAuth2/LocalServerCodeReceiver.cs
        // Their implementation is far more advanced and robust, you can check it for details.

        // http://localhost:49217/signin-microsoft must be registered in apps.dev.microsoft.com portal for your app. 49217 is just an arbitrary number, can be anything.
        private const string LoopbackCallback = "http://localhost:{0}/signin-microsoft/";
        //        private const string ClosePageResponse =
        //@"<html>
        //  <head><title>OAuth 2.0 Authentication Token Received</title></head>
        //  <body>
        //    Received verification code. Closing...
        //    <script type='text/javascript'>
        //      window.setTimeout(function() {
        //          window.open('', '_self', ''); 
        //          window.close(); 
        //        }, 1000);
        //      if (window.opener) { window.opener.checkToken(); }
        //    </script>
        //  </body>
        //</html>";

        // So we start a local web server on any unused port and generate the redirect_uri. Something like "http://localhost:60780/signin-microsoft/".
        // On apps.dev.microsoft.com portal, you registered redirect_uri like "http://localhost:49217/signin-microsoft". It's not a problem that port numbers are different.
        // You only need to make that that the parts before and after the port number match.
        //private string _redirectUriLocal;
        //public string RedirectUriLocal
        //{
        //    get
        //    {
        //        if (!string.IsNullOrEmpty(_redirectUriLocal))
        //        {
        //            return _redirectUriLocal;
        //        }

        //        return _redirectUriLocal = string.Format(LoopbackCallback, GetRandomUnusedPort());
        //    }
        //}

        //public async Task<string> GetAccessTokenFromLocalServerAsync(string clientSecret)
        //{
        //    string authorizationUrl = string.Format(this.AuthorizationUri + CodeQueryString, this._clientId, RedirectUriLocal);
        //    using (HttpListener listener = new HttpListener())
        //    {
        //        listener.Prefixes.Add(RedirectUriLocal);
        //        try
        //        {
        //            listener.Start();


        //            ProcessStartInfo psi = new ProcessStartInfo
        //            {
        //                FileName = authorizationUrl,
        //                UseShellExecute = true
        //            };

        //            Process.Start(psi);

        //            // Wait to get the authorization code response.
        //            HttpListenerContext context = await listener.GetContextAsync().ConfigureAwait(false);
        //            NameValueCollection coll = context.Request.QueryString;

        //            // Write a "close" response.
        //            using (var writer = new StreamWriter(context.Response.OutputStream))
        //            {
        //                writer.WriteLine(ClosePageResponse);
        //                writer.Flush();
        //            }
        //            context.Response.OutputStream.Close();

        //            Dictionary<string, string> dict = coll.AllKeys.ToDictionary(k => k, k => coll[k]);

        //            string authCode = dict["code"]; // In production, be sure to check that the "code" parameter is actually there.
        //            string accessTokenRequestBody = string.Format(AccessBody, this._clientId, authCode, WebUtility.UrlEncode(RedirectUriLocal));
        //            accessTokenRequestBody += "&client_secret=" + clientSecret; // Unlike GetAccessTokenAsync, GetAccessTokenFromLocalServerAsync also needs client secret.

        //            AccessTokenResponse tokensFromServer = await GetTokensAsync(this.RefreshUri, accessTokenRequestBody);

        //            this._accessToken = tokensFromServer.AccessToken;
        //            this._refreshToken = tokensFromServer.RefreshToken;
        //            this._expiration = tokensFromServer.Expiration;

        //            return this._accessToken;
        //        }
        //        finally
        //        {
        //            listener.Close();
        //        }
        //    }
        //}

        //private static int GetRandomUnusedPort()
        //{
        //    var listener = new TcpListener(IPAddress.Loopback, 0);
        //    try
        //    {
        //        listener.Start();
        //        return ((IPEndPoint)listener.LocalEndpoint).Port;
        //    }
        //    finally
        //    {
        //        listener.Stop();
        //    }
        //}
    }
}
