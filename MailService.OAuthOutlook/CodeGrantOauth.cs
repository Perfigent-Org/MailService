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
            this._uri = string.Format(this.AuthorizationUri + CodeQueryString + "&prompt=login", this._clientId, RedirectUri);
            this._logoutUri = string.Format(this.LogoutUri + LogoutBody, this._clientId, RedirectUri);
        }

        public async Task<string> GetAccessTokenAsync()
        {
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

            Form webBrowserForm = new Form
            {
                WindowState = FormWindowState.Normal,
                StartPosition = FormStartPosition.CenterScreen,
                Width = 500,
                Height = 700
            };

            WebBrowser webBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                Url = new Uri(url)
            };

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

            Form webBrowserForm = new Form
            {
                WindowState = FormWindowState.Minimized
            };

            WebBrowser webBrowser = new WebBrowser
            {
                Dock = DockStyle.None,
                Url = new Uri(url)
            };

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
    }
}
