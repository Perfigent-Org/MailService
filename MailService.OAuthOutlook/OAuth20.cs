using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace MailService.OAuthOutlook
{
    public class OAuth20
    {
        private static string _storedRefreshToken = null;
        private static CodeGrantOauth _oauth = null;
        private static string tokenFile = "token.json";

        public async static Task<KeyValuePair<string, string>> Login(string clientId)
        {
            PersistentAccessToken tokensFromStorage = null;

            Console.WriteLine("Full path to the token file: " + Path.GetFullPath(tokenFile));

            // You should delete this token.json in case if you're changing ClientId.
            if (File.Exists(tokenFile))
            {
                string json = File.ReadAllText(tokenFile);
                tokensFromStorage = JsonConvert.DeserializeObject(json, typeof(PersistentAccessToken)) as PersistentAccessToken;
                _storedRefreshToken = tokensFromStorage.RefreshToken;
            }

            bool tokenNeedsRefresh = true;
            if (tokensFromStorage != null)
            {
                DateTime expiresAtUtc = DateTime.Parse(tokensFromStorage.ExpiresAt).ToUniversalTime();

                // Uncommment this to make the token "expire" faster if you want to test the code which refreshes the token.
                // expiresAtUtc = expiresAtUtc.AddMinutes(-59);

                if (DateTime.UtcNow.AddMinutes(1) < expiresAtUtc)
                {
                    tokenNeedsRefresh = false;
                }
            }

            if (tokenNeedsRefresh)
            {
                _oauth = await GetOauthTokensAsync(_storedRefreshToken, clientId);

                string expiresAt = DateTime.UtcNow.AddSeconds(_oauth.Expiration).ToString("o", CultureInfo.InvariantCulture);

                tokensFromStorage = new PersistentAccessToken() { AccessToken = _oauth.AccessToken, RefreshToken = _oauth.RefreshToken, ExpiresAt = expiresAt };

                string json = JsonConvert.SerializeObject(tokensFromStorage);
                File.WriteAllText(tokenFile, json);
            }

            IDictionary<string, string> userData = await UserProfileUtils.GetUserDataAsync(tokensFromStorage.AccessToken);

            return new KeyValuePair<string, string>(userData["email"], tokensFromStorage.AccessToken);
        }

        private static DateTime _tokenExpiration;
        private static async Task<CodeGrantOauth> GetOauthTokensAsync(string refreshToken, string clientId)
        {
            CodeGrantOauth auth = new CodeGrantOauth(clientId);

            if (string.IsNullOrEmpty(refreshToken))
            {
                await auth.GetAccessTokenAsync();
            }
            else
            {
                await auth.RefreshAccessTokenAsync(refreshToken);

                // Refresh tokens can become invalid for several reasons
                // such as the user's password changed.
                if (!string.IsNullOrEmpty(auth.Error))
                {
                    auth = await GetOauthTokensAsync(null, clientId);
                }
            }

            if (!string.IsNullOrEmpty(auth.Error))
            {
                throw new Exception(auth.Error);
            }
            else
            {
                _storedRefreshToken = auth.RefreshToken;
                _tokenExpiration = DateTime.Now.AddSeconds(auth.Expiration);
            }

            return auth;
        }

        public static async Task Logout(string clientId)
        {
            try
            {
                if (File.Exists(tokenFile)) File.Delete(tokenFile);
                if (_oauth == null) _oauth = new CodeGrantOauth(clientId);
                if (!await _oauth.Logout()) throw new Exception("Please contact administrator");
                Console.WriteLine("User has signed-out");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error signing-out user: {ex.Message}");
            }
        }
    }
}
