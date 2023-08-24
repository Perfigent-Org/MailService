using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MailService.OAuthOffice365
{
    public static class OAuth20
    {
        private static string apiEndpoint = "https://outlook.office.com/api/v2.0/me";
        private static string[] scopes = new string[] { "offline_access", "https://outlook.office365.com/IMAP.AccessAsUser.All", "https://outlook.office365.com/SMTP.Send" };

        //private static string apiEndpoint = "https://graph.microsoft.com/v1.0/me";
        //private static string[] scopes = new string[] { "offline_access", "https://graph.microsoft.com/IMAP.AccessAsUser.All", "https://graph.microsoft.com/SMTP.Send" };

        public async static Task<KeyValuePair<string, string>> Login()
        {
            AuthenticationResult authResult = null;
            var app = Client.PublicClientApp;

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    try { await Logout(); } catch { }
                    throw new Exception($"Error Acquiring Token:{System.Environment.NewLine}{msalex}");
                }
            }
            catch (Exception ex)
            {
                try { await Logout(); } catch { }
                throw new Exception($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
            }

            // Note that you don't need to call this method if you're just looking for the email address, you can use authResult.Account.Username for that.
            // If you also need display name or any other details, then using this method is justified.
            //var userData = await GetUserData(authResult);

            return new KeyValuePair<string, string>(authResult.Account.Username, authResult.AccessToken);
        }

        private static async Task<UserData> GetUserData(AuthenticationResult authResult)
        {
            if (authResult != null)
            {
                WithExceptionOrError<UserData> userData = await GetUserDataHttpContentWithToken(apiEndpoint, authResult.AccessToken);
                if (userData.ErrorMessage == null)
                {
                    Console.WriteLine($"Email address: {userData.Value.EmailAddress}, Display name: {userData.Value.DisplayName}");
                    return userData.Value;
                }
                else
                {
                    throw new Exception(userData.ErrorMessage);
                }
            }

            return null;
        }

        // Note that even if no exception occurred, it can still return an object with an error message rather than the filled profile. This is how Outlook REST API works.
        private static async Task<WithExceptionOrError<UserData>> GetUserDataHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;

            var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);

            //Add the token in Authorization header
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            try
            {
                response = await httpClient.SendAsync(request);
                string content = await response.Content.ReadAsStringAsync();
                UserData data = JsonConvert.DeserializeObject<UserData>(content);
                return new WithExceptionOrError<UserData>(data, data.Error?.Message);
            }
            catch (Exception ex)
            {
                return new WithExceptionOrError<UserData>(null, ex);
            }
        }

        public static async Task Logout()
        {
            var accounts = await Client.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await Client.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    Console.WriteLine("User has signed-out");
                }
                catch (MsalException ex)
                {
                    throw new Exception($"Error signing-out user: {ex.Message}");
                }
            }
        }
    }
}
