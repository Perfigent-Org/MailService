using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace MailService.OAuthOutlook
{
    internal static class UserProfileUtils
    {
        const string LiveGetProfileUri = "https://apis.live.net/v5.0/me?access_token=";

        protected class ExtendedMicrosoftClientUserData
        {
            public string FirstName { get; set; }
            public string Gender { get; set; }
            public string Id { get; set; }
            public string LastName { get; set; }
            public Uri Link { get; set; }
            public string Name { get; set; }
            public Emails Emails { get; set; }
        }

        protected class Emails
        {
            public string Preferred { get; set; }
            public string Account { get; set; }
            public string Personal { get; set; }
            public string Business { get; set; }
        }

        private static readonly string[] UriRfc3986CharsToEscape = new string[] { "!", "*", "'", "(", ")" };
        private static string EscapeUriDataStringRfc3986(string value)
        {
            StringBuilder escaped = new StringBuilder(Uri.EscapeDataString(value));

            // Upgrade the escaping to RFC 3986, if necessary.
            for (int i = 0; i < UriRfc3986CharsToEscape.Length; i++)
            {
                escaped.Replace(UriRfc3986CharsToEscape[i], Uri.HexEscape(UriRfc3986CharsToEscape[i][0]));
            }

            // Return the fully-RFC3986-escaped string.
            return escaped.ToString();
        }

        // Inspired by http://answer.techwikihow.com/154458/getting-email-oauth-authentication-microsoft.html
        // Be sure to have "wl.emails" in the requested scopes if you're using this method.
        public static async Task<IDictionary<string, string>> GetUserDataAsync(string accessToken)
        {
            ExtendedMicrosoftClientUserData graph;
            WebRequest request =
                WebRequest.Create(LiveGetProfileUri + EscapeUriDataStringRfc3986(accessToken));

            WebResponse response = null;
            try
            {
                try
                {
                    response = (WebResponse)await request.GetResponseAsync();
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
                    using (StreamReader sr = new StreamReader(responseStream))
                    {
                        string data = await sr.ReadToEndAsync();
                        graph = JsonConvert.DeserializeObject<ExtendedMicrosoftClientUserData>(data);
                    }
                }
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }

            Dictionary<string, string> userData = new Dictionary<string, string>();
            userData.Add("id", graph.Id);
            userData.Add("username", graph.Name);
            userData.Add("name", graph.Name);
            userData.Add("link", graph.Link == null ? null : graph.Link.AbsoluteUri);
            userData.Add("gender", graph.Gender);
            userData.Add("firstname", graph.FirstName);
            userData.Add("lastname", graph.LastName);
            userData.Add("email", graph.Emails.Preferred);
            return userData;
        }
    }
}
