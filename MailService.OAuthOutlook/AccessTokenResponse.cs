using Newtonsoft.Json;

namespace MailService.OAuthOutlook
{
    [JsonObject(MemberSerialization.OptIn)]
    internal class AccessTokenResponse
    {
        [JsonProperty("expires_in")]
        public int Expiration { get; set; }

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; set; }
    }
}
