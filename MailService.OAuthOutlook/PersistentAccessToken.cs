using Newtonsoft.Json;

namespace MailService.OAuthOutlook
{
    [JsonObject(MemberSerialization.OptIn)]
    internal class PersistentAccessToken
    {
        [JsonProperty("expires_at")]
        public string ExpiresAt { get; set; }

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; set; }
    }
}
