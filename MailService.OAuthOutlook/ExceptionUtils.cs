using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace MailService.OAuthOutlook
{
    internal static class ExceptionUtils
    {
        // By default, WebExceptions are not very informative. Let's make them better.
        // https://stackoverflow.com/questions/11828843/webexception-how-to-get-whole-response-with-a-body
        public static async Task ThrowWithBodyAsync(WebException wex)
        {
            if (wex.Status == WebExceptionStatus.ProtocolError)
            {
                string responseBody;
                try
                {
                    // Get the message body for rethrow with body included
                    responseBody = await new StreamReader(wex.Response.GetResponseStream()).ReadToEndAsync();
                }
                catch (Exception)
                {
                    // In case of failure to get the body just rethrow the original web exception.
                    throw wex;
                }

                //include the body in the message
                throw new WebException(wex.Message + $" Response body: '{responseBody}'", wex, wex.Status, wex.Response);
            }

            // In case of non-protocol errors no body is available anyway, so just rethrow the original web exception.
            throw wex;
        }
    }
}
