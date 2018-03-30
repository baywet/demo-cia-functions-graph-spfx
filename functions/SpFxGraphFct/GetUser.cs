using System;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Linq;

namespace SpFxGraphFct
{
    public static class GetUser
    {
        [FunctionName(nameof(GetUser))]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            var accessToken = req.Headers.Authorization.Parameter ?? req.Headers.FirstOrDefault(x => x.Key.Equals("X-relaytoken", StringComparison.InvariantCultureIgnoreCase)).Value?.FirstOrDefault();
            var client = await getClient(accessToken);
            var messages = await client.Me.Messages.Request().GetAsync();
            var overAllScore = 0d;
            var numberOfMarks = 0;
            foreach (var message in messages)
            {
                try
                {
                    var extension = await client.Me.Messages[message.Id].Extensions["com.baywet.happy"].Request().GetAsync();
                    var strScore = extension.AdditionalData["score"] as string;
                    if (!string.IsNullOrEmpty(strScore))
                    {
                        var score = double.Parse(strScore);
                        overAllScore += score;
                        numberOfMarks++;
                    }
                }
                catch (ServiceException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
                {// this email is not scored
                }
            }
            var average = numberOfMarks == 0 ? 0d : overAllScore / numberOfMarks;
            return req.CreateResponse(HttpStatusCode.OK, new { score = average }, JsonMediaTypeFormatter.DefaultMediaType);
        }
        private static async Task<GraphServiceClient> getClient(string accessToken)
        {
            try
            {
                var applicationId = ConfigurationManager.AppSettings["WEBSITE_AUTH_CLIENT_ID"];
                var applicationSecret = ConfigurationManager.AppSettings["WEBSITE_AUTH_CLIENT_SECRET"];
                var tenant = ConfigurationManager.AppSettings["Tenant"];
                var cac = new ClientCredential(applicationId, applicationSecret);
                var ua = new UserAssertion(accessToken);
                var authContext = new AuthenticationContext($"https://login.microsoftonline.com/{tenant}", false);
                var authenticationResult = await authContext.AcquireTokenAsync("https://graph.microsoft.com", cac, ua);
                return new GraphServiceClient(new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);
                                await Task.FromResult(true);
                            }
                        ));
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.TraceError("Error during Graph Authentication: " + e.ToString());
                return null;
            }
        }
    }
}
