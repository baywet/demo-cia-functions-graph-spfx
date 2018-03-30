using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ServiceBus.Messaging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Net.Http;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;

namespace SpFxGraphFct
{
    public static class ScoreEmail
    {
        [FunctionName("ScoreEmail")]
        public static async Task Run([ServiceBusTrigger("emails", AccessRights.Manage, Connection = "ServiceBusConnection")]string myQueueItem, TraceWriter log)
        {
            log.Info($"C# ServiceBus queue trigger function processed message: {myQueueItem}");
            var notification = JsonConvert.DeserializeObject<notificationPoco>(myQueueItem);
            var client = await getClient();
            foreach (var item in notification.Value)
            {
                var message = await client.Users[item.UserId].Messages[item.MessageId].Request().GetAsync();
                var score = await GetScore(message.Body.Content);
                await SaveScore(client, item, score);
            }
            log.Info($"{notification.Value.First().Resource}");

        }
        private static async Task SaveScore(GraphServiceClient client, notificationItem item, double? score)
        {
            var extension = new OpenTypeExtension
            {
                ODataType = "Microsoft.Graph.OpenTypeExtension",
                ExtensionName = "com.baywet.happy",
            };
            extension.AdditionalData = new Dictionary<string, object>
                {
                    { "score", score.Value.ToString() }
                };
            await client.Users[item.UserId].Messages[item.MessageId].Extensions.Request().AddAsync(extension);
        }
        private static string ScrubHtml(string value)
        {
            var step1 = Regex.Replace(value, @"<[^>]+>|&nbsp;", "").Trim();
            var step2 = Regex.Replace(step1, @"\s{2,}", " ");
            return step2;
        }
        private static async Task<double?> GetScore(string text)
        {
            using (ITextAnalyticsAPI client = new TextAnalyticsAPI
            {
                SubscriptionKey = ConfigurationManager.AppSettings["CognitiveKey"],
                AzureRegion = AzureRegions.Eastus2

            })
            {
                var result = await client.SentimentAsync(
                    new MultiLanguageBatchInput(
                        new List<MultiLanguageInput>()
                        {
                          new MultiLanguageInput(id: "1", text: ScrubHtml( text)),
                        }));
                return result.Documents.First().Score;
            }
        }
        private static async Task<GraphServiceClient> getClient()
        {
            try
            {
                var applicationId = ConfigurationManager.AppSettings["WEBSITE_AUTH_CLIENT_ID"];
                var applicationSecret = ConfigurationManager.AppSettings["WEBSITE_AUTH_CLIENT_SECRET"];
                var tenant = ConfigurationManager.AppSettings["Tenant"];
                var cac = new ClientCredential(applicationId, applicationSecret);
                var authContext = new AuthenticationContext($"https://login.microsoftonline.com/{tenant}", false);
                var authenticationResult = await authContext.AcquireTokenAsync("https://graph.microsoft.com", cac);
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
        public class notificationPoco
        {
            public List<notificationItem> Value { get; set; }
        }
        public class notificationItem
        {
            public Guid SubscriptionId { get; set; }
            public DateTime SubscriptionExpirationDateTime { get; set; }
            public string ClientState { get; set; }
            public string ChangeType { get; set; }
            public string Resource { get; set; }
            private static string userIdKey = "userId";
            private static string messageIdKey = "messageId";
            private static Regex extractingRegex = new Regex($@"Users/(?<{userIdKey}>[\w-]*)/Messages/(?<{messageIdKey}>[\w_=]*)");
            public string UserId
            {
                get
                {
                    return extractingRegex.Match(Resource).Groups[userIdKey].Value;
                }
            }
            public string MessageId
            {
                get
                {
                    return extractingRegex.Match(Resource).Groups[messageIdKey].Value;
                }
            }
        }
    }
}
