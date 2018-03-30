using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ServiceBus.Messaging;

namespace SpFxGraphFctWebHook
{
    public static class HandleWebHook
    {
        const string validationKey = "validationToken";
        [FunctionName(nameof(HandleWebHook))]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req,
            //[ServiceBus("emails", AccessRights.Manage, Connection = "ServiceBusConnection")] IAsyncCollector<string> collector,
            TraceWriter log, IBinder binder)
        {
            log.Info("C# HTTP trigger function processed a request.");
            var qs = req.GetQueryNameValuePairs();
            var valValue = qs.FirstOrDefault(x => x.Key.Equals(validationKey, StringComparison.CurrentCultureIgnoreCase));
            if (!string.IsNullOrEmpty(valValue.Value))
                return req.CreateResponse(HttpStatusCode.OK, valValue.Value);
            else
            {
                var serviceBusQueueAttribute = new ServiceBusAttribute("emails", AccessRights.Manage)
                {
                    Connection = "ServiceBusConnection"
                };
                var outputMessages = await binder.BindAsync<IAsyncCollector<string>>(serviceBusQueueAttribute);


                var body = await req.Content.ReadAsStringAsync();
                await outputMessages.AddAsync(body);
                return req.CreateResponse(HttpStatusCode.Accepted);
            }
        }
    }
}
