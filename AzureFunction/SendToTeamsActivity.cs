// Default URL for triggering event grid function in the local environment.
// http://localhost:7071/runtime/webhooks/EventGrid?functionName={functionname}
using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Azure.EventGrid.Models;
using Microsoft.Azure.WebJobs.Extensions.EventGrid;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Collections.Generic;
using KeyValuePair = Microsoft.Graph.KeyValuePair;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace TeamsAzDemo
{
    public static class SendToTeamsActivity
    {
        [FunctionName("SendToTeamsActivity")]
        public static async Task Run([EventGridTrigger] EventGridEvent eventGridEvent, ILogger log)
        {
            log.LogInformation(eventGridEvent.Data.ToString());

            JObject dataObject = eventGridEvent.Data as JObject;
            ActivityDetails details = dataObject.ToObject<ActivityDetails>();

            string clientId = Environment.GetEnvironmentVariable("ClientId");
            string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");
            string tenantId = Environment.GetEnvironmentVariable("TenantId");

            string authority = $"https://login.microsoftonline.com/{tenantId}/v2.0";
            string userId = details.userId;
            string taskId = details.taskId;
            string notificationUrl = details.notificationUrl;

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
              .WithClientSecret(clientSecret)
              .WithAuthority(authority)
              .Build();

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            MSALAuthenticationProvider authenticationProvider = new MSALAuthenticationProvider(cca, scopes.ToArray());
            GraphServiceClient graphServiceClient = new GraphServiceClient(authenticationProvider);

            var topic = new TeamworkActivityTopic
            {
                Source = TeamworkActivityTopicSource.Text,
                Value = "New Task Created",
                WebUrl = notificationUrl
            };

            var activityType = "taskCreated";

            var previewText = new ItemBody
            {
                Content = "A new task has been created for you"
            };

            var templateParameters = new List<KeyValuePair>()
            {
                new KeyValuePair
                {
                    Name = "taskId",
                    Value = taskId
                }
            };

            await graphServiceClient.Users[userId].Teamwork
                .SendActivityNotification(topic, activityType, null, previewText, templateParameters)
                .Request()
                .PostAsync();
        }
    }
}
