using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Azure.Identity;
using GUser = Microsoft.Graph.User;

namespace DirectReports
{
    internal class DirectReportsDurable
    {
        [FunctionName("DirectReportsDurable")]
        public static async Task<List<User>> RunOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context)
        {
            var outputs = new List<User>();
            var rootUser = context.GetInput<string>();
            outputs.AddRange(await context.CallActivityAsync<IList<User>>("DirectReports_GetDirects", rootUser));
            return outputs;
        }        

        [FunctionName("DirectReports_GetDirects")]
        public static async Task<IList<User>> GetDirects([ActivityTrigger] string userPrincipalName, ILogger log)
        {
            log.LogInformation($"Getting direct reports for {userPrincipalName}");

            var results = new List<User>();

            var credential = new ClientSecretCredential(System.Environment.GetEnvironmentVariable("TenantId"), 
                System.Environment.GetEnvironmentVariable("ClientId"), 
                System.Environment.GetEnvironmentVariable("ClientSecret")
            ); //new DefaultAzureCredential();

            var graphClient = new GraphServiceClient(credential);
            var user = await graphClient.Users[userPrincipalName]
                //.DirectReports
                .Request()
                .Select("id, displayName, userPrincipalName")
                .Expand("manager")
                .GetAsync();
            var manager = (GUser)user.Manager;
            //var blobClient = new BlobClient(new Uri("https://myaccount.blob.core.windows.net/mycontainer/myblob"), credential);            

            results.Add(new User
            {
                UserId = user.Id,
                UserPrincipalName = user.UserPrincipalName,
                DirectManager = manager?.UserPrincipalName ?? ""
            });

            var directReports = await graphClient.Users[userPrincipalName]
                .DirectReports
                .Request()
                .GetAsync();

            log.LogInformation($"{userPrincipalName} has {directReports.Count} direct reports.");

            foreach(var report in directReports)
            {
                var direct = (GUser)report;
                results.AddRange(await GetDirects(direct.UserPrincipalName, log));
            }

            return results;
        }

        [FunctionName("DirectReports_HttpStart")]
        public static async Task<HttpResponseMessage> HttpStart(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequestMessage req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            // Function input comes from the request content.
            string instanceId = await starter.StartNewAsync("DirectReportsDurable", null, "pjirsa@36mphdev.onmicrosoft.com");

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            return starter.CreateCheckStatusResponse(req, instanceId);
        }
    }
}