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

namespace DirectReports
{
    public class DirectReports
    {
        [FunctionName("DirectReports")]
        public static async Task<List<User>> RunOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context)
        {
            var outputs = new List<User>();
            var rootUser = context.GetInput<string>();
            var names = new []{ rootUser };

            foreach (var name in names)
            {
                outputs.AddRange(await context.CallActivityAsync<IList<User>>("DirectReports_GetDirects", name));
            }
            return outputs;
        }        

        [FunctionName("DirectReports_GetDirects")]
        public static async Task<IList<User>> GetDirects([ActivityTrigger] string userPrincipalName, ILogger log)
        {
            log.LogInformation($"Getting direct reports for {userPrincipalName}");

            var credential = new ClientSecretCredential(System.Environment.GetEnvironmentVariable("TenantId"), 
                System.Environment.GetEnvironmentVariable("ClientId"), 
                System.Environment.GetEnvironmentVariable("ClientSecret")
            ); //new DefaultAzureCredential();

            var graphClient = new GraphServiceClient(credential);
            var results = await graphClient.Users[userPrincipalName].DirectReports.Request().Select("id, displayName, userPrincipalName").GetAsync();
            //var blobClient = new BlobClient(new Uri("https://myaccount.blob.core.windows.net/mycontainer/myblob"), credential);

            log.LogInformation($"{userPrincipalName} has {results.Count} direct reports.");

            var user = new User 
            {
                UserId = "123",
                UserPrincipalName = "dave@acme.com",
                DirectManager = userPrincipalName
            };

            return new []{ user };
        }

        [FunctionName("DirectReports_HttpStart")]
        public static async Task<HttpResponseMessage> HttpStart(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequestMessage req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            // Function input comes from the request content.
            string instanceId = await starter.StartNewAsync("DirectReports", null, "pjirsa@36mphdev.onmicrosoft.com");

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            return starter.CreateCheckStatusResponse(req, instanceId);
        }
    }
}