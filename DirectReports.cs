using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using GraphUser = Microsoft.Graph.User;
using Azure.Storage.Blobs;
using Newtonsoft.Json;
using System.IO;
using System.Text;

namespace DirectReports
{
    public class DirectReports
    {
        private readonly GraphServiceClient _graphClient;
        private readonly BlobServiceClient _blobClient;
        private readonly ILogger<DirectReports> _logger;

        public DirectReports(GraphServiceClient graphClient, BlobServiceClient blobClient, ILogger<DirectReports> logger)
        {
            _graphClient = graphClient;
            _blobClient = blobClient;
            _logger = logger;
        }

        [FunctionName("DirectReportsTimerTrigger")]
        public async Task DirectReportsTimerTrigger(
            [TimerTrigger("0 30 1 * * *", RunOnStartup = false)]TimerInfo timer)
        {
            _logger.LogInformation("Executing time trigger for DirectReports");
            string topLevelAlias = System.Environment.GetEnvironmentVariable("TopLevelAlias");

            var results = await GetDirects(topLevelAlias);

            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(results)));
            await _blobClient
                .GetBlobContainerClient("directreports")
                .UploadBlobAsync($"directreports-{System.DateTime.Now.ToString("yyyyMMdd")}.json", stream);
        }

        [FunctionName("DirectReports")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req)
        {
            string userPrincipalName = req.Query["alias"];
            if (string.IsNullOrEmpty(userPrincipalName)){
                return new BadRequestObjectResult("please specify an alias");
            }

            var results = await GetDirects(userPrincipalName);       

            return new OkObjectResult(results);
        }

        private async Task<IList<User>> GetDirects(string userPrincipalName)
        {
            _logger.LogInformation($"Getting direct reports for {userPrincipalName}");
            // Initialize results list
            var results = new List<User>();            

            // Create content wrapper for batch requests
            var batchRequest = new BatchRequestContent();

            // Create request for user profile
            var userRequest = _graphClient.Users[userPrincipalName]
                .Request()
                .Select("id, displayName, userPrincipalName")
                .Expand("manager");

            // Create request for direct reports
            var directReportsRequest = _graphClient.Users[userPrincipalName]
                .DirectReports
                .Request();
            
            // Add requests to batch
            var userReqId = batchRequest.AddBatchRequestStep(userRequest);
            var directReportsReqId = batchRequest.AddBatchRequestStep(directReportsRequest);
            
            // Post batch request
            var batchResponse = await _graphClient.Batch.Request().PostAsync(batchRequest);

            // Access results and build results list
            try 
            {
                var user = await batchResponse.GetResponseByIdAsync<GraphUser>(userReqId);
                var directReports = await batchResponse.GetResponseByIdAsync<UserDirectReportsCollectionWithReferencesResponse>(directReportsReqId);

                _logger.LogInformation($"{userPrincipalName} has {directReports.Value.Count} direct reports.");

                var manager = (GraphUser)user.Manager;
                results.Add(new User
                {
                    UserId = user.Id,
                    DisplayName = user.DisplayName,
                    UserPrincipalName = user.UserPrincipalName,
                    DirectManager = manager?.UserPrincipalName ?? ""
                });

                foreach(var report in directReports.Value)
                {
                    var direct = (GraphUser)report;
                    results.AddRange(await GetDirects(direct.UserPrincipalName));
                }
            }
            catch (ServiceException ex)
            {
                _logger.LogError(ex, $"Error getting direct reports for {userPrincipalName}");
            }           

            return results;
        }
    }
}
