using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Azure.Identity;
using System.Collections.Generic;
using GraphUser = Microsoft.Graph.User;

namespace DirectReports
{
    public static class DirectReports
    {
        [FunctionName("DirectReports")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            string userPrincipalName = req.Query["alias"];
            if (string.IsNullOrEmpty(userPrincipalName)){
                return new BadRequestObjectResult("please specify an alias");
            }

            var results = await GetDirects(userPrincipalName, log);       

            return new OkObjectResult(results);
        }

        public static async Task<IList<User>> GetDirects(string userPrincipalName, ILogger log)
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
            var manager = (GraphUser)user.Manager;
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
                var direct = (GraphUser)report;
                results.AddRange(await GetDirects(direct.UserPrincipalName, log));
            }

            return results;
        }
    }
}
