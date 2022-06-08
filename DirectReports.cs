using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using GraphUser = Microsoft.Graph.User;

namespace DirectReports
{
    public class DirectReports
    {
        private GraphServiceClient _graphClient;
        private ILogger<DirectReports> _logger;

        public DirectReports(GraphServiceClient graphClient, ILogger<DirectReports> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
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

            var results = new List<User>();            
            var user = await _graphClient.Users[userPrincipalName]
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

            var directReports = await _graphClient.Users[userPrincipalName]
                .DirectReports
                .Request()
                .GetAsync();

            _logger.LogInformation($"{userPrincipalName} has {directReports.Count} direct reports.");

            foreach(var report in directReports)
            {
                var direct = (GraphUser)report;
                results.AddRange(await GetDirects(direct.UserPrincipalName));
            }

            return results;
        }
    }
}
