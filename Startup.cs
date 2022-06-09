using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Azure.WebJobs.Host.Bindings;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Azure.Identity;
using Azure.Storage.Blobs;

[assembly: FunctionsStartup(typeof(AuthFuncDemo.Startup))]

namespace AuthFuncDemo
{
    public class Startup : FunctionsStartup
    {
        public Startup()
        {
        }

        IConfiguration Configuration { get; set; }

        public override void Configure(IFunctionsHostBuilder builder)
        {
            // Get the azure function application directory. 'C:\whatever' for local and 'd:\home\whatever' for Azure
            var executionContextOptions = builder.Services.BuildServiceProvider()
                .GetService<IOptions<ExecutionContextOptions>>().Value;

            var currentDirectory = executionContextOptions.AppDirectory;

            // Get the original configuration provider from the Azure Function
            var configuration = builder.Services.BuildServiceProvider().GetService<IConfiguration>();

            // Create a new IConfigurationRoot and add our configuration along with Azure's original configuration 
            Configuration = new ConfigurationBuilder()
                .SetBasePath(currentDirectory)
                .AddConfiguration(configuration) // Add the original function configuration 
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();

            // Replace the Azure Function configuration with our new one
            builder.Services.AddSingleton(Configuration);

            ConfigureServices(builder.Services);
        }

        private void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<GraphServiceClient>(o => {
                var credential = new ClientSecretCredential(
                    System.Environment.GetEnvironmentVariable("TenantId"),
                    System.Environment.GetEnvironmentVariable("ClientId"),
                    System.Environment.GetEnvironmentVariable("ClientSecret"));
                return new GraphServiceClient(credential);
            });

            services.AddSingleton<BlobServiceClient>(o => {
                string blobConnectionString = System.Environment.GetEnvironmentVariable("AzureWebJobsStorage");
                return new BlobServiceClient(blobConnectionString);
            });
        }
    }
}