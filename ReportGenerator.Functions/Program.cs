using Azure.Identity;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services;
using ReportGenerator.Business.Services.Interfaces;
using ReportGenerator.Data.Repositories;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Data;

var host = Host.CreateDefaultBuilder(args)
    .ConfigureFunctionsWebApplication((context, builder) =>
    {
        
    })
    .ConfigureAppConfiguration((context, config) =>
    {
        // Load configuration files and secrets
        config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
        config.AddJsonFile("host.json", optional: true, reloadOnChange: true);
        config.AddUserSecrets<Program>(optional: true);
        //config.AddEnvironmentVariables();
    })
    .ConfigureServices((context, services) =>
    {
        var configuration = context.Configuration;

        services.Configure<GraphSettings>(configuration.GetSection("GraphSettings"));
        services.Configure<SharePointSettings>(configuration.GetSection("SharePointSettings"));
        services.Configure<QueueSettings>(configuration.GetSection("QueueSettings"));

        var graphSettings = configuration.GetSection("GraphSettings").Get<GraphSettings>()
            ?? throw new InvalidOperationException("GraphSettings section is missing.");

        if (string.IsNullOrWhiteSpace(graphSettings.TenantId) ||
            string.IsNullOrWhiteSpace(graphSettings.ClientId) ||
            string.IsNullOrWhiteSpace(graphSettings.ClientSecret))
            throw new InvalidOperationException("GraphSettings has missing required fields.");

        var sharePointSettings = configuration.GetSection("SharePointSettings").Get<SharePointSettings>()
            ?? throw new InvalidOperationException("SharePointSettings section is missing.");

        if (string.IsNullOrWhiteSpace(sharePointSettings.SiteUrl) ||
            string.IsNullOrWhiteSpace(sharePointSettings.CNReportsBasePath) ||
            string.IsNullOrWhiteSpace(sharePointSettings.CNMasterTemplatePath) ||
            string.IsNullOrWhiteSpace(sharePointSettings.DraftsFolderName))
            throw new InvalidOperationException("SharePointSettings has missing required fields.");

        var queueSettings = configuration.GetSection("QueueSettings").Get<QueueSettings>()
            ?? throw new InvalidOperationException("QueueSettings section is missing.");

        if (string.IsNullOrWhiteSpace(queueSettings.ReportFolderQueue))
            throw new InvalidOperationException("QueueSettings:ReportFolderQueue is missing.");

        // Get connection string from config or environment
        string connectionString = configuration["SqlConnectionString"]
            ?? throw new InvalidOperationException("SqlConnectionString not set.");

        services.AddScoped<IDbConnection>(_ => new SqlConnection(connectionString));

        var credential = new ClientSecretCredential(
            graphSettings.TenantId,
            graphSettings.ClientId,
            graphSettings.ClientSecret
        );

        var scopes = new[] { "https://graph.microsoft.com/.default" };

        services.AddSingleton(new GraphServiceClient(credential, scopes));

        services.AddSingleton(sp =>
        {
            string storageConnectionString = configuration["AzureWebJobsStorage"]
                ?? throw new InvalidOperationException("AzureWebJobsStorage not configured.");

            var queueClient = new QueueClient(storageConnectionString, queueSettings.ReportFolderQueue);

            queueClient.CreateIfNotExists();

            return queueClient;
        });

        // Register dependencies
        services.AddScoped<IReportRepository, ReportRepository>();
        services.AddScoped<IReportService, ReportService>();
        services.AddScoped<ISharePointService, SharePointService>();
        services.AddScoped<IProjectFolderService, ProjectFolderService>();
        services.AddScoped<IBridgeExcelProcessor, BridgeExcelProcessor>();
        services.AddScoped<IExcelExtractionService, ExcelExtractionService>();
        services.AddScoped<IAlternativeOptionsRepository, AlternativeOptionsRepository>();
        services.AddScoped<IBridgeCharacteristicsRepository, BridgeCharacteristicsRepository>();

        services.AddHttpClient();

        // Add Application Insights telemetry for worker
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();
    })
    .ConfigureLogging((context, logging) =>
    {
        // Load logging levels from configuration (host.json / appsettings.json)
        logging.AddConfiguration(context.Configuration.GetSection("Logging"));

        // Add Application Insights and Console loggers
        logging.AddApplicationInsights();
        logging.AddConsole();

        // Remove the default AI logging rule so that host.json settings are honored
        logging.Services.Configure<LoggerFilterOptions>(options =>
        {
            var defaultRule = options.Rules.FirstOrDefault(rule =>
                rule.ProviderName ==
                "Microsoft.Extensions.Logging.ApplicationInsights.ApplicationInsightsLoggerProvider");
            if (defaultRule is not null)
            {
                options.Rules.Remove(defaultRule);
            }
        });
    })
    .Build();

await host.RunAsync();
