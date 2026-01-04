using Azure.Storage.Queues.Models;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using System.Text.Json;

namespace ReportGenerator.Functions.Report;

public class ReportFolderQueueFunction
{
    #region Private Fields

    private readonly IProjectFolderService _projectFolderService;
    private readonly ILogger<ReportFolderQueueFunction> _logger;


    #endregion

    #region Constructor

    public ReportFolderQueueFunction(
        IProjectFolderService projectFolderService,
        ILogger<ReportFolderQueueFunction> logger)
    {
        _projectFolderService = projectFolderService;
        _logger = logger;
    }

    #endregion

    #region Function Methods

    [Function(nameof(ReportFolderQueueFunction))]
    public async Task Run([QueueTrigger("report-folder-queue", Connection = "AzureWebJobsStorage")] QueueMessage message)
    {
        try
        {
            _logger.LogInformation("Queue trigger received message: {msg}", message.MessageText);

            var request = JsonSerializer.Deserialize<ReportFolderRequest>(message.MessageText);

            if (request == null || request.ReportId == Guid.Empty)
            {
                _logger.LogWarning("Invalid queue message format.");
                return;
            }

            _logger.LogInformation("Creating folder structure for ReportId: {id}", request.ReportId);

            await _projectFolderService.CreateFullFolderStructureAsync(request.ReportId);

            _logger.LogInformation("Folder structure created successfully for {id}", request.ReportId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing queue message.");
            throw;
        }
    }

    #endregion
}