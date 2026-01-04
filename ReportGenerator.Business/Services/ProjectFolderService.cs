using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ReportGenerator.Business.Extensions;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Text;
using System.Text.Json;

namespace ReportGenerator.Business.Services
{
    public class ProjectFolderService : IProjectFolderService
    {
        #region Private Fields

        private readonly QueueClient _queueClient;
        private readonly ISharePointService _sharePointService;
        private readonly IReportRepository _reportRepository;
        private readonly SharePointSettings _sharePointSettings;
        private readonly ILogger<ProjectFolderService> _logger;

        #endregion

        #region Constructor

        public ProjectFolderService(
            QueueClient queueClient,
            ISharePointService sharePointService,
            IReportRepository reportRepository,
            IOptions<SharePointSettings> sharePointOptions,
            ILogger<ProjectFolderService> logger)
        {
            _queueClient = queueClient;
            _sharePointService = sharePointService;
            _reportRepository = reportRepository;
            _sharePointSettings = sharePointOptions.Value ?? throw new ArgumentNullException(nameof(sharePointOptions));
            _logger = logger;
        }

        #endregion

        #region Public Methods

        public async Task EnqueueFolderCreationAsync(Guid reportId)
        {
            if (reportId == Guid.Empty)
                throw new ArgumentException("ReportId cannot be empty.");

            var payload = new { ReportId = reportId };

            var json = JsonSerializer.Serialize(payload);

            string base64Message = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

            await _queueClient.SendMessageAsync(base64Message);

            _logger.LogInformation("Queued folder creation request for Report {ReportId}", reportId);
        }

        public async Task CreateFullFolderStructureAsync(Guid reportId)
        {
            var baseRoot = _sharePointSettings.CNReportsBasePath
                ?? throw new InvalidOperationException("SharePointBasePath not configured.");

            var report = await _reportRepository.GetBasicReportInfoAsync(reportId)
                  ?? throw new InvalidOperationException($"Report {reportId} not found.");

            string projectRoot = $"{baseRoot}/{report.ProjectNumber}";

            _logger.LogInformation("Creating full folder structure under: {ProjectRoot}", projectRoot);

            var fullPaths = ProjectFolderBlueprint.CN
                .ToPaths()
                .Select(rel => $"{projectRoot}/{rel}")
                .ToList();

            foreach (var path in fullPaths)
            {
                await _sharePointService.EnsureFolderHierarchyAsync(path);

                _logger.LogInformation("Created folder: {Path}", path);
            }

            _logger.LogInformation("Folder structure completed for project {ProjectNumber}", report.ProjectNumber);
        }

        #endregion
    }
}
