using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;

namespace ReportGenerator.Business.Services
{
    public class SharePointService : ISharePointService
    {
        #region Private Fields

        private readonly SharePointSettings _sharePointSettings;
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<SharePointService> _logger;

        #endregion

        #region Constructor

        public SharePointService(
            IOptions<SharePointSettings> sharePointOptions,
            GraphServiceClient graphClient, 
            ILogger<SharePointService> logger)
        {
            _sharePointSettings = sharePointOptions.Value ?? throw new ArgumentNullException(nameof(sharePointOptions));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        }

        #endregion

        #region Public Methods

        public async Task<byte[]> DownloadDocumentAsync(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                throw new ArgumentException("relativePath must be provided.", nameof(relativePath));

            var drive = await GetDriveAsync();

            await using var stream = await _graphClient
                .Drives[drive.Id]
                .Root
                .ItemWithPath(relativePath)
                .Content
                .GetAsync()
                ?? throw new FileNotFoundException($"File not found in SharePoint: {relativePath}");

            using var memory = new MemoryStream();
            await stream.CopyToAsync(memory);
            return memory.ToArray();
        }

        public async Task<DriveItem> UploadContentAndReturnItemAsync(string relativePath, byte[] fileContent)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                throw new ArgumentException("relativePath must be provided.", nameof(relativePath));

            var drive = await GetDriveAsync();

            await EnsureFolderHierarchyAsync(drive.Id!, relativePath, true);

            await using var ms = new MemoryStream(fileContent);

            var uploaded = await _graphClient.Drives[drive.Id]
                .Root
                .ItemWithPath(relativePath)
                .Content
                .PutAsync(ms);

            _logger.LogInformation("Uploaded file to SharePoint: {Url}", uploaded?.WebUrl);

            return uploaded!;
        }

        public async Task EnsureFolderHierarchyAsync(string relativePath, bool isFilePath = false)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                throw new ArgumentException("relativePath must be provided.", nameof(relativePath));

            var drive = await GetDriveAsync();

            await EnsureFolderHierarchyAsync(drive.Id!, relativePath, isFilePath);

            _logger.LogInformation("Ensured folder hierarchy: {Path}", relativePath);
        }

        public async Task<List<string>> GetImagePathsAsync(string projectNumber, string bridgeCode)
        {
            var drive = await GetDriveAsync();

            var basePath = $"GeneratedReports/{projectNumber}/{bridgeCode}/Images";

            var imageList = new List<string>();

            try
            {
                DriveItem? imageFolder = null;

                try
                {
                    imageFolder = await _graphClient
                        .Drives[drive.Id]
                        .Root
                        .ItemWithPath(basePath)
                        .GetAsync();
                }
                catch
                {
                    _logger.LogWarning("Image folder does not exist: {Path}", basePath);
                    return imageList;
                }

                var items = await _graphClient
                    .Drives[drive.Id]
                    .Items[imageFolder!.Id]
                    .Children
                    .GetAsync();

                if (items?.Value != null)
                {
                    foreach (var file in items.Value.Where(i => i.File != null))
                    {
                        if (!string.IsNullOrEmpty(file.Name))
                        {
                            var relativePath = $"{basePath}/{file.Name}";
                            imageList.Add(relativePath);
                        }
                    }
                }

                _logger.LogInformation("Retrieved {Count} images for {Project}/{Bridge}", imageList.Count, projectNumber, bridgeCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving image files from SharePoint path: {Path}", basePath);
                throw;
            }

            return imageList;
        }

        public async Task<List<string>> GetAllImagePathsForReportAsync(string projectNumber)
        {
            var drive = await GetDriveAsync();
            var result = new List<string>();

            foreach (var folder in ImageUploadFolderMap.DistinctFolders)
            {
                var fullPath = $"{_sharePointSettings.CNReportsBasePath}/{projectNumber}/{folder}";

                try
                {
                    var folderItem = await _graphClient
                        .Drives[drive.Id]
                        .Root
                        .ItemWithPath(fullPath)
                        .GetAsync();

                    var children = await _graphClient
                        .Drives[drive.Id]
                        .Items[folderItem.Id]
                        .Children
                        .GetAsync();

                    if (children?.Value != null)
                    {
                        foreach (var file in children.Value.Where(f => f.File != null))
                        {
                            result.Add($"{fullPath}/{file.Name}");
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            _logger.LogInformation("Collected {Count} image(s) for report {ProjectNumber}", result.Count, projectNumber);

            return result;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Returns a Graph client and the default document drive for the configured site.
        /// </summary>
        private async Task<Drive> GetDriveAsync()
        {
            var siteUrl = _sharePointSettings.SiteUrl ?? throw new InvalidOperationException("SharePoint Site URL is not configured.");

            if (!Uri.TryCreate(siteUrl, UriKind.Absolute, out var uri))
                throw new InvalidOperationException($"Invalid SharePointSiteUrl format: {siteUrl}");

            var site = await _graphClient
                .Sites[$"{uri.Host}:{uri.AbsolutePath}"]
                .GetAsync()
                ?? throw new InvalidOperationException("Unable to resolve SharePoint site.");

            var drive = (await _graphClient.Sites[site.Id].Drives.GetAsync())?.Value?.FirstOrDefault()
                ?? throw new InvalidOperationException("No default drive found.");

            return drive;
        }

        /// <summary>
        /// Ensures all folders in a relative path exist
        /// </summary>
        private async Task EnsureFolderHierarchyAsync(string driveId, string relativePath, bool isFilePath)
        {
            var parts = relativePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
            {
                _logger.LogWarning("No valid folder parts found for path: {Path}", relativePath);
                return;
            }

            int lastIndex = isFilePath ? parts.Length - 1 : parts.Length;

            string currentPath = string.Empty;

            for (int i = 0; i < lastIndex; i++)
            {
                currentPath = string.IsNullOrEmpty(currentPath)
                    ? parts[i]
                    : $"{currentPath}/{parts[i]}";

                try
                {
                    await _graphClient.Drives[driveId].Root.ItemWithPath(currentPath).GetAsync();
                }
                catch (ODataError)
                {
                    // Need to create this folder
                    string parentPath = currentPath.Contains('/')
                        ? currentPath[..currentPath.LastIndexOf('/')]
                        : string.Empty;

                    string parentId;
                    if (string.IsNullOrEmpty(parentPath))
                    {
                        var root = await _graphClient.Drives[driveId].Root.GetAsync()
                            ?? throw new InvalidOperationException("Unable to read drive root item.");
                        parentId = root.Id!;
                    }
                    else
                    {
                        var parent = await _graphClient.Drives[driveId].Root.ItemWithPath(parentPath).GetAsync()
                            ?? throw new InvalidOperationException($"Unable to read parent folder: {parentPath}");
                        parentId = parent.Id!;
                    }

                    var folder = new DriveItem { Name = parts[i], Folder = new Folder() };
                    await _graphClient.Drives[driveId].Items[parentId].Children.PostAsync(folder);

                    _logger.LogInformation("Created folder: {Path}", currentPath);
                }
            }
        }

        #endregion
    }
}
