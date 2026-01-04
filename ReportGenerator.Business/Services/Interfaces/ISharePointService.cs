using Microsoft.Graph.Models;

namespace ReportGenerator.Business.Services.Interfaces
{
    public interface ISharePointService
    {
        Task<byte[]> DownloadDocumentAsync(string relativePath);

        Task<DriveItem> UploadContentAndReturnItemAsync(string relativePath, byte[] fileContent);

        Task EnsureFolderHierarchyAsync(string relativePath, bool isFilePath = false);

        Task<List<string>> GetImagePathsAsync(string projectNumber, string bridgeCode);

        Task<List<string>> GetAllImagePathsForReportAsync(string projectNumber);
    }
}
