using ReportGenerator.Business.Models;

namespace ReportGenerator.Business.Services.Interfaces
{
    public interface IReportService
    {
        Task<string> GenerateReportAsync(Guid id);

        Task EnsureFolderStructureByReportIdAsync(Guid reportId);

        Task<string> UploadImageToSharePointAsync(ImageUploadRequest request);
    }
}
