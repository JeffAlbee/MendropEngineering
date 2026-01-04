using ReportGenerator.Data.Models;

namespace ReportGenerator.Data.Repositories.Interfaces
{
    public interface IReportRepository
    {
        Task<Report?> GetByIdAsync(Guid id);

        Task<ReportBasicInfo?> GetBasicReportInfoAsync(Guid reportId);

        Task InsertReportImageAsync(ReportImage image);
    }
}
