using ReportGenerator.Business.Models;

namespace ReportGenerator.Business.Services.Interfaces
{
    public interface IBridgeExcelProcessor
    {
        Task<BridgeExcelResult> ProcessAsync(Guid reportId, string fileUrl);
    }
}
