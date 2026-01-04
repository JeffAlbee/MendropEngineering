using ReportGenerator.Business.Models;
namespace ReportGenerator.Business.Services.Interfaces
{
    public interface IExcelExtractionService
    {
        BridgeExcelResult ProcessBridgeExcel(Stream stream);
    }
}
