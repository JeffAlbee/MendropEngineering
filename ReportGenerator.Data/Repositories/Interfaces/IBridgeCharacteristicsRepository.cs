using ReportGenerator.Data.Models;

namespace ReportGenerator.Data.Repositories.Interfaces
{
    public interface IBridgeCharacteristicsRepository
    {
        Task UpsertBridgeCharacteristicsAsync(int bridgeId, BridgeCharacteristics data);
    }
}
