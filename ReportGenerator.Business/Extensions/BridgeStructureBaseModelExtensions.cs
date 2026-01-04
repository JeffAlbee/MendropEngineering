using ReportGenerator.Business.Models;

namespace ReportGenerator.Business.Extensions
{
    public static class BridgeStructureBaseModelExtensions
    {
        public static string GetLogId(this BridgeStructureBase model)
        {
            if (model is AlternativeData alt)
                return $"Alternative {alt.AlternativeNumber}";

            return "Existing";
        }
    }
}
