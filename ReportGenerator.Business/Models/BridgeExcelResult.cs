namespace ReportGenerator.Business.Models
{
    public class BridgeExcelResult
    {
        public ExistingBridgeData Existing { get; set; } = new ();

        public List<AlternativeData> Alternatives { get; set; } = new();
    }
}
