namespace ReportGenerator.Data.Models
{
    public class ReportBasicInfo
    {
        public int ReportID { get; set; }

        public string ProjectNumber { get; set; } = default!;

        public int BridgeID { get; set; }

        public string BridgeCode { get; set; } = default!;
    }
}
