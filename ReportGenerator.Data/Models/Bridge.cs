namespace ReportGenerator.Data.Models
{
    public class Bridge
    {
        public int BridgeID { get; set; }
        public int ReportID { get; set; }

        public string BridgeCode { get; set; } = default!;
        public decimal? ProjectMilepost { get; set; }
        public string? Location { get; set; }
        public string? LocationCounty { get; set; }
        public string? LocationState { get; set; }
        public string? LocationCity { get; set; }
        public string? LocationRoadInterchange { get; set; }
        public string? RailroadDivision { get; set; }
        public string? RailroadSubdivision { get; set; }
        public string? FirmPanelNumber { get; set; }
        public DateTime? FemaEffectiveDate { get; set; }

        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }

        public BridgeCharacteristics? Characteristics { get; set; }
    }
}
