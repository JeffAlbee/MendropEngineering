namespace ReportGenerator.Data.Models
{
    public class BridgeCharacteristics
    {
        public int CharacteristicsID { get; set; }
        public int BridgeID { get; set; }

        public string? BasinAreaSquareMiles { get; set; }

        // Watershed and Structure
        public decimal? StructureHeight { get; set; }
        public string? StructureHeightUnit { get; set; }
        public decimal? ElevationLowGround { get; set; }
        public string? ElevationLowGroundUnit { get; set; }

        public decimal? ExistingLengthFeet { get; set; }
        public decimal? ExistingChordHeight { get; set; }
        public string? ExistingLengthFeetUnit { get; set; }
        public string? ExistingChordHeightUnit { get; set; }

        public decimal? BaseRailElevation { get; set; }
        public string? BaseRailElevationUnit { get; set; }
        public string? NumberOfSpans { get; set; }

        public string? ExistingStructureType { get; set; }
        public string? StructureDescription { get; set; }
        public decimal? LengthFeet { get; set; }
        public decimal? SpanLength { get; set; }
        public decimal? BoxSpan { get; set; }
        public decimal? BoxRise { get; set; }
        public decimal? PipeDiameter { get; set; }
        public int? NumberOfSpansOrCulverts { get; set; }
        public decimal? LowChordElevation { get; set; }
        public decimal? ChannelInvertElevation { get; set; }
        public decimal? WaterSurfaceElevation25Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio25Yr { get; set; }
        public decimal? WaterSurfaceElevation100Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio100Yr { get; set; }
        public decimal? WaterSurfaceElevation200Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio200Yr { get; set; }

        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }
}
