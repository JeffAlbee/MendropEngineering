namespace ReportGenerator.Data.Models
{
    public abstract class AlternativeOptionBase
    {
        public int AlternativeOptionID { get; set; }
        public int ReportID { get; set; }

        public string StructureType { get; set; } = default!;

        public decimal? StructureLength { get; set; }
        public string? StructureLengthUnit { get; set; }

        public decimal? PipeOrBoxLength { get; set; }
        public string? PipeOrBoxLengthUnit { get; set; }

        public decimal? BoxSpan { get; set; }
        public string? BoxSpanUnit { get; set; }

        public decimal? BoxRise { get; set; }
        public string? BoxRiseUnit { get; set; }

        public decimal? PipeDiameter { get; set; }
        public string? PipeDiameterUnit { get; set; }

        //Span Length applies to bridges
        public decimal? CulvertOrSpanSize { get; set; }
        public string? CulvertOrSpanSizeUnit { get; set; }

        public int? NumberOfSpansOrCulverts { get; set; }

        //newly added fields
        public string? StructureDescription { get; set; }
        public decimal? TopOfCulvertElevation { get; set; } // only for culverts
        public decimal? LowChordElevation { get; set; }
        public decimal? WaterSurfaceElevation25Yr { get; set; }
        public decimal? WaterSurfaceElevation100Yr { get; set; }
        public decimal? WaterSurfaceElevation200Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio25Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio100Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio200Yr { get; set; }

    }
}
