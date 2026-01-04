namespace ReportGenerator.Business.Models
{
    public abstract class BridgeStructureBase
    {
        public string? StructureDescription { get; set; }
        public string? StructureType { get; set; }
        public decimal? LengthFeet { get; set; }
        /// <summary>
        /// For bridges: this stores the span length (ft).
        /// </summary>
        public decimal? SpanLength { get; set; }
        /// <summary>
        /// Box culvert width (span), extracted from Excel when dimensions are provided 
        /// in the form 'W x H' (e.g., '9 ft x 8 ft'). 
        /// Only applicable when StructureType represents a box culvert.
        /// </summary>
        public decimal? BoxSpan { get; set; }
        /// <summary>
        /// Box culvert height (rise), extracted from Excel when dimensions are provided 
        /// in the form 'W x H' (e.g., '9 ft x 8 ft').
        /// Only applicable when StructureType represents a box culvert.
        /// </summary>
        public decimal? BoxRise { get; set; }
        /// <summary>
        /// Pipe culvert diameter in inches or feet, depending on source sheet formatting.
        /// Extracted when Excel cell contains expressions like '84" dia' or similar.
        /// Only applicable when StructureType represents a pipe culvert.
        /// </summary>
        public decimal? PipeDiameter { get; set; }
        /// <summary>
        /// Number of structural units:
        /// - For bridges → number of spans  
        /// - For culverts → number of barrels (culvert cells)
        ///
        /// Extracted from Excel row labeled 
        /// ‘Number of Barrels or Spans’.
        //// </summary>
        public int? NumberOfSpansOrCulverts { get; set; }
        public decimal? LowChordElevation { get; set; }
        public decimal? ChannelInvertElevation { get; set; }

        public decimal? WaterSurfaceElevation25Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio25Yr { get; set; }

        public decimal? WaterSurfaceElevation100Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio100Yr { get; set; }

        public decimal? WaterSurfaceElevation200Yr { get; set; }
        public decimal? HeadwaterToDiameterRatio200Yr { get; set; }
    }
}
