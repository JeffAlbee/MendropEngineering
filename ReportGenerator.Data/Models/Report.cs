namespace ReportGenerator.Data.Models
{
    public class Report
    {
        public Guid Id { get; set; }
        public int ReportID { get; set; }

        // General Info
        public string? ReportTitle { get; set; }
        public string ProjectNumber { get; set; } = default!;

        // Dates
        public DateTime? RfpDate { get; set; }
        public DateTime? DraftDate { get; set; }
        public DateTime? FinalDate { get; set; }

        // Prepared For
        public string? PreparedForName { get; set; }
        public string? PreparedForTitle { get; set; }
        public string? PreparedForOrganization { get; set; }
        public string? PreparedForAddressLine1 { get; set; }
        public string? PreparedForAddressLine2 { get; set; }
        public string? PreparedForCity { get; set; }
        public string? PreparedForState { get; set; }
        public string? PreparedForZip { get; set; }

        // Prepared By
        public string? PreparedByAddressLine1 { get; set; }
        public string? PreparedByAddressLine2 { get; set; }
        public string? PreparedByCity { get; set; }
        public string? PreparedByState { get; set; }
        public string? PreparedByZip { get; set; }
        public string? PreparedByPhone { get; set; }


        public string? CnContactName { get; set; }

        public decimal? UpstreamDistance { get; set; }
        public string? UpstreamDistanceUnit { get; set; }

        public decimal? DownstreamDistance { get; set; }
        public string? DownstreamDistanceUnit { get; set; }

        public decimal? TotalReachLength { get; set; }
        public string? TotalReachLengthUnit { get; set; }

        public string? PreferredStructureOption1 { get; set; }
        public string? PreferredStructureOption2 { get; set; }
        public string? HECRASVersion { get; set; }

        public DateTime? CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }

        public Bridge? Bridge { get; set; }
        public Alternative1Option? Alternative1Option { get; set; }
        public Alternative2Option? Alternative2Option { get; set; }
        public Alternative3Option? Alternative3Option { get; set; }
        public Alternative4Option? Alternative4Option { get; set; }
    }
}
