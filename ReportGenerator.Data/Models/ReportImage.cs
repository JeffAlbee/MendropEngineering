namespace ReportGenerator.Data.Models
{
    public class ReportImage
    {
        public Guid Id { get; set; }
        public int ReportId { get; set; }
        public string Category { get; set; } = default!;
        public string FileName { get; set; } = default!;
        public string SharePointUrl { get; set; } = default!;
        public DateTime UploadedAt { get; set; }
        public string? UploadedBy { get; set; }
    }
}
