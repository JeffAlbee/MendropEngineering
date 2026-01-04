namespace ReportGenerator.Business.Models
{
    public class ImageUploadRequest
    {
        public Guid ReportId { get; set; }
        
        public string FileUrl { get; set; } = default!;

        public string FileName { get; set; } = default!;

        public string Category { get; set; } = default!;

        public string? UserName { get; set; }
    }
}
