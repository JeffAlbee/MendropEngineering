namespace ReportGenerator.Business.Models
{
    public class ProcessExcelRequest
    {
        public Guid ReportId { get; set; }

        /// <summary>
        /// URL pointing to the Excel file (.xlsx). 
        /// This can be a SAS URL, SharePoint file URL, or uploaded file URL.
        /// </summary>
        public string FileUrl { get; set; } = default!;
    }
}
