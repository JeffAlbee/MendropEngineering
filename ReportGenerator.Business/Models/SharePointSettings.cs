namespace ReportGenerator.Business.Models
{
    public class SharePointSettings
    {
        public string SiteUrl { get; set; } = default!;
        public string CNReportsBasePath { get; set; } = default!;
        public string CNMasterTemplatePath { get; set; } = default!;
        public string DraftsFolderName { get; set; } = default!;
    }
}
