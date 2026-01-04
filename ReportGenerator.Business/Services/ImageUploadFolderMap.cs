namespace ReportGenerator.Business.Services
{
    public static class ImageUploadFolderMap
    {
        public static readonly IReadOnlyDictionary<string, string> CategoryToFolder =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "upstream", "Photos/Raw" },
                { "downstream", "Photos/Raw" },
                { "location", "Photos/Raw" },
                { "detailed-location", "Photos/Raw" },
                { "topographic", "Photos/Raw" },
                { "sample-locations", "Photos/Raw" },

                { "fema", "Docs/Engineering Report/H&H/Appendix D - FEMA Documents" },
                { "appendix-a", "Docs/Engineering Report/H&H/Appendix A - Soil Report" },
                { "appendix-b", "Docs/Engineering Report/H&H/Appendix B - Web Soil Survey" },
                { "appendix-c", "Docs/Engineering Report/H&H/Appendix C - Hydrology" },
                { "appendix-d", "Docs/Engineering Report/H&H/Appendix D - FEMA Documents" },
                { "appendix-e", "Docs/Engineering Report/H&H/Appendix E - Model Output" }
            };

        public static IReadOnlyList<string> DistinctFolders =>
            CategoryToFolder.Values.Distinct().ToList();
    }
}
