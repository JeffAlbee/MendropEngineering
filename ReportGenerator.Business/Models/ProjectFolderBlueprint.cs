namespace ReportGenerator.Business.Models
{
    public static class ProjectFolderBlueprint
    {
        public static readonly FolderNode CN = new(string.Empty,
        [
            new FolderNode("CADD Files"),

            new FolderNode("Docs",
            [
                new FolderNode("Engineering Report",
                [
                    new FolderNode("H&H",
                    [
                        new FolderNode("Appendix A - Soil Report"),
                        new FolderNode("Appendix B - Web Soil Survey"),
                        new FolderNode("Appendix C - Hydrology"),
                        new FolderNode("Appendix D - FEMA Documents"),
                        new FolderNode("Appendix E - Model Output")
                    ]),
                    new FolderNode("Survey"),
                    new FolderNode("Drafts")
                ]),

                new FolderNode("Meeting Minutes"),
                new FolderNode("Project Closeout"),
                new FolderNode("Proposal Documents"),
                new FolderNode("QAQC"),
                new FolderNode("Work Assignment",
                [
                    new FolderNode("Project Setup")
                ])
            ]),

            new FolderNode("Hydraulics",
            [
                new FolderNode("Calculations"),
                new FolderNode("HEC-RAS"),
                new FolderNode("HY-8"),
                new FolderNode("SRH-2D")
            ]),

            new FolderNode("Hydrology",
            [
                new FolderNode("Calculations"),
                new FolderNode("HEC-HMS"),
                new FolderNode("HYDR2009"),
                new FolderNode("StreamStats")
            ]),

            new FolderNode("Photos",
            [
                new FolderNode("Raw"),
                new FolderNode("360 Photos",
                [
                    new FolderNode("Raw")
                ])
            ]),

            new FolderNode("Information Received from Outside Source", 
            [
                new FolderNode("811"),
                new FolderNode("CN"),
                new FolderNode("Waypoint Analytical")
            ]),

            new FolderNode("Information Sent to Outside Source"),

            new FolderNode("Submittals"),

            new FolderNode("Survey",
            [
                new FolderNode("Download"),
                new FolderNode("Field Notes"),
                new FolderNode("OPUS")
            ])
        ]);
    }
}
