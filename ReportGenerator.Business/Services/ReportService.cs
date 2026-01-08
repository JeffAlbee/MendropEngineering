using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ReportGenerator.Business.Helpers;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using ReportGenerator.Data.Models;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;

namespace ReportGenerator.Business.Services
{
    public class ReportService : IReportService
    {
        #region private fields

        private readonly IReportRepository _reportRepository;
        private readonly ISharePointService _sharePointService;
        private readonly SharePointSettings _sharePointSettings;
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly ILogger<ReportService> _logger;

        #endregion

        #region Constructor

        public ReportService(
            IReportRepository reportRepository, 
            ISharePointService sharePointService,
            IOptions<SharePointSettings> sharePointOptions,
            IHttpClientFactory httpClientFactory,
            ILogger<ReportService> logger)
        {
            _reportRepository = reportRepository;
            _sharePointService = sharePointService;
            _sharePointSettings = sharePointOptions.Value ?? throw new ArgumentNullException(nameof(sharePointOptions));
            _httpClientFactory = httpClientFactory;
            _logger = logger;
        }

        #endregion

        #region Public Methods

        public async Task<string> GenerateReportAsync(Guid id)
        {
            try
            {
                _logger.LogInformation("Generating report for Report Id: {ReportId}", id);

                var templatePath = _sharePointSettings.CNMasterTemplatePath;

                if (string.IsNullOrWhiteSpace(templatePath))
                    throw new InvalidOperationException("TemplatePath not configured in settings.");

                var templateBytes = await _sharePointService.DownloadDocumentAsync(templatePath);

                if (templateBytes == null || templateBytes.Length == 0)
                    throw new FileNotFoundException($"Template not found in SharePoint: {templatePath}");

                var report = await FetchReportAsync(id);

                var placeholders = WordTemplateHelper.GetPlaceholders(templateBytes);

                _logger.LogInformation("Found {Count} merge fields.", placeholders.Count);

                // Pull images from SharePoint
                var imagePaths = await _sharePointService.GetAllImagePathsForReportAsync(report.ProjectNumber);

                var downloadedImages = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                var tasks = imagePaths.Select(async imagePath =>
                {
                    var fileName = Path.GetFileName(imagePath);

                    var cleanKey = Regex.Replace(
                        Path.GetFileNameWithoutExtension(fileName).ToLower(),
                        "[^a-z0-9_]", "_"
                    );

                    var placeholderKey = $"img_{cleanKey}";

                    var localPath = Path.Combine(Path.GetTempPath(), $"{fileName}");

                    var bytes = await _sharePointService.DownloadDocumentAsync(imagePath);

                    await File.WriteAllBytesAsync(localPath, bytes);

                    downloadedImages[placeholderKey] = localPath;
                });

                await Task.WhenAll(tasks);

                var values = BuildPlaceholderValues(report);

                foreach (var kvp in downloadedImages)
                {
                    values[kvp.Key] = kvp.Value;
                }

                var outputBytes = WordTemplateHelper.ReplacePlaceholders(templateBytes, values);

                _logger.LogInformation("Report generated successfully for Report Id: {ReportId}", id);

                // Build SharePoint target path
                string relativePath = BuildDraftReportPath(report.ProjectNumber);

                _logger.LogInformation("Uploaded generated report to SharePoint: {Path}", relativePath);

                // Upload to SharePoint
                var driveItem = await _sharePointService.UploadContentAndReturnItemAsync(relativePath, outputBytes);

                return driveItem?.WebUrl ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error generating report for Report Id: {id} - {ex.Message}");
                throw;
            }  
        }

        public async Task EnsureFolderStructureByReportIdAsync(Guid reportId)
        {
            var reportInfo = await _reportRepository.GetBasicReportInfoAsync(reportId);

            if (reportInfo == null)
            {
                _logger.LogWarning("No basic report info found for Report Id {ReportId}", reportId);
                throw new InvalidOperationException($"Report with ID {reportId} not found.");
            }

            string relativePath = $"GeneratedReports/{reportInfo.ProjectNumber}/{reportInfo.BridgeCode}";

            await _sharePointService.EnsureFolderHierarchyAsync(relativePath);

            _logger.LogInformation("Ensured folder structure for Report {Id} -> {Path}", reportInfo.ReportID, relativePath);
        }

        public async Task<string> UploadImageToSharePointAsync(ImageUploadRequest request)
        {
            try
            {
                var client = _httpClientFactory.CreateClient();
                var imageBytes = await client.GetByteArrayAsync(request.FileUrl);

                var report = await _reportRepository.GetBasicReportInfoAsync(request.ReportId)
                   ?? throw new InvalidOperationException($"Report {request.ReportId} not found.");

                if (!ImageUploadFolderMap.CategoryToFolder.TryGetValue(request.Category, out var folderPath))
                    throw new InvalidOperationException($"Unsupported image category '{request.Category}'.");

                var relativePath =$"{_sharePointSettings.CNReportsBasePath}/{report.ProjectNumber}/{folderPath}/{request.FileName}";

                var uploaded = await _sharePointService.UploadContentAndReturnItemAsync(relativePath, imageBytes);

                _logger.LogInformation("Uploaded image to SharePoint: {Url}", uploaded.WebUrl);

                var imageRecord = new ReportImage
                {
                    ReportId = report.ReportID,
                    Category = request.Category,
                    FileName = request.FileName,
                    SharePointUrl = uploaded.WebUrl!,
                    UploadedBy = request.UserName
                };

                await _reportRepository.InsertReportImageAsync(imageRecord);

                _logger.LogInformation("Saved image metadata to SQL for ReportID {ReportId}", report.ReportID);

                return uploaded.WebUrl!;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading image from Retool Storage");
                throw;
            }
        }

        #endregion

        #region Private Methods

        private async Task<Report> FetchReportAsync(Guid id)
        {
            var report = await _reportRepository.GetByIdAsync(id);

            if (report == null)
            {
                _logger.LogWarning("No report found for Report Id {ReportId}", id);
                throw new InvalidOperationException($"Report with ID {id} not found.");
            }

            _logger.LogInformation("Report data retrieved for Report Id: {ReportId}", id);
            return report;
        }

        private static Dictionary<string, object?> BuildCoverPagePlaceholderValues(Report report)
        {
            var values = new Dictionary<string, object?>
            {
                ["project_number"] = report?.ProjectNumber,
                ["bridge_id"] = report?.Bridge?.BridgeCode,
                ["REPORT_TITLE"] = report?.ReportTitle,
                ["location"] = report?.Bridge?.Location,
                ["location_county"] = report?.Bridge?.LocationCounty,
                ["location_state"] = report?.Bridge?.LocationState,
                ["railroad_division"] = report?.Bridge?.RailroadDivision,
                ["railroad_subdivision"] = report?.Bridge?.RailroadSubdivision,
                ["draft_date"] = report?.DraftDate?.ToString("MMMM dd, yyyy"),
                ["final_date"] = report?.FinalDate?.ToString("MMMM dd, yyyy"),
                ["prepared_for_name"] = report?.PreparedForName,
                ["prepared_for_title"] = report?.PreparedForTitle,
                ["prepared_for_organization"] = report?.PreparedForOrganization,
                ["prepared_for_address_line1"] = report?.PreparedForAddressLine1,
                ["prepared_for_address_line2"] = report?.PreparedForAddressLine2,
                ["prepared_for_city"] = report?.PreparedForCity,
                ["prepared_for_state"] = report?.PreparedForState,
                ["prepared_for_zip"] = report?.PreparedForZip,
                ["prepared_by_address_line1"] = report?.PreparedByAddressLine1,
                ["prepared_by_address_line2"] = report?.PreparedByAddressLine2,
                ["prepared_by_city"] = report?.PreparedByCity,
                ["prepared_by_state"] = report?.PreparedByState,
                ["prepared_by_zip"] = report?.PreparedByZip,
                ["prepared_by_phone"] = report?.PreparedByPhone,
            };

            return values;
        }

        private static Dictionary<string, object?> BuildIntroductionPagePlaceholderValues(Report report)
        {
            var values = new Dictionary<string, object?>
            {
                ["rfp_date"] = report.RfpDate?.ToString("MMMM dd, yyyy"),
                ["cn_contact_name"] = report?.CnContactName,
                ["location_road_interchange"] = report?.Bridge?.LocationRoadInterchange,
                ["location_city"] = report?.Bridge?.LocationCity,
                ["location_county"] = report?.Bridge?.LocationCounty,
                ["firm_panel_number"] = report?.Bridge?.FirmPanelNumber,
                ["firm_effective_date"] = report?.Bridge?.FemaEffectiveDate?.ToString("MMMM dd, yyyy"),
                ["project_milepost"] = report?.Bridge?.ProjectMilepost
            };

            return values;
        }

        private static Dictionary<string, object?> BuildStudyAreaPagePlaceholderValues(Report report)
        {
            var values = new Dictionary<string, object?>
            {
                ["existing_structure_type"] = report?.Bridge?.Characteristics?.ExistingStructureType,
                ["existing_structure_description"] = report?.Bridge?.Characteristics?.StructureDescription,
                ["existing_length_feet"] = report?.Bridge?.Characteristics?.ExistingLengthFeet,
                ["existing_length_feet_unit"] = report?.Bridge?.Characteristics?.ExistingLengthFeet != null ? "feet" : null,
                ["structure_height"] = report?.Bridge?.Characteristics?.StructureHeight,
                ["structure_height_unit"] = report?.Bridge?.Characteristics?.StructureHeight != null ? "feet" : null,
                ["base_rail_elevation"] = report?.Bridge?.Characteristics?.BaseRailElevation,
                ["base_rail_elevation_unit"] = report?.Bridge?.Characteristics?.BaseRailElevation != null ? "feet" : null,
                ["existing_chord_height"] = report?.Bridge?.Characteristics?.ExistingChordHeight,
                ["existing_chord_height_unit"] = report?.Bridge?.Characteristics?.ExistingChordHeight != null ? "feet" : null,
                ["elevation_low_ground"] = report?.Bridge?.Characteristics?.ElevationLowGround,
                ["elevation_low_ground_unit"] = report?.Bridge?.Characteristics?.ElevationLowGround != null ? "feet" : null,
                ["basin_area_square_miles"] = report?.Bridge?.Characteristics?.BasinAreaSquareMiles,
                ["number_of_spans"] = report?.Bridge?.Characteristics?.NumberOfSpans
            };

            return values;
        }

        private static Dictionary<string, object?> BuildEvaluationMethodsPagePlaceholderValues(Report report)
        {
            var values = new Dictionary<string, object?>
            {
                ["hec_res_version"] = report?.HECRASVersion,
                ["upstream_distance"] = report?.UpstreamDistance,
                ["upstream_distance_unit"] = report?.UpstreamDistanceUnit,
                ["downstream_distance"] = report?.DownstreamDistance,
                ["downstream_distance_unit"] = report?.DownstreamDistanceUnit,
                ["total_reach_length"] = report?.TotalReachLength,
                ["total_reach_length_unit"] = report?.TotalReachLengthUnit,
                ["preferred_structure_option_1"] = report?.PreferredStructureOption1,
                ["preferred_structure_option_2"] = report?.PreferredStructureOption2,

                ["alt1_structure_type"] = report?.Alternative1Option?.StructureType,
                ["alt1_structure_desc"] = BuildStructureDescription(report?.Alternative1Option),

                ["alt2_structure_type"] = report?.Alternative2Option?.StructureType,
                ["alt2_structure_desc"] = BuildStructureDescription(report?.Alternative2Option),

                ["alt3_structure_type"] = report?.Alternative3Option?.StructureType,
                ["alt3_structure_desc"] = BuildStructureDescription(report?.Alternative3Option),

                ["alt4_structure_type"] = report?.Alternative4Option?.StructureType,
                ["alt4_structure_desc"] = BuildStructureDescription(report?.Alternative4Option),
            };

            return values;
        }

        private static Dictionary<string, object?> BuildConclusionSectionPlaceholderValues(Report report)
        {
            var values = new Dictionary<string, object?>
            {
                ["alt1_summary_structure_desc"] = BuildShortStructureSummary(report.Alternative1Option),
                ["alt2_summary_structure_desc"] = BuildShortStructureSummary(report.Alternative2Option),
                ["alt3_summary_structure_desc"] = BuildShortStructureSummary(report.Alternative3Option),
                ["alt4_summary_structure_desc"] = BuildShortStructureSummary(report.Alternative4Option),

                ["alt1_guideline_bullets"] = BuildGuidelineBullets(report.Alternative1Option, report.Bridge?.Characteristics),
                ["alt2_guideline_bullets"] = BuildGuidelineBullets(report.Alternative2Option, report.Bridge?.Characteristics),
                ["alt3_guideline_bullets"] = BuildGuidelineBullets(report.Alternative3Option, report.Bridge?.Characteristics),
                ["alt4_guideline_bullets"] = BuildGuidelineBullets(report.Alternative4Option, report.Bridge?.Characteristics),
            };

            return values;
        }

        private static Dictionary<string, object?> BuildPlaceholderValues(Report report)
        {
            var result = new Dictionary<string, object?>();

            void Merge(Dictionary<string, object?> source)
            {
                foreach (var kv in source)
                    result[kv.Key] = kv.Value; // overwrite if key already exists
            }

            Merge(BuildCoverPagePlaceholderValues(report));
            Merge(BuildIntroductionPagePlaceholderValues(report));
            Merge(BuildStudyAreaPagePlaceholderValues(report));
            Merge(BuildEvaluationMethodsPagePlaceholderValues(report));
            Merge(BuildConclusionSectionPlaceholderValues(report));

            return result;
        }

        private string BuildDraftReportPath(string projectNumber)
        {
            var basePath = _sharePointSettings.CNReportsBasePath.TrimEnd('/');
            var draftsFolder = _sharePointSettings.DraftsFolderName.Trim('/');

            return $"{basePath}/{projectNumber}/{draftsFolder}/Draft_Report_{DateTime.Now:yyyyMMdd_HHmm}.docx";
        }

        private static string BuildStructureDescription(AlternativeOptionBase? alt)
        {
            if (alt == null || string.IsNullOrWhiteSpace(alt.StructureType))
                return string.Empty;

            string type = alt.StructureType.Trim().ToLowerInvariant();
            int count = alt.NumberOfSpansOrCulverts ?? 0;

            static string CountWord(int n) => n switch
            {
                1 => "one",
                2 => "two",
                3 => "three",
                4 => "four",
                _ => n.ToString()
            };

            switch (type)
            {
                // BRIDGE
                case "bridge":
                    if (count > 0 &&
                        alt.StructureLength is decimal spanLength &&
                        !string.IsNullOrWhiteSpace(alt.StructureLengthUnit))
                    {
                        return $"{alt.StructureDescription} {count} span{(count > 1 ? "s" : "")} at {spanLength:0.##} {alt.StructureLengthUnit} long";
                    }
                    return alt.StructureDescription ?? string.Empty;

                // BOX CULVERT
                case "box culvert":
                    if (count > 0 &&
                        alt.BoxSpan is decimal w &&
                        alt.BoxRise is decimal h &&
                        alt.StructureLength is decimal len)
                    {
                        return
                            $"{CountWord(count)} box culvert{(count > 1 ? "s" : "")}, " +
                            $"each being {w:0.##}'(w) x {h:0.##}'(h) and {len:0.##} feet in length";
                    }
                    return string.Empty;

                // PIPE CULVERT
                case "pipe culvert":
                    if (count > 0 &&
                        alt.PipeDiameter is decimal dia &&
                        alt.StructureLength is decimal plen)
                    {
                        return
                            $"{CountWord(count)} pipe culvert{(count > 1 ? "s" : "")} " +
                            $"at {dia:0.##}\" diameter and {plen:0.##} feet in length";
                    }
                    return string.Empty;

                default:
                    return string.Empty;
            }
        }

        private static string BuildShortStructureSummary(AlternativeOptionBase? alt)
        {
            if (alt == null || string.IsNullOrWhiteSpace(alt.StructureType))
                return string.Empty;

            string type = alt.StructureType.Trim().ToLower();

            switch (type)
            {
                case "bridge":
                    // Example: "Flat slab girder bridge with 3 spans at 14’ long."
                    if (alt.NumberOfSpansOrCulverts is int spans &&
                        alt.CulvertOrSpanSize is decimal spanLength)
                    {
                        return $"{alt.StructureDescription} with {spans} span{(spans > 1 ? "s" : "")} at {spanLength}’ long.";
                    }
                    break;


                case "box culvert":
                    // Example: "Double concrete box culverts, 9’ x 8’, 36 ft long."
                    if (alt.NumberOfSpansOrCulverts is int boxes &&
                        alt.BoxSpan is decimal w &&
                        alt.BoxRise is decimal h &&
                        alt.StructureLength is decimal length)
                    {
                        string prefix = boxes switch
                        {
                            2 => "Double",
                            3 => "Triple",
                            _ => $"{boxes}"
                        };

                        return $"{prefix} {alt.StructureDescription}, " + $"{w}'(w) x {h}'(h) and {length} feet in length";
                    }
                    break;


                case "pipe culvert":
                    if (alt.NumberOfSpansOrCulverts is int pipes &&
                        alt.PipeDiameter is decimal diameter &&
                        alt.StructureLength is decimal structureLength)
                    {
                        string word = pipes switch
                        {
                            1 => "One",
                            2 => "Two",
                            3 => "Three",
                            4 => "Four",
                            _ => pipes.ToString()
                        };

                        return $"{word} ({pipes}) {diameter}\" diameter {alt.StructureDescription}, " +
                                $"{structureLength} feet in length";
                    }
                    break;
            }

            return string.Empty;
        }

        private static List<string> BuildGuidelineBullets(
            AlternativeOptionBase? alt,
            BridgeCharacteristics? bc)
        {
            if (alt == null || string.IsNullOrWhiteSpace(alt.StructureType) || bc == null)
                return new List<string>();

            var bullets = new List<string>();

            string type = alt.StructureType.Trim().ToLowerInvariant();
            decimal? baseRail = bc?.BaseRailElevation; // preferred: from bridge characteristics

            // Guideline #1 (ALL TYPES)
            if (alt.WaterSurfaceElevation100Yr.HasValue)
            {
                var wse100 = alt.WaterSurfaceElevation100Yr.Value;
                string compare = DescribeDifference(wse100, bc.WaterSurfaceElevation100Yr);

                bullets.Add(
                    $"Guideline #1 is achieved. " +
                    $"The 100-year water surface elevation of {FmtFt(wse100)} is {compare} " +
                    $"the existing conditions water surface elevation of {FmtFt(bc.WaterSurfaceElevation100Yr)} " +
                    $"at normal depth conditions."
                );
            }

            // Bridge guidelines (#2a, #2b)
            if (type == "bridge")
            {
                // #2a: 200-yr below base of rail
                if (alt.WaterSurfaceElevation200Yr.HasValue && baseRail.HasValue)
                {
                    var wse200 = alt.WaterSurfaceElevation200Yr.Value;
                    var diff = baseRail.Value - wse200;

                    bullets.Add(
                        $"Guideline #2a is achieved. " +
                        $"The 200-year water surface elevation of {FmtFt(wse200)} is {FmtFt(Math.Abs(diff))} below " +
                        $"the base of rail elevation of {FmtFt(baseRail.Value)} for normal depth conditions."
                    );
                }

                // #2b: 100-yr below low chord (min)
                if (alt.WaterSurfaceElevation100Yr.HasValue && alt.LowChordElevation.HasValue)
                {
                    var wse100 = alt.WaterSurfaceElevation100Yr.Value;
                    var lowChord = alt.LowChordElevation.Value;
                    var diff = lowChord - wse100;

                    bullets.Add(
                        $"Guideline #2b is achieved. " +
                        $"The 100-year water surface elevation of {FmtFt(wse100)} is {FmtFt(Math.Abs(diff))} below " +
                        $"the minimum low chord elevation of {FmtFt(lowChord)} for normal depth conditions."
                    );
                }

                return bullets;
            }

            // Culvert guidelines (#3a–#3e)
            if (type is "pipe culvert" or "box culvert")
            {
                // #3a: 100-yr below base of rail
                if (alt.WaterSurfaceElevation100Yr.HasValue && baseRail.HasValue)
                {
                    var wse100 = alt.WaterSurfaceElevation100Yr.Value;
                    var diff = baseRail.Value - wse100;

                    bullets.Add(
                        $"Guideline #3a is achieved. " +
                        $"The 100-year water surface elevation of {FmtFt(wse100)} is {FmtFt(Math.Abs(diff))} below " +
                        $"the base of rail elevation of {FmtFt(baseRail.Value)} for normal depth conditions."
                    );
                }

                // #3b: HW/D 25-year
                if (alt.HeadwaterToDiameterRatio25Yr.HasValue)
                {
                    bullets.Add(
                        $"Guideline #3b is achieved. " +
                        $"The HW/D for the 25-year event is {alt.HeadwaterToDiameterRatio25Yr.Value:0.00}."
                    );
                }

                // #3c: HW/D 100-year
                if (alt.HeadwaterToDiameterRatio100Yr.HasValue)
                {
                    bullets.Add(
                        $"Guideline #3c is achieved. " +
                        $"The HW/D for the 100-year event is {alt.HeadwaterToDiameterRatio100Yr.Value:0.00}."
                    );
                }

                // #3d / #3e: cover depth from top of pipe/box to base of rail
                if (alt.TopOfCulvertElevation.HasValue && baseRail.HasValue)
                {
                    var cover = baseRail.Value - alt.TopOfCulvertElevation.Value;

                    if (type == "pipe culvert")
                    {
                        bullets.Add(
                            $"Guideline #3d is achieved. " +
                            $"The approximate cover depth from top of pipe to base of rail is {FmtFt(Math.Abs(cover))}."
                        );
                    }
                    else // box culvert
                    {
                        bullets.Add(
                            $"Guideline #3e is achieved. " +
                            $"The approximate cover depth from top of box to base of rail is {FmtFt(Math.Abs(cover))}."
                        );
                    }
                }

                return bullets;
            }

            // Unknown structure type → only guideline #1 (if present) is returned
            return bullets;
        }

        private static string DescribeDifference(decimal value, decimal? reference)
        {
            if (!reference.HasValue)
                return string.Empty;

            decimal diff = value - reference.Value;

            // treat tiny difference as equal (avoid weird "0.0' lower than")
            if (Math.Abs(diff) < 0.05m)
                return "equal to";

            return $"approximately {Math.Abs(diff):0.0}' {(diff < 0 ? "lower than" : "higher than")}";
        }

        private static string FmtFt(decimal? value)
        {
            // Always show one decimal and a trailing foot mark, like 291.5'
            return $"{value:0.0}'";
        }

        #endregion
    }
}
