using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using ReportGenerator.Data.Models;
using ReportGenerator.Data.Repositories.Interfaces;

namespace ReportGenerator.Business.Services
{
    public class BridgeExcelProcessor(
        IReportRepository reportRepository,
        IExcelExtractionService excelExtractionService,
        IHttpClientFactory httpClientFactory,
        IAlternativeOptionsRepository alternativeOptionsRepository,
        IBridgeCharacteristicsRepository bridgeCharacteristicsRepository,
        ILogger<BridgeExcelProcessor> logger) : IBridgeExcelProcessor
    {
        private readonly IReportRepository _reportRepository = reportRepository;
        private readonly IExcelExtractionService _excelExtractionService = excelExtractionService;
        private readonly IHttpClientFactory _httpClientFactory = httpClientFactory;
        private readonly IAlternativeOptionsRepository _alternativeOptionsRepository = alternativeOptionsRepository;
        private readonly IBridgeCharacteristicsRepository _bridgeCharacteristicsRepository = bridgeCharacteristicsRepository;
        private readonly ILogger<BridgeExcelProcessor> _logger = logger;

        public async Task<BridgeExcelResult> ProcessAsync(Guid reportId, string fileUrl)
        {
            _logger.LogInformation("Starting bridge Excel processing for ReportId={ReportId}", reportId);

            var report = await _reportRepository.GetBasicReportInfoAsync(reportId);
            if (report == null)
            {
                _logger.LogWarning("Report {ReportId} not found.", reportId);
                throw new InvalidOperationException($"Report {reportId} not found.");
            }

            try
            {
                _logger.LogInformation("Downloading Excel file from URL: {Url}", fileUrl);

                var client = _httpClientFactory.CreateClient();
                var fileBytes = await client.GetByteArrayAsync(fileUrl);
                using var stream = new MemoryStream(fileBytes);

                _logger.LogInformation("Extracting Excel rows and building AlternativeData objects.");

                var result = _excelExtractionService.ProcessBridgeExcel(stream);

                _logger.LogInformation("Excel extraction complete. Extracted {Count} alternatives.", result.Alternatives.Count);

                ValidateAlternatives(result);

                await SaveExistingBridgeCharacteristicsAsync(report.BridgeID, result.Existing);

                await SaveAlternativesAsync(report.ReportID, result);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing Excel for ReportId={ReportId}", reportId);
                throw;
            }
        }

        private void ValidateAlternatives(BridgeExcelResult result)
        {
            foreach (var alt in result.Alternatives)
            {
                if (string.IsNullOrWhiteSpace(alt.StructureType))
                    _logger.LogWarning(  "Alternative {Alt} has missing StructureType. User input required.", alt.AlternativeNumber);

                if (alt.SpanLength == null && alt.BoxSpan == null && alt.PipeDiameter == null)
                    _logger.LogWarning("Alternative {Alt} has missing dimensions. Manual input required.", alt.AlternativeNumber);
            }
        }

        private async Task SaveExistingBridgeCharacteristicsAsync(int bridgeId, ExistingBridgeData existing)
        {
            if (existing == null)
            {
                _logger.LogWarning("No ExistingBridgeData extracted from Excel.");
                return;
            }

            _logger.LogInformation("Mapping ExistingBridgeData → BridgeCharacteristics");

            var mapped = MapExistingToCharacteristics(bridgeId, existing);

            await _bridgeCharacteristicsRepository.UpsertBridgeCharacteristicsAsync(bridgeId, mapped);

            _logger.LogInformation("BridgeCharacteristics updated successfully for BridgeID={BridgeID}", bridgeId);
        }

        public async Task SaveAlternativesAsync(int reportId, BridgeExcelResult result)
        {
            await _alternativeOptionsRepository.UpsertAlternative1Async(MapToAlt1(reportId, result.Alternatives[0]));

            await _alternativeOptionsRepository.UpsertAlternative2Async(MapToAlt2(reportId, result.Alternatives[1]));

            await _alternativeOptionsRepository.UpsertAlternative3Async(MapToAlt3(reportId, result.Alternatives[2]));

            await _alternativeOptionsRepository.UpsertAlternative4Async(MapToAlt4(reportId, result.Alternatives[3]));
        }

        private BridgeCharacteristics MapExistingToCharacteristics(int bridgeId, ExistingBridgeData src)
        {
            return new BridgeCharacteristics
            {
                BridgeID = bridgeId,
                ExistingStructureType = src.StructureType,
                StructureDescription = src.StructureDescription,
                LengthFeet = src.LengthFeet,
                SpanLength = src.SpanLength,
                BoxSpan = src.BoxSpan,
                BoxRise = src.BoxRise,
                PipeDiameter = src.PipeDiameter,
                NumberOfSpansOrCulverts = src.NumberOfSpansOrCulverts,
                LowChordElevation = src.LowChordElevation,
                ChannelInvertElevation = src.ChannelInvertElevation,
                WaterSurfaceElevation25Yr = src.WaterSurfaceElevation25Yr,
                HeadwaterToDiameterRatio25Yr = src.HeadwaterToDiameterRatio25Yr,
                WaterSurfaceElevation100Yr = src.WaterSurfaceElevation100Yr,
                HeadwaterToDiameterRatio100Yr = src.HeadwaterToDiameterRatio100Yr,
                WaterSurfaceElevation200Yr = src.WaterSurfaceElevation200Yr,
                HeadwaterToDiameterRatio200Yr = src.HeadwaterToDiameterRatio200Yr
            };
        }

        private Alternative1Option MapToAlt1(int reportId, AlternativeData src) => CreateAlternative<Alternative1Option>(reportId, src);
        
        private Alternative2Option MapToAlt2(int reportId, AlternativeData src) => CreateAlternative<Alternative2Option>(reportId, src);
        
        private Alternative3Option MapToAlt3(int reportId, AlternativeData src) => CreateAlternative<Alternative3Option>(reportId, src);

        private Alternative4Option MapToAlt4(int reportId, AlternativeData src) => CreateAlternative<Alternative4Option>(reportId, src);

        private T CreateAlternative<T>(int reportId, AlternativeData a) where T : AlternativeOptionBase, new()
        {
            return new T
            {
                ReportID = reportId,
                StructureType = a.StructureType!,
                StructureDescription = a.StructureDescription,
                StructureLength = a.LengthFeet,
                StructureLengthUnit = "feet",

                BoxSpan = a.BoxSpan,
                BoxSpanUnit = "feet",

                BoxRise = a.BoxRise,
                BoxRiseUnit = "feet",

                PipeDiameter = a.PipeDiameter,
                PipeDiameterUnit = "inches",

                CulvertOrSpanSize = a.SpanLength,
                CulvertOrSpanSizeUnit = "feet",

                NumberOfSpansOrCulverts = a.NumberOfSpansOrCulverts,

                LowChordElevation = a.LowChordElevation,

                WaterSurfaceElevation25Yr = a.WaterSurfaceElevation25Yr,
                WaterSurfaceElevation100Yr = a.WaterSurfaceElevation100Yr,
                WaterSurfaceElevation200Yr = a.WaterSurfaceElevation200Yr,

                HeadwaterToDiameterRatio25Yr = a.HeadwaterToDiameterRatio25Yr,
                HeadwaterToDiameterRatio100Yr = a.HeadwaterToDiameterRatio100Yr,
                HeadwaterToDiameterRatio200Yr = a.HeadwaterToDiameterRatio200Yr
            };
        }
    }
}
