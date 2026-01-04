using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using ReportGenerator.Business.Extensions;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;

namespace ReportGenerator.Business.Services
{
    public class ExcelExtractionService(ILogger<ExcelExtractionService> logger) : IExcelExtractionService
    {
        private readonly ILogger<ExcelExtractionService> _logger = logger;

        public BridgeExcelResult ProcessBridgeExcel(Stream stream)
        {
            var result = new BridgeExcelResult();

            using var package = new ExcelPackage(stream);
            var workbook = package.Workbook;

            var lastSheet = workbook.Worksheets.Last();

            ExtractExistingBridgeData(lastSheet, result);

            ExtractAlternativeBridgeData(lastSheet, result);

            return result;
        }

        private void ExtractExistingBridgeData(ExcelWorksheet ws, BridgeExcelResult result)
        {
            const int col = 2;

            var description = ws.Cells[4, col].Text?.Trim();

            var existing = new ExistingBridgeData
            {
                StructureDescription = description,
                StructureType = DetectStructureType(description),

                LengthFeet = GetValue(ws, 5, col),
                NumberOfSpansOrCulverts = (int?)GetValue(ws, 7, col),

                LowChordElevation = GetValue(ws, 10, col),
                ChannelInvertElevation = GetValue(ws, 11, col),

                WaterSurfaceElevation25Yr = GetValue(ws, 14, col),
                HeadwaterToDiameterRatio25Yr = GetValue(ws, 17, col),

                WaterSurfaceElevation100Yr = GetValue(ws, 23, col),
                HeadwaterToDiameterRatio100Yr = GetValue(ws, 26, col),

                WaterSurfaceElevation200Yr = GetValue(ws, 32, col),
                HeadwaterToDiameterRatio200Yr = GetValue(ws, 35, col),
            };

            // Parse Dimensions row 6 same as alternatives
            NormalizeDimensions(ws.Cells[6, col].Text?.Trim(), existing);

            result.Existing = existing;

            _logger.LogInformation("Extracted Existing Bridge Data: {@Existing}", existing);
        }

        private void ExtractAlternativeBridgeData(ExcelWorksheet ws, BridgeExcelResult result)
        {
            _logger.LogInformation("Starting alternative bridge data extraction from worksheet: {Sheet}", ws.Name);

            for (int alt = 1; alt <= 4; alt++)
            {
                int col = alt + 2;

                _logger.LogInformation("Processing Alternative {Alt} (Column {Col})", alt, col);

                var description = ws.Cells[4, col].Text?.Trim();

                if (string.IsNullOrWhiteSpace(description))
                    _logger.LogWarning("Alternative {Alt}: StructureDescription (Row 4) is EMPTY.", alt);

                var model = new AlternativeData
                {
                    AlternativeNumber = alt,
                    StructureDescription = description,
                    StructureType = DetectStructureType(description),
                    LengthFeet = GetValue(ws, 5, col),
                    NumberOfSpansOrCulverts = (int?)GetValue(ws, 7, col),
                    LowChordElevation = GetValue(ws, 10, col),
                    ChannelInvertElevation = GetValue(ws, 11, col),
                    // Hydraulics: 25-year
                    WaterSurfaceElevation25Yr = GetValue(ws, 14, col),
                    HeadwaterToDiameterRatio25Yr = GetValue(ws, 17, col),
                    // Hydraulics: 100-year
                    WaterSurfaceElevation100Yr = GetValue(ws, 23, col),
                    HeadwaterToDiameterRatio100Yr = GetValue(ws, 26, col),
                    // Hydraulics: 200-year
                    WaterSurfaceElevation200Yr = GetValue(ws, 32, col),
                    HeadwaterToDiameterRatio200Yr = GetValue(ws, 35, col),
                };

                if (model.StructureType == "Unknown")
                    _logger.LogWarning("Alternative {Alt}: Unable to detect structure type from description: '{Desc}'", alt, description);

                NormalizeDimensions(ws.Cells[6, col].Text?.Trim(), model);

                result.Alternatives.Add(model);

                _logger.LogInformation("Alternative {Alt} extraction complete: {@Model}", alt, model);
            }

            _logger.LogInformation("HYDRAULICS extraction complete. Total alternatives: {Count}", result.Alternatives.Count);
        }

        private void NormalizeDimensions(string? rawDim, BridgeStructureBase model)
        {
            string id = model.GetLogId();

            if (string.IsNullOrWhiteSpace(rawDim))
            {
                _logger.LogWarning("Alternative {Alt}: Dimension cell (Row 6) is EMPTY. User input required.", id);
                return;
            }

            var raw = rawDim.ToLower().Replace("ft", "").Replace("\"", "").Replace("'", "").Trim();

            switch (model.StructureType)
            {
                // BRIDGE
                case "Bridge":
                    // Try to get span from row 6
                    if (decimal.TryParse(raw, out var culvertOrSpanSize))
                    {
                        model.SpanLength = culvertOrSpanSize;
                        _logger.LogInformation("Alternative {Alt}: Bridge span extracted: {Span}", id, culvertOrSpanSize);
                    }
                    else
                    {
                        _logger.LogWarning("Alternative {Alt}: Unable to parse bridge span from '{Raw}'.", id, raw);
                    }

                    break;

                // BOX CULVERT
                case "Box Culvert":
                    if (raw.Contains("x"))
                    {
                        var parts = raw.Split("x", StringSplitOptions.TrimEntries);

                        if (parts.Length == 2)
                        {
                            var spanRaw = new string(parts[0].Where(char.IsDigit).ToArray());
                            var riseRaw = new string(parts[1].Where(char.IsDigit).ToArray());

                            if (decimal.TryParse(spanRaw, out var span))
                                model.BoxSpan = span;
                            else
                                _logger.LogWarning("Alternative {Alt}: Could not parse BoxSpan from '{Val}'", id, parts[0]);

                            if (decimal.TryParse(riseRaw, out var rise))
                                model.BoxRise = rise;
                            else
                                _logger.LogWarning("Alternative {Alt}: Could not parse BoxRise from '{Val}'", id, parts[1]);
                        }
                        else
                        {
                            _logger.LogWarning("Alternative {Alt}: Unexpected box culvert format: '{Raw}'", id, raw);
                        }
                    }
                    else
                        _logger.LogWarning("Alternative {Alt}: Missing 'x' separator for box culvert: '{Raw}'", id, raw);
                    break;

                // PIPE CULVERT
                case "Pipe Culvert":
                    var digits = new string(raw.Where(char.IsDigit).ToArray());
                    if (decimal.TryParse(digits, out var diameter))
                    {
                        model.PipeDiameter = diameter;
                    }
                    else
                    {
                        _logger.LogWarning("Alternative {Alt}: Unable to parse pipe diameter from '{Raw}'. User input required.",
                            id, raw);
                    }
                    break;

                default:
                    _logger.LogWarning("Alternative {Alt}: Unknown structure type, cannot reliably parse dimensions. Raw='{Raw}'",
                        id, raw);
                    break;
            }
        }

        private decimal? GetValue(ExcelWorksheet ws, int row, int col)
        {
            var raw = ws.Cells[row, col].Text?
                .Replace("ft", "", StringComparison.OrdinalIgnoreCase)
                .Replace("\"", "")
                .Replace("'", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(raw, out var value))
                return value;

            return null;
        }

        private string DetectStructureType(string? description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return "Unknown";

            var text = description.ToLower();

            // --- Bridge variations ---
            if (text.Contains("bridge") ||
                text.Contains("trestle") ||
                text.Contains("girder") ||
                text.Contains("span") ||
                text.Contains("slab") ||
                text.Contains("beam"))
            {
                return "Bridge";
            }

            // --- Culverts ---
            if (text.Contains("box"))
                return "Box Culvert";

            if (text.Contains("pipe"))
                return "Pipe Culvert";

            // If the description contains dimensions like "9' x 8'"
            if (text.Contains("x"))
                return "Box Culvert";

            return "Unknown";
        }
    }
}
