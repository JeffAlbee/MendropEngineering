using Dapper;
using ReportGenerator.Data.Models;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Data;

namespace ReportGenerator.Data.Repositories
{
    public class BridgeCharacteristicsRepository : IBridgeCharacteristicsRepository
    {
        private readonly IDbConnection _connection;

        public BridgeCharacteristicsRepository(IDbConnection connection)
        {
            _connection = connection;
        }

        public async Task UpsertBridgeCharacteristicsAsync(int bridgeId, BridgeCharacteristics data)
        {
            const string sql = @"
            MERGE dbo.BridgeCharacteristics AS target
            USING (VALUES
            (
                @BridgeID,
                @ExistingStructureType,
                @StructureDescription,
                @LengthFeet,
                @SpanLength,
                @BoxSpan,
                @BoxRise,
                @PipeDiameter,
                @NumberOfSpansOrCulverts,
                @LowChordElevation,
                @ChannelInvertElevation,
                @WaterSurfaceElevation25Yr,
                @HeadwaterToDiameterRatio25Yr,
                @WaterSurfaceElevation100Yr,
                @HeadwaterToDiameterRatio100Yr,
                @WaterSurfaceElevation200Yr,
                @HeadwaterToDiameterRatio200Yr
            )) AS source
            (
                BridgeID,
                ExistingStructureType,
                StructureDescription,
                LengthFeet,
                SpanLength,
                BoxSpan,
                BoxRise,
                PipeDiameter,
                NumberOfSpansOrCulverts,
                LowChordElevation,
                ChannelInvertElevation,
                WaterSurfaceElevation25Yr,
                HeadwaterToDiameterRatio25Yr,
                WaterSurfaceElevation100Yr,
                HeadwaterToDiameterRatio100Yr,
                WaterSurfaceElevation200Yr,
                HeadwaterToDiameterRatio200Yr
            )
            ON target.BridgeID = source.BridgeID

            WHEN MATCHED THEN
                UPDATE SET
                    ExistingStructureType          = source.ExistingStructureType,
                    StructureDescription           = source.StructureDescription,
                    LengthFeet                     = source.LengthFeet,
                    SpanLength                     = source.SpanLength,
                    BoxSpan                        = source.BoxSpan,
                    BoxRise                        = source.BoxRise,
                    PipeDiameter                   = source.PipeDiameter,
                    NumberOfSpansOrCulverts        = source.NumberOfSpansOrCulverts,
                    LowChordElevation              = source.LowChordElevation,
                    ChannelInvertElevation         = source.ChannelInvertElevation,
                    WaterSurfaceElevation25Yr      = source.WaterSurfaceElevation25Yr,
                    HeadwaterToDiameterRatio25Yr   = source.HeadwaterToDiameterRatio25Yr,
                    WaterSurfaceElevation100Yr     = source.WaterSurfaceElevation100Yr,
                    HeadwaterToDiameterRatio100Yr  = source.HeadwaterToDiameterRatio100Yr,
                    WaterSurfaceElevation200Yr     = source.WaterSurfaceElevation200Yr,
                    HeadwaterToDiameterRatio200Yr  = source.HeadwaterToDiameterRatio200Yr,
                    UpdatedAt = GETUTCDATE()

            WHEN NOT MATCHED THEN
                INSERT
                (
                    BridgeID,
                    ExistingStructureType,
                    StructureDescription,
                    LengthFeet,
                    SpanLength,
                    BoxSpan,
                    BoxRise,
                    PipeDiameter,
                    NumberOfSpansOrCulverts,
                    LowChordElevation,
                    ChannelInvertElevation,
                    WaterSurfaceElevation25Yr,
                    HeadwaterToDiameterRatio25Yr,
                    WaterSurfaceElevation100Yr,
                    HeadwaterToDiameterRatio100Yr,
                    WaterSurfaceElevation200Yr,
                    HeadwaterToDiameterRatio200Yr,
                    CreatedAt,
                    UpdatedAt
                )
                VALUES
                (
                    source.BridgeID,
                    source.ExistingStructureType,
                    source.StructureDescription,
                    source.LengthFeet,
                    source.SpanLength,
                    source.BoxSpan,
                    source.BoxRise,
                    source.PipeDiameter,
                    source.NumberOfSpansOrCulverts,
                    source.LowChordElevation,
                    source.ChannelInvertElevation,
                    source.WaterSurfaceElevation25Yr,
                    source.HeadwaterToDiameterRatio25Yr,
                    source.WaterSurfaceElevation100Yr,
                    source.HeadwaterToDiameterRatio100Yr,
                    source.WaterSurfaceElevation200Yr,
                    source.HeadwaterToDiameterRatio200Yr,
                    GETUTCDATE(),
                    GETUTCDATE()
                );";

            var parameters = new
            {
                BridgeID = bridgeId,
                data.ExistingStructureType,
                data.StructureDescription,
                data.LengthFeet,
                data.SpanLength,
                data.BoxSpan,
                data.BoxRise,
                data.PipeDiameter,
                data.NumberOfSpansOrCulverts,
                data.LowChordElevation,
                data.ChannelInvertElevation,
                data.WaterSurfaceElevation25Yr,
                data.HeadwaterToDiameterRatio25Yr,
                data.WaterSurfaceElevation100Yr,
                data.HeadwaterToDiameterRatio100Yr,
                data.WaterSurfaceElevation200Yr,
                data.HeadwaterToDiameterRatio200Yr
            };

            await _connection.ExecuteAsync(sql, parameters);
        }
    }
}
