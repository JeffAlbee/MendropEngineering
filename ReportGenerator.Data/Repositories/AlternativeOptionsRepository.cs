using Dapper;
using ReportGenerator.Data.Models;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Data;

namespace ReportGenerator.Data.Repositories
{
    public class AlternativeOptionsRepository : IAlternativeOptionsRepository
    {
        private readonly IDbConnection _connection;

        public AlternativeOptionsRepository(IDbConnection connection)
        {
            _connection = connection;
        }

        public async Task UpsertAlternative1Async(Alternative1Option entity)
        {
            await UpsertAsync("dbo.Alternative1Options", entity);
        }

        public async Task UpsertAlternative2Async(Alternative2Option entity)
        {
            await UpsertAsync("dbo.Alternative2Options", entity);
        }

        public async Task UpsertAlternative3Async(Alternative3Option entity)
        {
            await UpsertAsync("dbo.Alternative3Options", entity);
        }

        public async Task UpsertAlternative4Async(Alternative4Option entity)
        {
            await UpsertAsync("dbo.Alternative4Options", entity);
        }

        private async Task UpsertAsync(string table, AlternativeOptionBase entity)
        {
            const string sqlTemplate = @"
                MERGE {TABLE} AS Target
                USING (SELECT @ReportID AS ReportID) AS Source
                    ON Target.ReportID = Source.ReportID
                WHEN MATCHED THEN
                    UPDATE SET 
                        StructureType = @StructureType,
                        StructureDescription = @StructureDescription,
                        StructureLength = @StructureLength,
                        StructureLengthUnit = @StructureLengthUnit,
                        PipeOrBoxLength = @PipeOrBoxLength,
                        PipeOrBoxLengthUnit = @PipeOrBoxLengthUnit,
                        BoxSpan = @BoxSpan,
                        BoxSpanUnit = @BoxSpanUnit,
                        BoxRise = @BoxRise,
                        BoxRiseUnit = @BoxRiseUnit,
                        PipeDiameter = @PipeDiameter,
                        PipeDiameterUnit = @PipeDiameterUnit,
                        CulvertOrSpanSize = @CulvertOrSpanSize,
                        CulvertOrSpanSizeUnit = @CulvertOrSpanSizeUnit,
                        NumberOfSpansOrCulverts = @NumberOfSpansOrCulverts,
                        TopOfCulvertElevation = @TopOfCulvertElevation,
                        LowChordElevation = @LowChordElevation,
                        WaterSurfaceElevation25Yr = @WaterSurfaceElevation25Yr,
                        WaterSurfaceElevation100Yr = @WaterSurfaceElevation100Yr,
                        WaterSurfaceElevation200Yr = @WaterSurfaceElevation200Yr,
                        HeadwaterToDiameterRatio25Yr = @HeadwaterToDiameterRatio25Yr,
                        HeadwaterToDiameterRatio100Yr = @HeadwaterToDiameterRatio100Yr,
                        HeadwaterToDiameterRatio200Yr = @HeadwaterToDiameterRatio200Yr
                WHEN NOT MATCHED THEN
                    INSERT (
                        ReportID, StructureType, StructureDescription, StructureLength, StructureLengthUnit,
                        PipeOrBoxLength, PipeOrBoxLengthUnit,
                        BoxSpan, BoxSpanUnit, BoxRise, BoxRiseUnit,
                        PipeDiameter, PipeDiameterUnit,
                        CulvertOrSpanSize, CulvertOrSpanSizeUnit,
                        NumberOfSpansOrCulverts, TopOfCulvertElevation, LowChordElevation,
                        WaterSurfaceElevation25Yr, WaterSurfaceElevation100Yr, WaterSurfaceElevation200Yr,
                        HeadwaterToDiameterRatio25Yr, HeadwaterToDiameterRatio100Yr, HeadwaterToDiameterRatio200Yr
                    )
                    VALUES (
                        @ReportID, @StructureType, @StructureDescription, @StructureLength, @StructureLengthUnit,
                        @PipeOrBoxLength, @PipeOrBoxLengthUnit,
                        @BoxSpan, @BoxSpanUnit, @BoxRise, @BoxRiseUnit,
                        @PipeDiameter, @PipeDiameterUnit,
                        @CulvertOrSpanSize, @CulvertOrSpanSizeUnit,
                        @NumberOfSpansOrCulverts, @TopOfCulvertElevation, @LowChordElevation,
                        @WaterSurfaceElevation25Yr, @WaterSurfaceElevation100Yr, @WaterSurfaceElevation200Yr,
                        @HeadwaterToDiameterRatio25Yr, @HeadwaterToDiameterRatio100Yr, @HeadwaterToDiameterRatio200Yr
                    );";

            var sql = sqlTemplate.Replace("{TABLE}", table);

            await _connection.ExecuteAsync(sql, entity);
        }
    }
}
