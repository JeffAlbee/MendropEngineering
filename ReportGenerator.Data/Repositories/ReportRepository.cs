using Dapper;
using ReportGenerator.Data.Models;
using ReportGenerator.Data.Repositories.Interfaces;
using System.Data;

namespace ReportGenerator.Data.Repositories
{
    public class ReportRepository : IReportRepository
    {
        #region Private Fields

        private readonly IDbConnection _connection;

        #endregion

        #region Constructor

        public ReportRepository(IDbConnection connection)
        {
            _connection = connection;
        }

        #endregion

        #region Public Methods

        public async Task<Report?> GetByIdAsync(Guid id)
        {
            const string sql = @"
                SELECT 
                    r.*, 
                    b.*, 
                    bc.*,
                    a1.*,
                    a2.*,
                    a3.*,
                    a4.*
                FROM dbo.Report r
                LEFT JOIN dbo.Bridge b ON r.ReportID = b.ReportID
                LEFT JOIN dbo.BridgeCharacteristics bc ON b.BridgeID = bc.BridgeID
                LEFT JOIN dbo.Alternative1Options a1 ON r.ReportID = a1.ReportID
                LEFT JOIN dbo.Alternative2Options a2 ON r.ReportID = a2.ReportID
                LEFT JOIN dbo.Alternative3Options a3 ON r.ReportID = a3.ReportID
                LEFT JOIN dbo.Alternative4Options a4 ON r.ReportID = a4.ReportID
                WHERE r.Id = @ReportID";

            var result = await _connection.QueryAsync<
                Report, 
                Bridge, 
                BridgeCharacteristics,
                Alternative1Option,
                Alternative2Option,
                Alternative3Option,
                Alternative4Option,
                Report>(
                sql,
                (report, bridge, characteristics, a1, a2, a3, a4) =>
                {
                    report.Bridge = bridge;
                    if (bridge != null)
                        bridge.Characteristics = characteristics;

                    report.Alternative1Option = a1;
                    report.Alternative2Option = a2;
                    report.Alternative3Option = a3;
                    report.Alternative4Option = a4;

                    return report;
                },
                new { ReportID = id },
                splitOn: "BridgeID,CharacteristicsID,AlternativeOptionID,AlternativeOptionID,AlternativeOptionID,AlternativeOptionID"
            );

            return result.FirstOrDefault();
        }

        public async Task<ReportBasicInfo?> GetBasicReportInfoAsync(Guid reportId)
        {
            const string sql = @"
                SELECT 
                    r.ReportID, 
                    r.ProjectNumber, 
                    b.BridgeID, 
                    b.BridgeCode
                FROM dbo.Report r
                LEFT JOIN dbo.Bridge b ON r.ReportID = b.ReportID
                WHERE r.Id = @ReportID";

            var result = await _connection.QueryAsync<ReportBasicInfo>(
                sql,
                new { ReportId = reportId }
            );

            return result.FirstOrDefault();
        }

        public async Task InsertReportImageAsync(ReportImage image)
        {
            const string sql = @"
                INSERT INTO dbo.ReportImage
                (Id, ReportId, Category, FileName, SharePointUrl, UploadedAt, UploadedBy)
                VALUES
                (@Id, @ReportId, @Category, @FileName, @SharePointUrl, @UploadedAt, @UploadedBy)";

            image.Id = Guid.NewGuid();
            image.UploadedAt = DateTime.UtcNow;

            await _connection.ExecuteAsync(sql, image);
        }

        #endregion
    }
}
