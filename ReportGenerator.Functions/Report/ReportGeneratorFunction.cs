using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Services.Interfaces;
using System.Net;
using System.Web;

namespace ReportGenerator.Functions.Report;

public class ReportGeneratorFunction
{
    #region private fields

    private readonly IReportService _reportService;
    private readonly ILogger<ReportGeneratorFunction> _logger;

    #endregion

    #region Constructor

    public ReportGeneratorFunction(IReportService reportService, ILogger<ReportGeneratorFunction> logger)
    {
        _reportService = reportService;
        _logger = logger;
    }

    #endregion

    #region Function Method

    [Function("report-generate")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get")] HttpRequestData req)
    {
        _logger.LogInformation("Report generation triggered.");

        var query = HttpUtility.ParseQueryString(req.Url.Query);

        if (!Guid.TryParse(query["reportId"], out var reportId))
            return req.CreateResponse(HttpStatusCode.BadRequest);

        string sharePointUrl = await _reportService.GenerateReportAsync(reportId);

        var response = req.CreateResponse(HttpStatusCode.OK);

        await response.WriteAsJsonAsync(new
        {
            message = "Report generated successfully.",
            sharePointUrl
        });

        _logger.LogInformation("Report uploaded and URL returned: {Url}", sharePointUrl);

        return response;
    }

    #endregion
}