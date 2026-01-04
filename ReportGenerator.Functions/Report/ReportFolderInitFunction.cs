using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using System.Net;

namespace ReportGenerator.Functions.Report;

public class ReportFolderInitFunction
{
    #region private fields

    private readonly IProjectFolderService _projectFolderService;
    private readonly ILogger<ReportFolderInitFunction> _logger;

    #endregion

    #region Constructor

    public ReportFolderInitFunction(
        IProjectFolderService projectFolderService,
        ILogger<ReportFolderInitFunction> logger)
    {
        _projectFolderService = projectFolderService;
        _logger = logger;
    }

    #endregion

    #region Function Methods

    [Function("report-create-folders")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        _logger.LogInformation("Report folder initialization triggered.");

        var request = await req.ReadFromJsonAsync<ReportFolderRequest>();

        if (request == null || request.ReportId == Guid.Empty)
        {
            var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
            await badResponse.WriteStringAsync("Missing or invalid reportId.");
            return badResponse;
        }

        await _projectFolderService.EnqueueFolderCreationAsync(request.ReportId);

        var response = req.CreateResponse(HttpStatusCode.OK);

        await response.WriteAsJsonAsync(new
        {
            message = "Folder creation request enqueued successfully.",
            reportId = request.ReportId
        });

        return response;
    }

    #endregion
}