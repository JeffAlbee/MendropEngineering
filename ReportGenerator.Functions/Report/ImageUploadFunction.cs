using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using System.Net;

namespace ReportGenerator.Functions.Report;

public class ImageUploadFunction
{
    #region private fields

    private readonly IReportService _reportService;
    private readonly ILogger<ImageUploadFunction> _logger;

    #endregion

    #region Constructor

    public ImageUploadFunction(
        IReportService reportService,
        ILogger<ImageUploadFunction> logger
        )
    {
        _reportService = reportService;
        _logger = logger;
    }

    #endregion

    #region Function Methods

    [Function("upload-image-to-sharepoint")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        _logger.LogInformation("Upload image to sharepoint triggered.");

        var request = await req.ReadFromJsonAsync<ImageUploadRequest>();

        if (request == null || request.ReportId == Guid.Empty || string.IsNullOrWhiteSpace(request.FileUrl))
        {
            var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
            await badResponse.WriteStringAsync("Missing required fields.");
            return badResponse;
        }

        var uploadedUrl = await _reportService.UploadImageToSharePointAsync(request);

        var response = req.CreateResponse(HttpStatusCode.OK);

        await response.WriteAsJsonAsync(new
        {
            message = "Image uploaded successfully to SharePoint.",
            sharePointUrl = uploadedUrl
        });

        return response;
    }

    #endregion
}