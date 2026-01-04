using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using ReportGenerator.Business.Models;
using ReportGenerator.Business.Services.Interfaces;
using System.Net;

namespace ReportGenerator.Functions.Report;

public class ProcessBridgeExcelFunction
{
    private readonly IBridgeExcelProcessor _bridgeExcelProcessor;
    private readonly ILogger<ProcessBridgeExcelFunction> _logger;

    public ProcessBridgeExcelFunction(ILogger<ProcessBridgeExcelFunction> logger, IBridgeExcelProcessor bridgeExcelProcessor)
    {
        _bridgeExcelProcessor = bridgeExcelProcessor;
        _logger = logger;
    }

    [Function("process-bridge-excel")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        _logger.LogInformation("Bridge Excel processing triggered.");

        var request = await req.ReadFromJsonAsync<ProcessExcelRequest>();

        if (request == null || request.ReportId == Guid.Empty || string.IsNullOrWhiteSpace(request.FileUrl))
        {
            _logger.LogWarning("Invalid request payload.");
            var badResponse = req.CreateResponse(HttpStatusCode.BadRequest);
            await badResponse.WriteStringAsync("Missing ReportId or FileUrl.");
            return badResponse;
        }

        _logger.LogInformation("Request received for ReportId={ReportId}, FileUrl={Url}",
            request.ReportId, request.FileUrl);

        var result = await _bridgeExcelProcessor.ProcessAsync(request.ReportId, request.FileUrl);

        var response = req.CreateResponse(HttpStatusCode.OK);
        await response.WriteAsJsonAsync(new
        {
            message = "Excel processed successfully",
            reportId = request.ReportId,
            alternatives = result.Alternatives,
            existing = result.Existing
        });

        return response;
    }
}