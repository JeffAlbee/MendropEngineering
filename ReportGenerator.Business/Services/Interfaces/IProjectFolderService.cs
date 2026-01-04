namespace ReportGenerator.Business.Services.Interfaces
{
    public interface IProjectFolderService
    {
        Task EnqueueFolderCreationAsync(Guid reportId);

        Task CreateFullFolderStructureAsync(Guid reportId);
    }
}
