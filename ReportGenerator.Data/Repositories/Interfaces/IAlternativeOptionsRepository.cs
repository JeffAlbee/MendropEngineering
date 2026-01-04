using ReportGenerator.Data.Models;

namespace ReportGenerator.Data.Repositories.Interfaces
{
    public interface IAlternativeOptionsRepository
    {
        Task UpsertAlternative1Async(Alternative1Option entity);

        Task UpsertAlternative2Async(Alternative2Option entity);

        Task UpsertAlternative3Async(Alternative3Option entity);

        Task UpsertAlternative4Async(Alternative4Option entity);
    }
}
