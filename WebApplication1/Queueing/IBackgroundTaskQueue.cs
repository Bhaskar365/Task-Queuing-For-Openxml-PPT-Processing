using SharedModels.DTO;

namespace WebApplicationAPI.Queueing
{
    public interface IBackgroundTaskQueue
    {
        void Enqueue(ReportStatusDto reportStatus);

        ValueTask EnqueueAsync(Func<CancellationToken, Task> workItem);
        ValueTask<Func<CancellationToken, Task>?> DequeueAsync(CancellationToken cancellationToken);

        Task<ReportStatusDto> DequeueAsync1(CancellationToken cancellationToken);

        IEnumerable<ReportStatusDto> GetAll();
        ReportStatusDto? GetById(Guid id);
    }
}
