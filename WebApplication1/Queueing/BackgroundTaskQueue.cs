using SharedModels.DTO;
using System.Collections.Concurrent;
using System.Threading.Channels;

namespace WebApplicationAPI.Queueing
{
    public class BackgroundTaskQueue : IBackgroundTaskQueue
    {

        private readonly Channel<Func<CancellationToken, Task>> _queue;
        private readonly Channel<ReportStatusDto> _queue2 = Channel.CreateUnbounded<ReportStatusDto>();
        private readonly ConcurrentDictionary<Guid, ReportStatusDto> _store = new();

        public BackgroundTaskQueue(int capacity = 100)
        {
            _queue = Channel.CreateBounded<Func<CancellationToken,Task>>(capacity);
        }

        public async ValueTask<Func<CancellationToken, Task>?> DequeueAsync(CancellationToken cancellationToken)
        {
            Console.WriteLine("Task started...");
            var workItem = await _queue.Reader.ReadAsync(cancellationToken);
            Console.WriteLine("Task finished.");
            return workItem;
        }

        public async Task<ReportStatusDto> DequeueAsync1(CancellationToken cancellationToken)
        {
            return await _queue2.Reader.ReadAsync(cancellationToken);
        }

        public void Enqueue(ReportStatusDto reportStatus)
        {
            _store[reportStatus.TaskId] = reportStatus;
            _queue2.Writer.TryWrite(reportStatus);
        }

        public async ValueTask EnqueueAsync(Func<CancellationToken, Task> workItem)
        {
            Console.WriteLine("Task started...");

            if(workItem == null) throw new ArgumentNullException(nameof(workItem));

            await _queue.Writer.WriteAsync(workItem);

            Console.WriteLine("Task finished.");
        }

        public IEnumerable<ReportStatusDto> GetAll() => _store.Values;

        public ReportStatusDto? GetById(Guid id)
        {
            _store.TryGetValue(id, out var task);
            return task;
        }

    }
}
