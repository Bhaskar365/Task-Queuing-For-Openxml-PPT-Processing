using ClassLibrary2.Models;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace ClassLibrary2.BackgroundServicesDLL
{
    public class BackgroundServiceDLL : IBackgroundServiceDLL
    {
        private readonly Channel<Func<CancellationToken, Task>> _queue;
        private readonly Channel<PageStatusDto> _queue2 = Channel.CreateUnbounded<PageStatusDto>();
        private readonly ConcurrentDictionary<Guid, PageStatusDto> _store = new ConcurrentDictionary<Guid, PageStatusDto>();

        public BackgroundServiceDLL(int capacity = 100)
        {
            _queue = Channel.CreateBounded<Func<CancellationToken, Task>>(capacity);
        }

        public async ValueTask<Func<CancellationToken, Task>> DequeueAsync(CancellationToken cancellationToken)
        {
            Console.WriteLine("Task started...");
            var workItem = await _queue.Reader.ReadAsync(cancellationToken);
            Console.WriteLine("Task finished.");
            return workItem;
        }

        public async Task<PageStatusDto> DequeueAsync1(CancellationToken cancellationToken)
        {
            return await _queue2.Reader.ReadAsync(cancellationToken);
        }

        public void Enqueue(PageStatusDto reportStatus)
        {
            _store[reportStatus.TaskId] = reportStatus;
            _queue2.Writer.TryWrite(reportStatus);
        }

        public async ValueTask EnqueueAsync(Func<CancellationToken, Task> workItem)
        {
            Console.WriteLine("Task started...");

            if (workItem == null) throw new ArgumentNullException(nameof(workItem));

            await _queue.Writer.WriteAsync(workItem);

            Console.WriteLine("Task finished.");
        }

        public IEnumerable<PageStatusDto> GetAll()
        {
            return _store.Values;
        }

        public PageStatusDto GetById(Guid id)
        {
            _store.TryGetValue(id, out var task);
            return task;
        }
    }
}
