using ClassLibrary2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClassLibrary2.BackgroundServicesDLL
{
    public interface IBackgroundServiceDLL
    {
        void Enqueue(PageStatusDto reportStatus);

        ValueTask EnqueueAsync(Func<CancellationToken, Task> workItem);
        ValueTask<Func<CancellationToken, Task>> DequeueAsync(CancellationToken cancellationToken);

        Task<PageStatusDto> DequeueAsync1(CancellationToken cancellationToken);

        IEnumerable<PageStatusDto> GetAll();
        PageStatusDto GetById(Guid id);
    }
}
