
using Microsoft.Extensions.Logging;

namespace WebApplicationAPI.Queueing.HostedService
{
    public class QueuedHostedService : BackgroundService
    {
        private readonly IBackgroundTaskQueue _taskQueue;
        private readonly ILogger<QueuedHostedService> _logger;

        public QueuedHostedService(IBackgroundTaskQueue taskQueue, ILogger<QueuedHostedService> logger)
        {
            _taskQueue = taskQueue;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("Queued Hosted Service is running.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    _logger.LogInformation("Dequeued a task...");

                    var workItem = await _taskQueue.DequeueAsync(stoppingToken);

                    if (workItem != null)
                    {
                        // await workItem(stoppingToken);
                        _ = Task.Run(() => workItem(stoppingToken), stoppingToken);

                        _logger.LogInformation("Task completed.");
                    }

                }
                catch (OperationCanceledException ex)
                {
                    _logger.LogCritical(ex.Message);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error occurred executing task work item.");
                }
              
            }
        }
    }
}