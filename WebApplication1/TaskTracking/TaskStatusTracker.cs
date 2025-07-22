
using System.Collections.Concurrent;

namespace WebApplicationAPI.TaskTracking
{
    public class TaskStatusTracker : ITaskStatusTracker
    {
        private readonly ConcurrentDictionary<Guid, string> _status = new();
        public string GetStatus(Guid taskId)
        {
           return _status.TryGetValue(taskId, out var status) ? status : "Not Found";
        }

        public void SetStatus(Guid taskId, string status)
        {
            _status[taskId] = status;
        }
    }
}
