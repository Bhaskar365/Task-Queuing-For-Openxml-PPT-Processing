using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.TaskTrackingDLL
{
    public class PageTracking : IPageTracking
    {
        private readonly ConcurrentDictionary<Guid, string> _status = new ConcurrentDictionary<Guid, string>();
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
