using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.TaskTrackingDLL
{
    public interface IPageTracking
    {
        void SetStatus(Guid taskId, string status);
        string GetStatus(Guid taskId);
    }
}
