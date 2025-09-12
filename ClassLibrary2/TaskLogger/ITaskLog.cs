using ClassLibrary2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.TaskLogger
{
    public interface ITaskLog
    {
        void InsertTask(TaskLogDLL task);

        void MarkTaskAsCompleted(string taskId, DateTime completedTime, string status);

        void SetTaskStatusState(Guid taskId, string status, string createdBy);

        Task<IEnumerable<TaskLogDLL>> GetUnfinishedTasks(string createdBy);

        Task<IEnumerable<TaskLogDLL>> GetAllData();
    }
}
