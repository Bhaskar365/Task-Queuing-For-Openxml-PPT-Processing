using SharedModels.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.Task
{
    public interface ITask
    {
        void InsertTaskDto(TaskLog task, string connectionString);

        void MarkTaskAsCompletedDto(string taskId, DateTime completedTime, string connectionString, string status);

        void SetTaskStatusStateDto(Guid taskId, string status, string connectionString, string createdBy);

        Task<IEnumerable<TaskLog>> GetUnfinishedTasksDto(string connectionString, string createdBy);

        Task<IEnumerable<TaskLog>> GetAllDataDto(string connectionString);
    }
}
