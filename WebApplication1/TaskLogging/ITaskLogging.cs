using SharedModels.DTO;

namespace WebApplicationAPI.TaskLogging
{
    public interface ITaskLogging
    {
        void InsertTask(TaskLog task, string connectionString);

        void MarkTaskAsCompleted(string taskId, DateTime completedTime, string connectionString,string status);

        void SetTaskStatusState(Guid taskId, string status, string connectionString, string createdBy);

        Task<IEnumerable<TaskLog>> GetUnfinishedTasks(string connectionString, string createdBy);

        Task<IEnumerable<TaskLog>> GetAllData(string connectionString);
    }
}

