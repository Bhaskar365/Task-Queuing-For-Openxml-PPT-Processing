using SharedModels.DTO;

namespace WebApplicationAPI.TaskLogging
{
    public interface ITaskLogging
    {
        void InsertTask(TaskLog task, string connectionString);
        void MarkTaskAsCompleted(string taskId, DateTime completedTime, string connectionString);
    }
}
