namespace WebApplicationAPI.TaskTracking
{
    public interface ITaskStatusTracker
    {
        void SetStatus(Guid taskId, string status);
        string GetStatus(Guid taskId);
    }
}
