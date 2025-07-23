using Microsoft.Data.SqlClient;
using SharedModels.DTO;

namespace WebApplicationAPI.TaskLogging
{
    public class TaskLogging : ITaskLogging
    {
        public void InsertTask(TaskLog task, string connectionString)
        {
            using (SqlConnection conn = new SqlConnection(connectionString)) 
            {
                string sql = @"INSERT INTO TaskLoggingTable (ProjectType, TaskId, CreatedOn, CreatedBy) 
                              VALUES(@ProjectType, @TaskId, @CreatedOn, @CreatedBy)";

                using (SqlCommand cmd = new SqlCommand(sql, conn)) 
                {
                    cmd.Parameters.AddWithValue("@ProjectType", task.ProjectType);
                    cmd.Parameters.AddWithValue("@TaskId", task.TaskId);
                    cmd.Parameters.AddWithValue("@CreatedOn", task.CreatedOn);
                    cmd.Parameters.AddWithValue("@CreatedBy", task.CreatedBy);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                
            };
        }

        public void MarkTaskAsCompleted(string taskId, DateTime completedTime, string connectionString)
        {
            using(SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE TaskLoggingTable 
                               SET CompletedOn = @CompletedOn
                               WHERE TaskId = @TaskId;
                              ";

                using (SqlCommand cmd = new SqlCommand(sql, conn))  
                {
                    cmd.Parameters.AddWithValue("@CompletedOn", completedTime);
                    cmd.Parameters.AddWithValue("@TaskId", Guid.Parse(taskId));

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
