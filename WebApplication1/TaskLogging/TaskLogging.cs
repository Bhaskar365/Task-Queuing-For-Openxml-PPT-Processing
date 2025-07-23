using Microsoft.AspNetCore.Http.HttpResults;
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

        public void SetTaskStatusState(Guid taskId,string status,string connectionString,string createdBy) 
        {
            using(SqlConnection conn = new SqlConnection(connectionString)) 
            {
                string sql = @"UPDATE TaskLoggingTable
                               SET status = @CurrentStatus
                               WHERE TaskId = @TaskId
                               AND CreatedBy=@CreatedBy
                              ";

                using(SqlCommand cmd = new SqlCommand(sql,conn))
                {
                    cmd.Parameters.AddWithValue("@CurrentStatus",status);
                    cmd.Parameters.AddWithValue("@TaskId",taskId);
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }   
        }

        public async Task<IEnumerable<TaskLog>> GetUnfinishedTasks(string connectionString,string createdBy)
        {
            var tasks = new List<TaskLog>();

            using (SqlConnection conn = new SqlConnection(connectionString)) 
            {
                string sql = @"SELECT * FROM TaskLoggingTable 
                               WHERE (Status = 'Queued' OR Status = 'Processing') 
                               AND CompletedOn IS NULL AND CreatedBy=@CreatedBy"; 

                using(SqlCommand cmd = new SqlCommand(sql,conn))
                {
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy);

                    await conn.OpenAsync();

                    using (var reader = await cmd.ExecuteReaderAsync()) 
                    {
                        while(await reader.ReadAsync())
                        {
                            tasks.Add(new TaskLog 
                            {
                                TaskId = reader.GetGuid(reader.GetOrdinal("TaskId")),
                                ProjectType = reader["ProjectType"].ToString()!,
                                CreatedBy = reader["CreatedBy"].ToString()!,
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                CompletedOn = reader["CompletedOn"] == DBNull.Value ? DateTime.MinValue : reader.GetDateTime(reader.GetOrdinal("CompletedOn"))
                            });
                        }
                    }
                }
            }
            return tasks;
        }

    }
}
