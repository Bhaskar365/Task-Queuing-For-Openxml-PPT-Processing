using SharedModels.DTO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.Task
{
    public class Task : ITask
    {
        public async Task<IEnumerable<TaskLog>> GetAllDataDto(string connectionString)
        {
            var tasks = new List<TaskLog>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT * FROM TaskLoggingTable WHERE CreatedBy IS NOT NULL";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    await conn.OpenAsync();

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            tasks.Add(new TaskLog
                            {
                                TaskId = reader.GetGuid(reader.GetOrdinal("TaskId")),
                                ProjectType = reader["ProjectType"].ToString()!,
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                CompletedOn = reader.GetDateTime(reader.GetOrdinal("CompletedOn")),
                                CreatedBy = reader["CreatedBy"].ToString()!,
                                CurrentStatus = reader["CurrentStatus"].ToString()!
                            });
                        }
                    }
                }
            }
            return tasks;
        }

        public async Task<IEnumerable<TaskLog>> GetUnfinishedTasksDto(string connectionString, string createdBy)
        {
            var tasks = new List<TaskLog>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT * FROM TaskLoggingTable 
                               WHERE (CurrentStatus = 'Fail') 
                               AND CompletedOn IS NOT NULL AND CreatedBy=@CreatedBy";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy);

                    await conn.OpenAsync();

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            tasks.Add(new TaskLog
                            {
                                TaskId = reader.GetGuid(reader.GetOrdinal("TaskId")),
                                ProjectType = reader["ProjectType"].ToString()!,
                                CreatedBy = reader["CreatedBy"].ToString()!,
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                CurrentStatus = reader["CurrentStatus"].ToString()!,
                                CompletedOn = reader.GetDateTime(reader.GetOrdinal("CompletedOn"))
                            });
                        }
                    }
                }
            }
            return tasks;
        }

        public void InsertTaskDto(TaskLog task, string connectionString)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"INSERT INTO TaskLoggingTable (ProjectType, TaskId, CreatedOn, CreatedBy, CurrentStatus) 
                              VALUES(@ProjectType, @TaskId, @CreatedOn, @CreatedBy, @CurrentStatus)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ProjectType", task.ProjectType);
                    cmd.Parameters.AddWithValue("@TaskId", task.TaskId);
                    cmd.Parameters.AddWithValue("@CreatedOn", task.CreatedOn);
                    cmd.Parameters.AddWithValue("@CreatedBy", task.CreatedBy);
                    cmd.Parameters.AddWithValue("@CurrentStatus", task.CurrentStatus);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }

            }
        }

        public void MarkTaskAsCompletedDto(string taskId, DateTime completedTime, string connectionString, string status)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE TaskLoggingTable 
                               SET CompletedOn = @CompletedOn,
                                   CurrentStatus = @CurrentStatus
                               WHERE TaskId = @TaskId;
                              ";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@CompletedOn", completedTime);
                    cmd.Parameters.AddWithValue("@TaskId", Guid.Parse(taskId));
                    cmd.Parameters.AddWithValue("@CurrentStatus", status);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void SetTaskStatusStateDto(Guid taskId, string status, string connectionString, string createdBy)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE TaskLoggingTable
                               SET CurrentStatus = @CurrentStatus
                               WHERE TaskId = @TaskId
                               AND CreatedBy=@CreatedBy
                              ";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@CurrentStatus", status);
                    cmd.Parameters.AddWithValue("@TaskId", taskId);
                    cmd.Parameters.AddWithValue("@CreatedBy", createdBy);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
