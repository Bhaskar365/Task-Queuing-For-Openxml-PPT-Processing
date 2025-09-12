using ClassLibrary2.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.TaskLogger
{
    public class TaskLog : ITaskLog
    {
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=OpenxmlCharts;Integrated Security=True;";

        public void InsertTask(TaskLogDLL task)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"INSERT INTO PagesLoggingTable (ProjectType, TaskId, CreatedOn, CreatedBy, CurrentStatus) 
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

        public void MarkTaskAsCompleted(string taskId, DateTime completedTime, string status)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE PagesLoggingTable 
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

        public void SetTaskStatusState(Guid taskId, string status, string createdBy)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"UPDATE PagesLoggingTable
                               SET CurrentStatus = @CurrentStatus
                               WHERE TaskId = @TaskId
                               AND CreatedBy=@CreatedBy";

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

        public async Task<IEnumerable<TaskLogDLL>> GetUnfinishedTasks(string createdBy)
        {
            var tasks = new List<TaskLogDLL>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT * FROM PagesLoggingTable 
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
                            tasks.Add(new TaskLogDLL
                            {
                                TaskId = reader.GetGuid(reader.GetOrdinal("TaskId")),
                                ProjectType = reader["ProjectType"].ToString(),
                                CreatedBy = reader["CreatedBy"].ToString(),
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                CurrentStatus = reader["CurrentStatus"].ToString(),
                                CompletedOn = reader.GetDateTime(reader.GetOrdinal("CompletedOn"))
                            });
                        }
                    }
                }
            }
            return tasks;
        }

        public async Task<IEnumerable<TaskLogDLL>> GetAllData()
        {
            var tasks = new List<TaskLogDLL>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT * FROM PagesLoggingTable WHERE CreatedBy IS NOT NULL";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    await conn.OpenAsync();

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (await reader.ReadAsync())
                        {
                            tasks.Add(new TaskLogDLL
                            {
                                TaskId = reader.GetGuid(reader.GetOrdinal("TaskId")),
                                ProjectType = reader["ProjectType"].ToString(),
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                CompletedOn = reader.GetDateTime(reader.GetOrdinal("CompletedOn")),
                                CreatedBy = reader["CreatedBy"].ToString(),
                                CurrentStatus = reader["CurrentStatus"].ToString(),
                            });
                        }
                    }
                }
            }
            return tasks;
        }
    }
}
