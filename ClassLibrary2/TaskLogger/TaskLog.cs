using ClassLibrary2.Models;
using System;
using System.Collections.Generic;
using System.Data;
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


        //stored procedure insert final
        public Guid InsertFinalReport(string projectName, int userId, int statusId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_InsertFinalReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                // Add parameters
                cmd.Parameters.AddWithValue("@ProjectName", projectName);
                cmd.Parameters.AddWithValue("@UserID", userId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);

                conn.Open();

                // Because we used OUTPUT INSERTED.TaskID in SQL, ExecuteScalar() returns TaskID
                return (Guid)cmd.ExecuteScalar();
            }
        }


        //stored procedure update final
        public void UpdateFinalReport(Guid taskId, int statusId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_UpdateFinalReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@TaskID", taskId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);
                cmd.Parameters.AddWithValue("@CompletedOn", DBNull.Value); // SQL handles it

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }


        //stored procedure delete final
        public void DeleteFinalReport(Guid taskId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_DeleteFinalReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@TaskID", taskId);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }





        //stored procedure individual insert
        public int InsertIndividualReport(IndividualReportModel taskLog, string templateName, string statusMessage)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_InsertIndividualReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@TaskID", taskLog.TaskID);
                cmd.Parameters.AddWithValue("@UserID", taskLog.UserID);
                cmd.Parameters.AddWithValue("@StatusID", taskLog.StatusID);
                cmd.Parameters.AddWithValue("@TemplateName", templateName);
                cmd.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                cmd.Parameters.AddWithValue("@StatusMessage", statusMessage);

                conn.Open();
                return Convert.ToInt32(cmd.ExecuteScalar()); // SubtaskId returned
            }
        }



        //stored procedure individual update
        public void UpdateIndividualReport(int subtaskId, string currentStatus, string statusMessage)
        {

            int statusId = GetStatusIdByName(currentStatus);

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_UpdateIndividualReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@SubtaskId", subtaskId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);
                cmd.Parameters.AddWithValue("@StatusMessage", statusMessage);
                cmd.Parameters.AddWithValue("@CompletedOn", DateTime.Now);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }



        //stored procedure individual delete
        public void DeleteIndividualReport(int subtaskId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_DeleteIndividualReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@SubtaskId", subtaskId);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }










        // insert individual reports
        public void InsertIndividualReportTask(IndividualReportModel report,string userName,string currentStatus)
        {
            report.UserID = getUserIdByName(userName);

            report.StatusID = GetStatusIdByName(currentStatus);

            report.TaskID = CreateFinalReport(report.TemplateName, report.UserID, report.StatusID);

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"INSERT INTO xlChartGenerationPortal.IndividualReport
                              (TaskID, UserID, StatusID, TemplateName, CreatedOn, CompletedOn, StatusMessage) 
                              VALUES 
                              (@TaskID, @UserID, @StatusID, @TemplateName, @CreatedOn, @CompletedOn, @StatusMessage)";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@TaskID", report.TaskID);
                    cmd.Parameters.AddWithValue("@UserID", report.UserID);
                    cmd.Parameters.AddWithValue("@StatusID", report.StatusID);
                    cmd.Parameters.AddWithValue("@TemplateName", report.TemplateName);
                    cmd.Parameters.AddWithValue("@CreatedOn", report.CreatedOn);
                    cmd.Parameters.AddWithValue("@CompletedOn", report.CompletedOn);
                    cmd.Parameters.AddWithValue("@StatusMessage", report.StatusMessage);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }


        // update individual report status
        public void UpdateStatusForIndividualReportTask(IndividualReportModel report, string StatusMessage, string currentStatus)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                int statusId = GetStatusIdByName(currentStatus);

                string sql = @"UPDATE xlChartGenerationPortal.IndividualReport
                               SET CompletedOn = GETDATE(),
                                   StatusMessage = @StatusMessage,
                                   StatusID = @statusId
                                WHERE SubtaskId = @SubtaskId
                              ";

                using(SqlCommand cmd = new SqlCommand(sql,conn))
                {
                    cmd.Parameters.AddWithValue("@StatusMessage",report.StatusMessage);
                    cmd.Parameters.AddWithValue("@StatusID",statusId);
                    cmd.Parameters.AddWithValue("@SubtaskId", report.SubtaskId);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }


        // create guid for final report
        public Guid CreateFinalReport(string ProjectName, int UserID, int StatusID)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"INSERT INTO xlChartGenerationPortal.FinalReport 
                               (ProjectName, CreatedOn, UserID, StatusID)
                               OUTPUT INSERTED.TaskID
                               VALUES (@ProjectName, GETDATE(), @UserID, @StatusID)
                              ";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ProjectName", ProjectName);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@StatusID", StatusID);

                    conn.Open();
                    return (Guid)cmd.ExecuteScalar();
                }
            }
        }


        //update final report status
        public void UpdateStatusForFinalReportTask(string currentStatus,string username,string ProjectName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                int userId = getUserIdByName(username);

                Guid taskId = GetTaskIdByProjectAndUser(ProjectName, userId);

                int statusId = GetStatusIdByName(currentStatus);

                string sql = (currentStatus == "Done" || currentStatus == "Failed")
                            ? @"UPDATE xlChartGenerationPortal.FinalReport
                                   SET CompletedOn = GETDATE(),
                                       StatusID = @StatusID
                                 WHERE TaskID = @TaskID"
                            : @"UPDATE xlChartGenerationPortal.FinalReport
                                   SET StatusID = @StatusID
                                 WHERE TaskID = @TaskID";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@StatusID", statusId);
                    cmd.Parameters.AddWithValue("@TaskID", taskId);


                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }


        // get project details by project and user
        public Guid GetTaskIdByProjectAndUser(string ProjectName, int UserID)
        {
            using (SqlConnection conn = new SqlConnection(connectionString)) 
            {
                string sql = @" SELECT TOP 1 TaskID
                                FROM xlChartGenerationPortal.FinalReport
                                WHERE ProjectName = @ProjectName AND UserID = @UserID
                                ORDER BY CreatedOn DESC
                              ";

                using(SqlCommand cmd = new SqlCommand(sql,conn))
                {
                    cmd.Parameters.AddWithValue("@ProjectName", ProjectName);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    conn.Open();

                    object result = cmd.ExecuteScalar();
                    if (result == null || result == DBNull.Value)
                        throw new Exception($"No FinalReport found for project {ProjectName} and user {UserID}.");

                    return (Guid)result;
                }
            }
        }



        //get user ID by name
        public int getUserIdByName(string UserName)
        {
            using(SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT UserID FROM xlChartGenerationPortal.Users WHERE UserName=@UserName";

                using(SqlCommand cmd = new SqlCommand(sql,conn))
                {
                    cmd.Parameters.AddWithValue("@UserName", UserName);
                    conn.Open();

                    object result = cmd.ExecuteScalar();

                    if(result==null)
                    {
                        throw new Exception($"User {UserName} not found");
                    }

                    return Convert.ToInt32(result);
                }
            }
        }

        //get status id by name
        public int GetStatusIdByName(string StatusName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT StatusID FROM xlChartGenerationPortal.Status WHERE StatusName=@StatusName";

                using(SqlCommand cmd = new SqlCommand(sql,conn)) 
                {
                    cmd.Parameters.AddWithValue("@StatusName", StatusName);
                    conn.Open();

                    object result = cmd.ExecuteScalar();
                    if (result == null) 
                    {
                        throw new Exception($"Status {StatusName} not found");
                    }
                    return Convert.ToInt32(result);
                }
            }
        }

       


    }
}
 