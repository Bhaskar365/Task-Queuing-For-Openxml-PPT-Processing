using System.Data.SqlClient;
using System.Data;
using SharedModels.DTO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelChartsBlazorOpenxml.SharedTaskTracking
{
    public class TaskTrackingSp : ITaskTrackingSp
    {
        //string connectionString = "Data Source=SQL02;Initial Catalog=BI_Methodology;Persist Security Info=True;User ID=mrcharts;Password='pwdmrchae0_d';Connect Timeout=0;TrustServerCertificate=True;";

        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=OpenxmlCharts;Integrated Security=True;";

        //stored procedure insert final
        public Guid InsertFinalReport(string projectName, string userName, string currentStatus)
        {
            try
            {
                int userId = GetUserIdByNameSp(userName);

                int statusId = GetStatusIdByNameSp(currentStatus);


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
            catch (Exception)
            {

                throw;
            }
        }


        //stored procedure update final
        public void UpdateFinalReport(string status, string projectName, string userName)
        {
            int statusId = GetStatusIdByNameSp(status);

            int userId = GetUserIdByNameSp(userName);

            Guid taskId = GetTaskIdByProjectAndUserSp(projectName, userId);


            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_UpdateFinalReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@TaskID", taskId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);
                cmd.Parameters.AddWithValue("@CompletedOn", DateTime.Now); // SQL handles it

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
        public int InsertIndividualReport(Guid taskId, int userId, int statusId, string templateName, string statusMessage)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_InsertIndividualReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@TaskID", taskId);
                cmd.Parameters.AddWithValue("@UserID", userId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);
                cmd.Parameters.AddWithValue("@TemplateName", templateName);
                cmd.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                cmd.Parameters.AddWithValue("@StatusMessage", statusMessage);

                conn.Open();
                return Convert.ToInt32(cmd.ExecuteScalar()); // SubtaskId returned
            }
        }



        //stored procedure individual update
        public void UpdateIndividualReport(int subtaskId, int statusId, string statusMessage)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_UpdateIndividualReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@SubtaskId", subtaskId);
                cmd.Parameters.AddWithValue("@StatusID", statusId);
                cmd.Parameters.AddWithValue("@CompletedOn", DateTime.Now);
                cmd.Parameters.AddWithValue("@StatusMessage", statusMessage);

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

        //get user ID by name
        public int getUserIdByName(string UserName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT UserID FROM xlChartGenerationPortal.Users WHERE UserName=@UserName";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@UserName", UserName);
                    conn.Open();

                    object result = cmd.ExecuteScalar();

                    if (result == null)
                    {
                        throw new Exception($"User {UserName} not found");
                    }

                    return Convert.ToInt32(result);
                }
            }
        }


        //get user id from name stored procedures
        public int GetUserIdByNameSp(string UserName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetUserByName", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", UserName);

                conn.Open();

                object result = cmd.ExecuteScalar();

                if (result == null || result == DBNull.Value)
                    throw new Exception($"No status was found for {UserName}");

                return Convert.ToInt16(result);
            }
        }


        //get status id from name stored procedure
        public int GetStatusIdByNameSp(string StatusName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetStatusIdFromName", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StatusName", StatusName);

                conn.Open();

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    throw new Exception($"No status was found for {StatusName}");

                return Convert.ToInt16(result);
            }
        }


        //get status name by id stored procedure
        public string GetUserNameFromIdSp(int userId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetUserNameById", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserID", userId);

                conn.Open();

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    throw new Exception($"No status was found for {userId}");

                return Convert.ToString(result)!;
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

                using (SqlCommand cmd = new SqlCommand(sql, conn))
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



        // get project details by project and user stored procedure
        public Guid GetTaskIdByProjectAndUserSp(string projectName, int userId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetTaskIdByProjectAndUser", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ProjectName", projectName);
                cmd.Parameters.AddWithValue("@UserID", userId);

                conn.Open();

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    throw new Exception($"No FinalReport found for project {projectName} and user {userId}.");

                return (Guid)result;
            }
        }

        // get individual user report by task id stored procedure
        public async Task<List<IndividualReportModel>> GetIndividualUserReport(Guid taskId)
        {
            var reports = new List<IndividualReportModel>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetIndividualUserReports", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@TaskID", taskId);

               await conn.OpenAsync();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    while (reader.Read())
                    {
                        reports.Add(new IndividualReportModel
                        {
                            SubtaskId = reader.GetInt32(reader.GetOrdinal("SubtaskId")),
                            TaskID = reader.GetGuid(reader.GetOrdinal("TaskID")),
                            UserID = reader.GetInt16(reader.GetOrdinal("UserID")),
                            StatusID = reader.GetInt16(reader.GetOrdinal("StatusID")),
                            TemplateName = reader.GetString(reader.GetOrdinal("TemplateName")),
                            CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                            CompletedOn = (DateTime)(reader.IsDBNull(reader.GetOrdinal("CompletedOn")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("CompletedOn"))),
                            StatusMessage = reader.IsDBNull(reader.GetOrdinal("StatusMessage"))
                                            ? null
                                            : reader.GetString(reader.GetOrdinal("StatusMessage"))
                        });
                    }
                }
            }

            return reports;
        }

        public List<FinalReportModel> GetFinalReportsByName(int UserID)
        {
            try
            {
                var reports = new List<FinalReportModel>();

                using (SqlConnection conn = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetAllFinalReportsByName", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    conn.Open();

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            reports.Add(new FinalReportModel
                            {
                                UserID = reader.GetInt16(reader.GetOrdinal("UserID")),
                                TaskID = reader.GetGuid(reader.GetOrdinal("TaskID")),
                                StatusID = reader.GetInt16(reader.GetOrdinal("StatusID")),
                                CompletedOn = (reader.IsDBNull(reader.GetOrdinal("CompletedOn")) ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("CompletedOn"))),
                                CreatedOn = reader.GetDateTime(reader.GetOrdinal("CreatedOn")),
                                ProjectName = reader.GetString(reader.GetOrdinal("ProjectName"))
                            });
                        }
                    }
                }
                return reports;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public string GetStatusNameFromIdSp(int StatusID)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand("xlChartGenerationPortal.sp_GetStatusNameFromId", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@StatusID", StatusID);

                conn.Open();

                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                    throw new Exception($"No status was found for {StatusID}");

                return Convert.ToString(result)!;
            }
        }


    }
}


