using System.Data.SqlClient;
using System.Data;

namespace ExcelChartsBlazorOpenxml.SharedTaskTracking
{
    public class TaskTrackingSp : ITaskTrackingSp
    {
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=OpenxmlCharts;Integrated Security=True;";



        //stored procedure insert final
        public Guid InsertFinalReport(string projectName,string userName,string currentStatus)
        {

            int userId = getUserIdByName(userName);

            int statusId = GetStatusIdByName(currentStatus);


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
        public void UpdateFinalReport(string status,string projectName,string userName)
        {
            int statusId = GetStatusIdByName(status);

            int userId = getUserIdByName(userName);

            Guid taskId = GetTaskIdByProjectAndUser(projectName, userId);


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



        //get status id by name
        public int GetStatusIdByName(string StatusName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"SELECT StatusID FROM xlChartGenerationPortal.Status WHERE StatusName=@StatusName";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
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


    }
}

