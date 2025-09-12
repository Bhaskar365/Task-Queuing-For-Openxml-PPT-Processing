using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.SharedTaskTracking
{
    public class TaskTrack : ITaskTrack
    {
        string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=OpenxmlCharts;Integrated Security=True;";



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



    }
}
