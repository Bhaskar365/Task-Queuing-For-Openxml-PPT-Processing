using SharedModels.DTO;

namespace ExcelChartsBlazorOpenxml.SharedTaskTracking
{
    public interface ITaskTrackingSp
    {
        //stored procedures

        ////////////////////////////////////////////////////////final reports//////////////////////////////////////////////////////

        //final insertion
        Guid InsertFinalReport(string projectName, string userName, string currentStatus);



        //final update
        void UpdateFinalReport(string status, string projectName, string userName);


        //final delete
        void DeleteFinalReport(Guid taskId);



        //////////////////////////// /////////////////individual reports //////////////////////////// ////////////////////////////


        //individual insertion
        int InsertIndividualReport(Guid taskId, int userId, int statusId, string templateName, string statusMessage);


        //individual update
        void UpdateIndividualReport(int subtaskId, int statusId, string statusMessage);


        //get status id from name
        int GetStatusIdByNameSp(string StatusName);


        //individual delete
        void DeleteIndividualReport(int subtaskId);

        //get user id from name
        int GetUserIdByNameSp(string UserName);

        // get user  name from id stored procedure
        string GetUserNameFromIdSp(int userId);

        Guid GetTaskIdByProjectAndUser(string ProjectName, int UserID);

        Guid GetTaskIdByProjectAndUserSp(string projectName, int userId);

        List<IndividualReportModel> GetIndividualUserReport(Guid taskId);

    }
}
