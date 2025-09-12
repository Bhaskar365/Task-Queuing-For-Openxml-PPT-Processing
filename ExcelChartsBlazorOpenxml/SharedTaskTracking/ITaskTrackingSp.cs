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


        //individual delete
        void DeleteIndividualReport(int subtaskId);


        Guid GetTaskIdByProjectAndUser(string ProjectName, int UserID);
    }
}
