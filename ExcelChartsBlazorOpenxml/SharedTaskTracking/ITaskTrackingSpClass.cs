namespace ExcelChartsBlazorOpenxml.SharedTaskTracking
{
    public interface ITaskTrackingSpClass
    {
        //final insertion
        Guid InsertFinalReport(string projectName, int userId, int statusId);



        //final update
        void UpdateFinalReport(Guid taskId, int statusId);


        //final delete
        void DeleteFinalReport(Guid taskId);



        //individual insertion
        int InsertIndividualReport(Guid taskId, int userId, int statusId, string templateName, string statusMessage);


        //individual update
        void UpdateIndividualReport(int subtaskId, int statusId, string statusMessage);


        //individual delete
        void DeleteIndividualReport(int subtaskId);
    }
}
