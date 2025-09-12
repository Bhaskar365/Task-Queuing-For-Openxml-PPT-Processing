using ClassLibrary2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.TaskLogger
{
    public interface ITaskLog
    {
        void InsertTask(TaskLogDLL task);

        void MarkTaskAsCompleted(string taskId, DateTime completedTime, string status);

        void SetTaskStatusState(Guid taskId, string status, string createdBy);

        Task<IEnumerable<TaskLogDLL>> GetUnfinishedTasks(string createdBy);

        Task<IEnumerable<TaskLogDLL>> GetAllData();


        void InsertIndividualReportTask(IndividualReportModel report, string userName, string currentStatus);

        void UpdateStatusForIndividualReportTask(IndividualReportModel report, string StatusMessage, string currentStatus);

        int getUserIdByName(string UserName);

        int GetStatusIdByName(string StatusName);

        Guid GetTaskIdByProjectAndUser(string ProjectName, int UserID);

        //stored procedures
        //final insertion
        Guid InsertFinalReport(string projectName, int userId, int statusId);


        //final update
        void UpdateFinalReport(Guid taskId, int statusId);


        //final delete
        void DeleteFinalReport(Guid taskId);



        //////////////////////////// /////////////////individual reports //////////////////////////// ////////////////////////////


        //individual insertion
        int InsertIndividualReport(IndividualReportModel taskLog, string templateName, string statusMessage);


        //individual update
        void UpdateIndividualReport(int subtaskId, string currentStatus, string statusMessage);


        //individual delete
        void DeleteIndividualReport(int subtaskId);


    }
}
