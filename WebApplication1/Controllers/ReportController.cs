using Azure.Core;
using ClassLibrary1;
using DocumentFormat.OpenXml.Office2016.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SharedModels.DTO;
using System.Threading.Tasks;
using WebApplication1.Repositories;
using WebApplicationAPI.Queueing;
using WebApplicationAPI.TaskLogging;
using WebApplicationAPI.TaskTracking;

namespace WebApplicationAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        private readonly IBackgroundTaskQueue _queue;
        private readonly ITaskStatusTracker _tracker;
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly IConfiguration _configuration;
        private readonly ITaskLogging _taskLogging;
        private readonly IDataRepository _repository;

        public ReportController(IBackgroundTaskQueue queue,
                                ITaskStatusTracker tracker,
                                IServiceScopeFactory scopeFactory,
                                IConfiguration configuration,
                                ITaskLogging taskLogging,
                                IDataRepository repository)
        {
            _queue = queue;
            _tracker = tracker;
            _scopeFactory = scopeFactory;
            _configuration = configuration;
            _taskLogging = taskLogging;
            _repository = repository;
        }




        [HttpPost("generate")]
        public async Task<IActionResult> GenerateReport([FromBody] ReportGenerationRequest request)
        {

            var connectionString = _configuration.GetConnectionString("DBConnection");
            string user = "testUser";

            _tracker.SetStatus(request.TaskId, "Queued");

            await _queue.EnqueueAsync(async token =>
            {
                try
                {
                    _tracker.SetStatus(request.TaskId, "Processing");

                    using var scope = _scopeFactory.CreateScope();
                    var repository = scope.ServiceProvider.GetRequiredService<IDataRepository>();
                    var dll = new DLLCls();
                    string sourcePath = "";

                    TaskLog taskLog = new TaskLog
                    {
                        TaskId = request.TaskId,
                        CreatedOn = DateTime.UtcNow,
                        ProjectType = request.ProjectTemplateType,
                        CreatedBy = user,
                        CurrentStatus = "Processing"
                    };


                    _taskLogging.InsertTask(taskLog, connectionString!);
                    _taskLogging.SetTaskStatusState(request.TaskId, "Processing", connectionString!, user);

                    switch (request.ProjectTemplateType)
                    {
                        case "FitToConcept":

                            var fitToConceptdata = await repository.GetAllFitToConceptData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept{fitToConceptdata.Count()}.pptx";
                            dll.FitToConceptMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), fitToConceptdata.ToList());
                            break;

                        case "OverallImpressions":
                            var overallData = await repository.GetAllOverallImpressionsData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\OverallImpressions{overallData.Count()}.pptx";
                            dll.OverallImpressionsMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), overallData.ToList());
                            break;

                        case "Att1":
                            var att1Data = await repository.GetAtt1Data();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation{att1Data.Count()}.pptx";
                            dll.Attribute1Method(CreateTargetPath(sourcePath, request.ProjectTemplateType), att1Data.ToList());
                            break;

                        case "Att2":
                            var att2Data = await repository.GetAtt2Data();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation{att2Data.Count()}.pptx";
                            dll.Attribute2Method(CreateTargetPath(sourcePath, request.ProjectTemplateType), att2Data.ToList());
                            break;

                        case "AttrAggr":
                            var attAggData = await repository.GetAttrEvalAggregData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluationAggregate{attAggData.Count()}.pptx";
                            dll.AttributeMethodForAttributeEvalAggreg(CreateTargetPath(sourcePath, request.ProjectTemplateType), attAggData.ToList());
                            break;

                        case "Memorability":
                            var memoData = await repository.GetMemorabilityData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Memorability{memoData.Count()}.pptx";
                            dll.MemorabilityMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), memoData.ToList());
                            break;

                        case "PersonalPref":
                            var persPrefData = await repository.GetPersonalPreferenceData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PersonalPreferences{persPrefData.Count()}.pptx";
                            dll.PersonalPreferencesMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), persPrefData.ToList());
                            break;

                        case "Suffix":
                            var suffData = await repository.GetSuffixData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Suffix{suffData.Count()}.pptx";
                            dll.SuffixMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), suffData.ToList());
                            break;

                        case "VerbalUnder":
                            var verbData = await repository.GetVerbalUnderstandingData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\VerbalUnderstanding-Bar{verbData.Count()}.pptx";
                            dll.VerbalUnderstandingBarMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), verbData.ToList());
                            break;

                        case "writtenUnd":
                            var writtData = await repository.GetWrittenUnderstandingData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\WrittenUnderstanding{writtData.Count()}.pptx";
                            dll.WrittenUnderstandingMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), writtData.ToList());
                            break;

                        case "Exaggerative":
                            var exagData = await repository.GetExagg();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative{exagData.Count()}.pptx";
                            dll.ExaggerativeMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), exagData.ToList());
                            break;
                    }

                    taskLog.CompletedOn = DateTime.UtcNow;
                    _taskLogging.MarkTaskAsCompleted(request.TaskId.ToString(), (DateTime)taskLog.CompletedOn, connectionString!,"Done");

                    _tracker.SetStatus(request.TaskId, "Done");

                     _taskLogging.SetTaskStatusState(request.TaskId, "Done", connectionString!, user);

                }
                catch (Exception ex)
                {
                    _taskLogging.SetTaskStatusState(request.TaskId, "Fail", connectionString!, user);
                    _tracker.SetStatus(request.TaskId, $"Error: {ex.Message}");
                }
            });

            //await Task.Run(async () =>
            //{
            //    _tracker.SetStatus(request.TaskId, "Processing");
            //    var data = await _repository.GetAllFitToConceptData();
            //    var dll = new DLLCls();
            //    string sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept{data.Count()}.pptx";
            //    dll.FitToConceptMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), data.ToList());
            //    _tracker.SetStatus(request.TaskId, "Done");
            //});

            //return Ok(new ReportStatusDto { TaskId = request.TaskId, 
            //    ProjectType = request.ProjectTemplateType, Status= "Done" });

            return Ok(new ReportStatusDto
            {
                TaskId = request.TaskId,
                ProjectType = request.ProjectTemplateType,
            });
        }

        [HttpGet("status/{taskId}")]
        public IActionResult GetStatus(Guid taskId)
        {
            var status = _tracker.GetStatus(taskId);
            //return Ok(new { taskId, status });

            return Ok(new ReportStatusDto
            {
                TaskId = taskId,
                Status = status ?? "Unknown",
            });

        }

        [HttpGet("user/taskLogs")]
        public async Task<IActionResult> GetTaskLogsDetails()
        {
            var logs = await _repository.GetTaskLogs();
            return Ok(logs);
        }

        [HttpGet("{user}/logs")]
        public async Task<ActionResult<IEnumerable<TaskLog>>> GetUserSpecificLogs(string user)
        {
            try
            {
                string conn = _configuration.GetConnectionString("DBConnection")!;

                var userLogs = await _taskLogging.GetUnfinishedTasks(conn, user);
                return Ok(userLogs);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpPost("retry")]
        public async Task<IActionResult> RetryReport([FromBody] TaskLog task) 
        {
            var connectionString = _configuration.GetConnectionString("DBConnection");
            
            string user = task.CreatedBy;
            Guid taskId = task.TaskId;

            _tracker.SetStatus(task.TaskId, "Queued");
            _taskLogging.SetTaskStatusState(taskId, "Queued", connectionString!, user);

            await _queue.EnqueueAsync(async token => 
            {
                try
                {
                    _taskLogging.SetTaskStatusState(taskId, "Processing", connectionString!, user);
                    _tracker.SetStatus(task.TaskId,"Processing");

                    using var scope = _scopeFactory.CreateScope();
                    var repository = scope.ServiceProvider.GetRequiredService<IDataRepository>();
                    var dll = new DLLCls();
                    string sourcePath = "";

                    switch (task.ProjectType)
                    {
                        case "FitToConcept":

                            var fitToConceptdata = await repository.GetAllFitToConceptData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept{fitToConceptdata.Count()}.pptx";
                            dll.FitToConceptMethod(CreateTargetPath(sourcePath, task.ProjectType), fitToConceptdata.ToList());
                            break;

                        case "OverallImpressions":
                            var overallData = await repository.GetAllOverallImpressionsData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\OverallImpressions{overallData.Count()}.pptx";
                            dll.OverallImpressionsMethod(CreateTargetPath(sourcePath, task.ProjectType), overallData.ToList());
                            break;

                        case "Att1":
                            var att1Data = await repository.GetAtt1Data();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation{att1Data.Count()}.pptx";
                            dll.Attribute1Method(CreateTargetPath(sourcePath, task.ProjectType), att1Data.ToList());
                            break;

                        case "Att2":
                            var att2Data = await repository.GetAtt2Data();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation{att2Data.Count()}.pptx";
                            dll.Attribute2Method(CreateTargetPath(sourcePath, task.ProjectType), att2Data.ToList());
                            break;

                        case "AttrAggr":
                            var attAggData = await repository.GetAttrEvalAggregData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluationAggregate{attAggData.Count()}.pptx";
                            dll.AttributeMethodForAttributeEvalAggreg(CreateTargetPath(sourcePath, task.ProjectType), attAggData.ToList());
                            break;

                        case "Memorability":
                            var memoData = await repository.GetMemorabilityData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Memorability{memoData.Count()}.pptx";
                            dll.MemorabilityMethod(CreateTargetPath(sourcePath, task.ProjectType), memoData.ToList());
                            break;

                        case "PersonalPref":
                            var persPrefData = await repository.GetPersonalPreferenceData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PersonalPreferences{persPrefData.Count()}.pptx";
                            dll.PersonalPreferencesMethod(CreateTargetPath(sourcePath, task.ProjectType), persPrefData.ToList());
                            break;

                        case "Suffix":
                            var suffData = await repository.GetSuffixData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Suffix{suffData.Count()}.pptx";
                            dll.SuffixMethod(CreateTargetPath(sourcePath, task.ProjectType), suffData.ToList());
                            break;

                        case "VerbalUnder":
                            var verbData = await repository.GetVerbalUnderstandingData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\VerbalUnderstanding-Bar{verbData.Count()}.pptx";
                            dll.VerbalUnderstandingBarMethod(CreateTargetPath(sourcePath, task.ProjectType), verbData.ToList());
                            break;

                        case "writtenUnd":
                            var writtData = await repository.GetWrittenUnderstandingData();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\WrittenUnderstanding{writtData.Count()}.pptx";
                            dll.WrittenUnderstandingMethod(CreateTargetPath(sourcePath, task.ProjectType), writtData.ToList());
                            break;

                        case "Exaggerative":
                            var exagData = await repository.GetExagg();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative{exagData.Count()}.pptx";
                            dll.ExaggerativeMethod(CreateTargetPath(sourcePath, task.ProjectType), exagData.ToList());
                            break;
                    }

                    task.CompletedOn = DateTime.UtcNow;
                    _taskLogging.MarkTaskAsCompleted(task.TaskId.ToString(), (DateTime)task.CompletedOn, connectionString!, "Done");
                    
                    _taskLogging.SetTaskStatusState(taskId, "Done", connectionString!, user);

                    _tracker.SetStatus(taskId, "Done");


                }
                catch (Exception ex)
                {
                    _taskLogging.SetTaskStatusState(taskId, "Fail", connectionString!, user);
                    _tracker.SetStatus(taskId, $"Fail: {ex.Message}");
                }
            });


            return Ok(task);
        }

        [HttpPost("remove")]
        public async Task<IActionResult> RemoveFailedTask([FromBody] TaskLog task)
        {
            var connectionString = _configuration.GetConnectionString("DBConnection");

            var getAllTaskIds = await _taskLogging.GetAllData(connectionString!);

            List<TaskLog> taskLogs = new List<TaskLog>();

            if(getAllTaskIds!=null)
            {
                taskLogs = getAllTaskIds.Where(x => x.TaskId == task.TaskId).ToList();
            }
            return Ok(taskLogs);
        }


        private string CreateTargetPath(string template, string project)
        {
            string path = $"C:\\ExcelChartFiles\\{project}";
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            path = $"C:\\Users\\bdas\\Downloads\\{project}_sample.pptx";
            System.IO.File.Copy(template, path, true);
            return path;
        }
    }
}
