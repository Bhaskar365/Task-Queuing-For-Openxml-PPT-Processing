using Azure.Core;
using ClassLibrary1;
using DocumentFormat.OpenXml.Office2016.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using SharedModels;
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

                        case "QTC":
                            var qtcData = await repository.GetQTC();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTC.pptx";
                            dll.QTCMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), qtcData.ToList());
                            break;

                        case "Sala":
                            var salaData = await repository.GetSala();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";
                            dll.SALANewMethod(CreateTargetPath(sourcePath, request.ProjectTemplateType), salaData.ToList());
                            break;

                        case "Brandex Safety":

                            var brandexSafetyData = (await repository.GetBrandexSafety()).ToList();

                            List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

                            double average1Max = 0.0000;
                            double average2Max = 0.0000;
                            double average3Max = 0.0;
                            double average4Max = 0.0000;
                            double average5Max = 0.0000;

                            for (int i = 0; i < brandexSafetyData.Count; i++)
                            {
                                var dblAverage1max = brandexSafetyData[i]?.dblAveragePage1;
                                if (dblAverage1max == 0.0 || dblAverage1max == null)
                                {
                                    average1Max = 0;
                                }
                                if (average1Max < dblAverage1max)
                                {
                                    average1Max = (double)dblAverage1max;
                                }

                                var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
                                if (dblAverage2max == 0.0 || dblAverage2max == null)
                                {
                                    average2Max = 0;
                                }
                                if (average2Max < dblAverage2max)
                                {
                                    average2Max = (double)dblAverage2max;
                                }

                                var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
                                if (dblAverage3max == 0.0 || dblAverage3max == null)
                                {
                                    average3Max = 0;
                                }
                                if (average3Max < dblAverage3max)
                                {
                                    average3Max = (double)dblAverage3max;
                                }

                                var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
                                if (dblAverage4max == 0.0 || dblAverage4max == null)
                                {
                                    average4Max = 0;
                                }
                                if (average4Max < dblAverage4max)
                                {
                                    average4Max = (double)dblAverage4max;
                                }

                                var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
                                if (dblAverage5max == 0.0 || dblAverage5max == null)
                                {
                                    average5Max = 0;
                                }
                                if (average5Max < dblAverage5max)
                                {
                                    average5Max = (double)dblAverage5max;
                                }
                            }

                            double scalingFactor = 0.22;

                            for (int i = 0; i < brandexSafetyData.Count; i++)
                            {
                                //for the table
                                BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

                                var dataEl = brandexSafetyData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;
                                double averagePage3WeightedValue = 0.0;
                                double averagePage4WeightedValue = 0.0;
                                double averagePage5WeightedValue = 0.0;

                                if (average1Max > 0)
                                {
                                    // averagePage1WeightedValue =(dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) ;

                                    averagePage1WeightedValue = (double)((double)(dataEl.dblAveragePage1 ?? 0.0 / average1Max) * (double)(dataEl.dblPage1Weight));
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (average2Max > 0)
                                {
                                    // averagePage2WeightedValue = (double)(dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;

                                    averagePage2WeightedValue = (double)((double)(dataEl.dblAveragePage2 ?? 0.0 / average2Max) * (double)(dataEl.dblPage2Weight));
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (average3Max > 0)
                                {
                                    averagePage3WeightedValue = (double)((double)(dataEl.dblAveragePage3 ?? 0.0 / average3Max) * (double)(dataEl.dblPage3Weight));
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (average4Max > 0)
                                {
                                    //averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;

                                    averagePage4WeightedValue = (double)((double)(dataEl.dblAveragePage4 ?? 0.0 / average4Max) * (double)(dataEl.dblPage4Weight));
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                if (average5Max > 0)
                                {
                                    // averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
                                    averagePage5WeightedValue = (double)((double)(dataEl.dblAveragePage5 ?? 0.0 / average5Max) * (double)(dataEl.dblPage5Weight));


                                }
                                else
                                {
                                    averagePage5WeightedValue = 0;
                                }


                                //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
                                //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
                                //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
                                // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
                                //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                                                  averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;


                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1 ?? 0.0;
                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2 ?? 0.0;
                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3 ?? 0.0;
                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4 ?? 0.0;
                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5 ?? 0.0;
                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6 ?? 0.0;
                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage6Weighted = (double)dataEl.dblAveragePage6Weighted;

                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7 ?? 0.0;
                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage7Weighted = (double)dataEl.dblAveragePage7Weighted;

                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8 ?? 0.0;
                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage8Weighted = (double)dataEl.dblAveragePage8Weighted;

                                brandexSafetyShortModel.dblIndex = indexSum;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = (int)dataEl.intRed;
                                brandexSafetyShortModel.intGreen = (int)dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = (int)dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = (bool)dataEl.boolBold;


                                //for the chart - 

                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;
                                double averagePage3WeightedValueForChart = 0.0;
                                double averagePage4WeightedValueForChart = 0.0;
                                double averagePage5WeightedValueForChart = 0.0;

                                if (average1Max > 0)
                                {
                                    averagePage1WeightedValueForChart = (double)((dataEl.dblAveragePage1 ?? 0.0 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (average2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = (double)((dataEl.dblAveragePage2 ?? 0.0 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (average3Max > 0)
                                {
                                    averagePage3WeightedValueForChart = (double)((dataEl.dblAveragePage3 ?? 0.0 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (average4Max > 0)
                                {
                                    averagePage4WeightedValueForChart = (double)((dataEl.dblAveragePage4 ?? 0.0 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                if (average5Max > 0)
                                {
                                    averagePage5WeightedValueForChart = (double)((dataEl.dblAveragePage5 ?? 0.0 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage5WeightedValueForChart = 0;
                                }

                                //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
                                //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
                                //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
                                //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
                                //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

                                double indexSumForChart = averagePage1WeightedValueForChart +
                                                          averagePage2WeightedValueForChart +
                                                          averagePage3WeightedValueForChart +
                                                          averagePage4WeightedValueForChart +
                                                          averagePage5WeightedValueForChart;

                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

                                brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

                                brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

                                brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

                                brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

                                brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;

                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = (int)dataEl.intRed;
                                brandexSafetyShortModel.intGreen = (int)dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = (int)dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = (bool)dataEl.boolBold;


                                brandexData.Add(brandexSafetyShortModel);
                            }


                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific{brandexSafetyData.Count()}.pptx";
                            dll.BrandexSafetyMethod(CreateTargetPath(sourcePath, brandexSafetyData.First().ProjectTemplateType!), brandexData);
                            break;

                        case "Brandex Strategic Distinctiveness":

                            var brandexStrategicData = (await repository.GetBrandexStrategicDistinctivenessData()).ToList();

                            List<BrandexStrategicDistinctivenessShortModel> brandexStrategicDistinctivenessesShortData = new List<BrandexStrategicDistinctivenessShortModel>();

                            double dblAverage1Max = 0.0;
                            double dblAverage2Max = 0.0;
                            double dblAverage3Max = 0.0;
                            double dblAverage4Max = 0.0;

                            for (int i = 0; i < brandexStrategicData.Count; i++)
                            {
                                var dblAverage1MaxValue = brandexStrategicData[i].dblAveragePage1;
                                if (dblAverage1MaxValue == 0)
                                {
                                    dblAverage1Max = 0.0;
                                }
                                if (dblAverage1MaxValue > dblAverage1Max)
                                {
                                    dblAverage1Max = (double)dblAverage1MaxValue;
                                }

                                var dblAverage2MaxValue = brandexStrategicData[i].dblAveragePage2;
                                if (dblAverage2MaxValue == 0)
                                {
                                    dblAverage2Max = 0;
                                }
                                if (dblAverage2MaxValue > dblAverage2Max)
                                {
                                    dblAverage2Max = (double)dblAverage2MaxValue;
                                }

                                var dblAverage3MaxValue = brandexStrategicData[i].dblAveragePage3;
                                if (dblAverage3MaxValue == 0)
                                {
                                    dblAverage3Max = 0;
                                }
                                if (dblAverage3MaxValue > dblAverage3Max)
                                {
                                    dblAverage3Max = (double)dblAverage3MaxValue;
                                }

                                var dblAverage4MaxValue = brandexStrategicData[i].dblAveragePage4;
                                if (dblAverage4MaxValue == 0)
                                {
                                    dblAverage4Max = 0;
                                }
                                if (dblAverage4MaxValue > dblAverage4Max)
                                {
                                    dblAverage4Max = (double)dblAverage4MaxValue;
                                }
                            }

                            // scaling factor 
                            double scalingFactorForStrategicDistinctiveness = 1.00777;

                            for (int i = 0; i < brandexStrategicData.Count; i++)
                            {
                                // for the table
                                BrandexStrategicDistinctivenessShortModel brandexStrategicDistinctivenessModelData = new BrandexStrategicDistinctivenessShortModel();

                                var dataForDistinctivessModel = brandexStrategicData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;
                                double averagePage3WeightedValue = 0.0;
                                double averagePage4WeightedValue = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValue =
                                        (double)((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * dataForDistinctivessModel.dblPage1Weight)!;
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValue =
                                       (double)((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * dataForDistinctivessModel.dblPage2Weight)!;
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValue =
                                        (double)((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * dataForDistinctivessModel.dblPage3Weight)!;
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValue =
                                        (double)((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * dataForDistinctivessModel.dblPage4Weight)!;
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                                             averagePage4WeightedValue) * scalingFactorForStrategicDistinctiveness;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName!;

                                brandexStrategicDistinctivenessModelData.dblAveragePage1 = dataForDistinctivessModel.dblAveragePage1;
                                brandexStrategicDistinctivenessModelData.dblPage1Weight = dataForDistinctivessModel.dblPage1Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

                                brandexStrategicDistinctivenessModelData.dblAveragePage2 = dataForDistinctivessModel.dblAveragePage2;
                                brandexStrategicDistinctivenessModelData.dblPage2Weight = dataForDistinctivessModel.dblPage2Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

                                brandexStrategicDistinctivenessModelData.dblAveragePage3 = dataForDistinctivessModel.dblAveragePage3;
                                brandexStrategicDistinctivenessModelData.dblPage3Weight = dataForDistinctivessModel.dblPage3Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

                                brandexStrategicDistinctivenessModelData.dblAveragePage4 = dataForDistinctivessModel.dblAveragePage4;
                                brandexStrategicDistinctivenessModelData.dblPage4Weight = dataForDistinctivessModel.dblPage4Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

                                brandexStrategicDistinctivenessModelData.dblAveragePage5 = dataForDistinctivessModel.dblAveragePage5;
                                brandexStrategicDistinctivenessModelData.dblPage5Weight = dataForDistinctivessModel.dblPage5Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage5Weighted = dataForDistinctivessModel.dblAveragePage5;

                                brandexStrategicDistinctivenessModelData.dblAveragePage6 = dataForDistinctivessModel.dblAveragePage6;
                                brandexStrategicDistinctivenessModelData.dblPage6Weight = dataForDistinctivessModel.dblPage6Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage6Weighted = dataForDistinctivessModel.dblAveragePage6Weighted;

                                brandexStrategicDistinctivenessModelData.dblAveragePage7 = dataForDistinctivessModel.dblAveragePage7;
                                brandexStrategicDistinctivenessModelData.dblPage7Weight = dataForDistinctivessModel.dblPage7Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage7Weighted = dataForDistinctivessModel.dblAveragePage7Weighted;

                                brandexStrategicDistinctivenessModelData.dblAveragePage8 = dataForDistinctivessModel.dblAveragePage8;
                                brandexStrategicDistinctivenessModelData.dblPage8Weight = dataForDistinctivessModel.dblPage8Weight;
                                brandexStrategicDistinctivenessModelData.dblAveragePage8Weighted = dataForDistinctivessModel.dblAveragePage8Weighted;

                                brandexStrategicDistinctivenessModelData.dblIndex = indexSum;
                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

                                // for chart distinctiveness

                                int fitToConceptStrategicWeightPage = 40;
                                int memorabilityStrategicWeightPage = 15;
                                int personalPreferenceStrategicWeightPage = 15;
                                int attributeEvaluationStrategicWeightPage = 30;

                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;
                                double averagePage3WeightedValueForChart = 0.0;
                                double averagePage4WeightedValueForChart = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValueForChart = (double)(((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptStrategicWeightPage) * scalingFactorForStrategicDistinctiveness)!;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = (double)(((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness)!;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValueForChart = (double)(((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness)!;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValueForChart = (double)(((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationStrategicWeightPage) * scalingFactorForStrategicDistinctiveness)!;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                double indexSumForChartDistinctiveness = averagePage1WeightedValueForChart +
                                                          averagePage2WeightedValueForChart +
                                                          averagePage3WeightedValueForChart +
                                                          averagePage4WeightedValueForChart;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName!;
                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForDistinctivenessChart = averagePage1WeightedValueForChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForDistinctivenessChart = averagePage2WeightedValueForChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForDistinctivenessChart = averagePage3WeightedValueForChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForDistinctivenessChart = averagePage4WeightedValueForChart;

                                brandexStrategicDistinctivenessModelData.dblIndexForMarketingChart = indexSumForChartDistinctiveness;

                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

                                // for chart marketing

                                int fitToConceptDistinctivenessWeightPage = 10;
                                int memorabilityDistinctivenssWeightPage = 30;
                                int personalPreferenceDistinctivenessWeightPage = 40;
                                int attributeEvaluationDistinctivenessWeightPage = 20;

                                double scalingFactorForMarketingChart = 1.00327;

                                double averagePage1WeightedValueForMarketingChart = 0.0;
                                double averagePage2WeightedValueForMarketingChart = 0.0;
                                double averagePage3WeightedValueForMarketingChart = 0.0;
                                double averagePage4WeightedValueForMarketingChart = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValueForMarketingChart = (double)(((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptDistinctivenessWeightPage) * scalingFactorForMarketingChart)!;
                                }
                                else
                                {
                                    averagePage1WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForMarketingChart = (double)(((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityDistinctivenssWeightPage) * scalingFactorForMarketingChart)!;
                                }
                                else
                                {
                                    averagePage2WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValueForMarketingChart = (double)(((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceDistinctivenessWeightPage) * scalingFactorForMarketingChart)!;
                                }
                                else
                                {
                                    averagePage3WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValueForMarketingChart = (double)(((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationDistinctivenessWeightPage) * scalingFactorForMarketingChart)!;
                                }
                                else
                                {
                                    averagePage4WeightedValueForMarketingChart = 0;
                                }

                                double indexSumForChartMarketing = averagePage1WeightedValueForMarketingChart +
                                                          averagePage2WeightedValueForMarketingChart +
                                                          averagePage3WeightedValueForMarketingChart +
                                                          averagePage4WeightedValueForMarketingChart;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName!;
                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForMarketingChart = averagePage1WeightedValueForMarketingChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForMarketingChart = averagePage2WeightedValueForMarketingChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForMarketingChart = averagePage3WeightedValueForMarketingChart;
                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForMarketingChart = averagePage4WeightedValueForMarketingChart;

                                brandexStrategicDistinctivenessModelData.dblIndexForDistinctivenessChart = indexSumForChartMarketing;
                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

                                brandexStrategicDistinctivenessesShortData.Add(brandexStrategicDistinctivenessModelData);
                            }

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexStrategicDistinctiveness{brandexStrategicData.Count()}.pptx";
                            dll.BrandexStrategicDistinctivenessMethod(CreateTargetPath(sourcePath, brandexStrategicData.First().ProjectTemplateType!), brandexStrategicDistinctivenessesShortData);

                            break;

                        case "Medical Terms":
                            var medicalTermsData = (await repository.GetMedicalTerms()).ToList();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholderRow.pptx";
                            dll.MedicalTermsMethod(CreateTargetPath(sourcePath, medicalTermsData.First().ProjectTemplateType!), medicalTermsData);
                            break;

                        case "Non-Medical Terms":
                            var nonMedicalTermsData = (await repository.GetNonMedicalTerms()).ToList();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholderRow.pptx";
                            dll.NonMedicalTermsMethod(CreateTargetPath(sourcePath, nonMedicalTermsData.First().ProjectTemplateType!), nonMedicalTermsData);
                            break;

                        case "01 Untrue":
                            var untrue01Data = (await repository.GetUntrue()).ToList();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative{untrue01Data.Count}.pptx";
                            dll.Untrue01Method(CreateTargetPath(sourcePath, untrue01Data.First().ProjectTemplateType), untrue01Data);
                            break;

                        case "02 Misleading":
                            var misleading02Data = (await repository.GetMisleading()).ToList();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative{misleading02Data.Count}.pptx";
                            dll.Misleading02Method(CreateTargetPath(sourcePath, misleading02Data.First().ProjectTemplateType!), misleading02Data);
                            break;

                        case "03 Exagg":
                            var exagg03Data = (await repository.Get03Exagg()).ToList();
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative{exagg03Data.Count}.pptx";
                            dll.Exagg03Method(CreateTargetPath(sourcePath, exagg03Data.First().ProjectTemplateType!), exagg03Data);
                            break;
                    }

                    taskLog.CompletedOn = DateTime.UtcNow;
                    _taskLogging.MarkTaskAsCompleted(request.TaskId.ToString(), (DateTime)taskLog.CompletedOn, connectionString!, "Done");

                    _tracker.SetStatus(request.TaskId, "Done");

                    _taskLogging.SetTaskStatusState(request.TaskId, "Done", connectionString!, user);

                }
                catch (Exception ex)
                {
                    _taskLogging.SetTaskStatusState(request.TaskId, "Fail", connectionString!, user);
                    _tracker.SetStatus(request.TaskId, $"Error: {ex.Message}");
                }
            });

            /* --- sample code -- */
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
