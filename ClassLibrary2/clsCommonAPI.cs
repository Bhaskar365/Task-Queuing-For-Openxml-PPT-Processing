////using DocumentFormat.OpenXml.Packaging;
////using DocumentFormat.OpenXml.Presentation;
////using Newtonsoft.Json;
////using OpenXmlDll;
////using prjData;
////using System;
////using System.Collections.Generic;
////using System.Data;
////using System.IO;
////using System.Linq;
////using System.Net.Http;
////using System.Text.RegularExpressions;
////using System.Threading.Tasks;
////using System.Xml.Linq;
////using static OpenXmlDLLDotnetFramework.DLLTemplate;

////namespace OpenXmlDLLDotnetFramework
////{
////    public class APIWrapper
////    {
////        string project = "";
////        string template = "";
////        string templateType = "";
////        string finalTemplate = "";
////        string group = "";
////        string breakdown = "";
////        string HistoricalMeanType = "";
////        string HistoricalMeanDescription = "";
////        //#pragma warning disable CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used
////        string genPath = "";
////        //#pragma warning restore CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used

////        public APIWrapper() { }

////        public APIWrapper(string project,
////                            string template,
////                            string group,
////                            string breakdown,
////                            string HistoricalMeanType,
////                            string HistoricalMeanDescription, string finalTemplate)
////        {
////            this.project = project;
////            this.template = template;
////            this.group = group;
////            this.breakdown = breakdown;
////            this.HistoricalMeanDescription = HistoricalMeanDescription;
////            this.HistoricalMeanType = HistoricalMeanType;
////            this.finalTemplate = finalTemplate;
////        }

////        public async Task<string> CallAPI(string project,
////                                        string template,
////                                        string group,
////                                        string breakdown,
////                                        string HistoricalMeanType,
////                                        string HistoricalMeanDescription)
////        {
////            using (HttpClient http = new HttpClient())
////            {
////                const string URL = "https://tools.brandinstitute.com/wsXlCharts/wsExcelCharts.asmx/GetChartBreakDownData";

////                //this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");



////                var payload = new[]
////                {
////                    new KeyValuePair<string,string>("token","2BF27A11-E318-447A-98FD-70AFE3871AA9"),
////                    new KeyValuePair<string,string>("project",project),
////                    new KeyValuePair<string,string>("template",template),
////                   //new KeyValuePair<string,string>("template",this.templateType),
////                    new KeyValuePair<string,string>("group",group),
////                    new KeyValuePair<string,string>("breakdown",breakdown),
////                    new KeyValuePair<string,string>("historicalMeanType",HistoricalMeanType),
////                    new KeyValuePair<string,string>("historicalMeanDescription",HistoricalMeanDescription)
////                };

////                var content = new FormUrlEncodedContent(payload);

////                try
////                {
////                    var response = await http.PostAsync(URL, content);

////                    if (response.IsSuccessStatusCode)
////                    {
////                        return await response.Content.ReadAsStringAsync();
////                    }
////                    else
////                    {
////                        return $"Error: {response.StatusCode}, {response.ReasonPhrase}";
////                    }
////                }
////                catch (Exception ex)
////                {
////                    return ex.Message;
////                }
////            }
////        }

////        public string CreateTargetPath(string myTemplate)
////        {
////            string path = $"C:\\excelfiles\\{this.project}";
////            if (!Directory.Exists(path))
////            {
////                Directory.CreateDirectory(path);
////            }
////            path = $"C:\\excelfiles\\{this.project}\\{this.template}_{this.breakdown}.pptx";
////            File.Copy(myTemplate, path, true);
////            return path;
////        }

////        public string getChartPathToCopyToFinal(string myTemplate, string project, string breakdown)
////        {
////            string path = $"C:\\excelfiles\\{project}";
////            if (!Directory.Exists(path))
////            {
////                Directory.CreateDirectory(path);
////            }
////            path = $"C:\\excelfiles\\{project}\\{myTemplate}_{breakdown}.pptx";
////            return path;
////        }

////        public async Task OpenXMLParallelProcess(string project,
////                                            List<string> templates,
////                                            List<string> breakdowns,
////                                            string HistoricalMeanType,
////                                            string HistoricalMeanDescription, string finalTemplateName)
////        {

////            //copy the template folder to the local system

////            //copyTemplates();


////            List<Task> taskArr = new List<Task>();

////            foreach (var breakdown in breakdowns)
////            {
////                foreach (var template in templates)
////                {

////                    APIWrapper wrapper = new APIWrapper(project, template, template, breakdown, HistoricalMeanType, HistoricalMeanDescription, finalTemplateName);
////                    taskArr.Add(Task.Run(() => wrapper.Process()));
////                }
////            }
////            await Task.WhenAll(taskArr);



////            //adding the template to the final



////        }



////        public Task<clsFinalPageNumberRange> getPPTFinalPageSettings(string strFinalReport, string year, string chartName)
////        {

////            return Task.Run(() =>
////            {

////                System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPPTFinalSettings] " + "'" + strFinalReport + "'," + "'" + year + "'");


////                List<clsFinalPageModel> lstFinalPageModel = clsData.MRData.ConvertDataTableToListGeneric<clsFinalPageModel>(dt);


////                List<clsFinalPageModel> lstFinal = lstFinalPageModel.Where(n => n.strPageGroupName == chartName).ToList();

////                clsFinalPageNumberRange objPageNum = new clsFinalPageNumberRange();

////                foreach (clsFinalPageModel obj in lstFinal)
////                {
////                    objPageNum.firstPage = obj.intPPTSlideIndexFirst;
////                    objPageNum.lastPage = obj.intPPTSlideIndexLast;

////                }

////                //  Task<clsFinalPageNumberRange> obj = new Task<clsFinalPageNumberRange>();


////                return objPageNum;
////            });


////        }



////        public async Task ProcessMultipleNoParam()
////        {
////            //APIWrapper wrapper = new APIWrapper("RACKEM", "Fit to Concept", "Fit to Concept", "OVERALL", "Fit to Concept", "2022-HCPS");
////            //APIWrapper wrapper3 = new APIWrapper("DESIGNATION_REDO", "Potential For Error - Bar", "Potential For Error - Bar", "OVERALL", "", "");
////            //APIWrapper wrapper4 = new APIWrapper("EDAT_2", "Personal Preferences", "Personal Preferences", "OVERALL", "", "");
////            //APIWrapper wrapper5 = new APIWrapper("SOLITAIRE", "Likeability", "01 Untrue", "OVERALL", "", "");
////            //APIWrapper wrapper6 = new APIWrapper("MEADOW", "Memorability", "Memorability", "OVERALL", "", "");
////            //APIWrapper wrapper7 = new APIWrapper("TRIBUTE_C2", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
////            //APIWrapper wrapper8 = new APIWrapper("HABITABLE", "Overall Impressions", "Overall Impressions", "Canada Overall", "", "");
////            //APIWrapper wrapper9 = new APIWrapper("ZIPLOCK", "Suffix", "Modifier Confusion", "U.S. Overall", "", "");
////            //APIWrapper wrapper10 = new APIWrapper("ZIPLOCK", "PromotionalReview", "Modifier Confusion", "U.S. Overall", "", "");
////            //APIWrapper wrapper11 = new APIWrapper("MEADOW", "Ease Of Pronounciation", "Att 1", "OVERALL", "", "");
////            //APIWrapper wrapper12 = new APIWrapper("MEADOW", "Ease Of Spelling", "Att 1", "OVERALL", "", "");
////            //APIWrapper wrapper13 = new APIWrapper("DOMINOS", "03 exagg", "03 exagg", "OVERALL", "", "");
////            //APIWrapper wrapper14 = new APIWrapper("DORAEMON", "Innovation", "Innovation", "OVERALL", "", "");
////            //APIWrapper wrapper15 = new APIWrapper("ATEMPORAL", "Modifier", "Modifier Meaning (Aided)", "Canada Overall", "", "");
////            //APIWrapper wrapper16 = new APIWrapper("QUEENSLAND", "Written Understanding - Bar", "Written Understanding - Bar", "Canada and Europe Overall", "", "");

////            //APIWrapper wrapper17 = new APIWrapper("RACKEM", "Attribute 1", "Attribute 1", "OVERALL", "", "");
////            //APIWrapper wrapper18 = new APIWrapper("RACKEM", "Attribute 2", "Attribute 2", "OVERALL", "", "");
////            //APIWrapper wrapper19 = new APIWrapper("RACKEM", "Attribute 3", "Attribute 3", "OVERALL", "", "");
////            //APIWrapper wrapper20 = new APIWrapper("RACKEM", "01 Untrue", "01 Untrue", "OVERALL", "", "");
////            //APIWrapper wrapper21 = new APIWrapper("RACKEM", "02 Misleading", "02 Misleading", "OVERALL", "", "");
////            //APIWrapper wrapper22 = new APIWrapper("RACKEM", "03 Exagg", "03 Exagg", "OVERALL", "", "");
////            //APIWrapper wrapper24 = new APIWrapper("RACKEM", "Memorability", "Memorability", "OVERALL", "", "");
////            //APIWrapper wrapper25 = new APIWrapper("RACKEM", "Overall Impressions", "Overall Impressions", "OVERALL", "", "");
////            //APIWrapper wrapper26 = new APIWrapper("RACKEM", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
////            //APIWrapper wrapper27 = new APIWrapper("RACKEM", "Written Understanding - Bar", "Written Understanding - Bar", "Overall", "", "");
////            ////APIWrapper wrapper28 = new APIWrapper("RACKEM", "Potential For Error - Bar", "Potential For Error - Bar", "Overall", "", "");
////            //APIWrapper wrapper29 = new APIWrapper("VGT_BND", "Preference Ranking", "Preference Ranking", "Overall", "", "");
////            //APIWrapper wrapper30 = new APIWrapper("RACKEM", "Attribute evaluation Aggregate", "Attribute evaluation Aggregate", "Overall", "", "");
////            //APIWrapper wrapper31 = new APIWrapper("VAYNSMR_HEME", "Initial Recall", "Initial Recall", "Canada and EU Medical Professionals", "", "");
////            //APIWrapper wrapper32 = new APIWrapper("ALL_FAME", "Chemical Structure Appropriateness", "", "Overall", "", "");

////            //APIWrapper wrapper33 = new APIWrapper("COPEMVIBO", "Likeability Rationale", "Likeability Rationale", "Overall", "", "");


////            Task[] taskArr = new Task[]
////            {
////                //Task.Run(() => wrapper.Process()),
////                //Task.Run(() => wrapper3.Process()),
////                //Task.Run(() => wrapper4.Process()),
////                //Task.Run(() => wrapper5.Process()),
////                //Task.Run(() => wrapper6.Process()),
////                //Task.Run(() => wrapper7.Process()),
////                //Task.Run(() => wrapper8.Process()),
////                //Task.Run(() => wrapper9.Process()),
////                //Task.Run(() => wrapper10.Process()),
////                //Task.Run(() => wrapper11.Process()),
////                //Task.Run(() => wrapper12.Process()),
////                //Task.Run(() => wrapper13.Process()),
////                //Task.Run(() => wrapper14.Process()),
////                //Task.Run(() => wrapper15.Process()),
////                //Task.Run(() => wrapper16.Process()),

////                // Task.Run(() => wrapper17.Process()),
////                // Task.Run(() => wrapper18.Process()),
////                // Task.Run(() => wrapper19.Process()),
////                // Task.Run(() => wrapper20.Process()),
////                // Task.Run(() => wrapper21.Process()),
////                // Task.Run(() => wrapper22.Process()),
////                // Task.Run(() => wrapper24.Process()),
////                // Task.Run(() => wrapper25.Process()),
////                // Task.Run(() => wrapper26.Process()),
////                // Task.Run(() => wrapper27.Process()),
////                //Task.Run(() => wrapper28.Process()),

////                //Task.Run(() => wrapper29.Process()),
////                //Task.Run(() => wrapper30.Process()),
////                //Task.Run(() => wrapper31.Process()),
////                //Task.Run(() => wrapper32.Process()),
////                //Task.Run(() => wrapper33.Process())
////            };

////            await Task.WhenAll(taskArr);
////        }

////        public async Task Process()
////        {

////            //get the pageGroup

////            this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");


////            string sourcePath = "";
////            var apiCall = await CallAPI(project, template, group, breakdown, HistoricalMeanType, HistoricalMeanDescription);

////            DLLClass dLLClass = new DLLClass();
////            XDocument xDoc = XDocument.Parse(apiCall);

////            var data = xDoc.Root.Value;

////            var jsonSettings = new JsonSerializerSettings
////            {
////                NullValueHandling = NullValueHandling.Ignore,
////                MissingMemberHandling = MissingMemberHandling.Ignore
////            };


////            if (data.Length == 0)
////            {
////                Console.WriteLine("No data present in api");
////                return;
////            }

////            if (data == null || data.ToString() == "[]")
////            {
////                sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                notAvailable(CreateTargetPath(sourcePath), template, breakdown);
////            }


////            if (data.Length != 0)
////            {
////                switch (this.templateType)
////                {
////                    case "Fit to Concept":


////                        List<DLLTemplate.FitToConceptModel> fitToConceptData = new List<DLLTemplate.FitToConceptModel>();

////                        fitToConceptData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


////                        if (fitToConceptData.Count == 0)
////                        {
////                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            //CreateTargetPath(sourcePath);

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept" + fitToConceptData.Count + ".pptx";
////                            dLLClass.FitToConceptMethod(CreateTargetPath(sourcePath), fitToConceptData, HistoricalMeanType, HistoricalMeanDescription);
////                        }
////                        break;

////                    case "Attribute 1":
////                    case "Attribute 2":
////                    case "Attribute 3":
////                    case "Att 1":
////                    case "Att 2":
////                    case "Att 3":
////                    case "Attribute Evaluation":



////                        List<DLLTemplate.FitToConceptModel> Att1Data = new List<DLLTemplate.FitToConceptModel>();

////                        Att1Data = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

////                        if (Att1Data.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                            //update the notavailable 


////                        }
////                        else
////                        {


////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation" + Att1Data.Count + ".pptx";
////                            dLLClass.AttributeMethod(CreateTargetPath(sourcePath), Att1Data, group, HistoricalMeanType, HistoricalMeanDescription, project);


////                        }
////                        break;

////                    case "Attribute evaluation Aggregate":
////                    case "Attribute Evaluation Aggregate":

////                        List<DLLTemplate.FitToConceptModel> attEvalData = new List<DLLTemplate.FitToConceptModel>();

////                        attEvalData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


////                        if (attEvalData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluationAggregate" + attEvalData.Count + ".pptx";
////                            dLLClass.AttributeMethodForAttributeEvalAggreg(CreateTargetPath(sourcePath), attEvalData);
////                        }
////                        break;



////                    case "BRANDEX SUFFIX STRATEGIC":
////                        List<BrandexSuffixStrategicModel> brandexSuffixStrategicData = new List<BrandexSuffixStrategicModel>();
////                        brandexSuffixStrategicData = JsonConvert.DeserializeObject<List<BrandexSuffixStrategicModel>>(data);

////                        if (brandexSuffixStrategicData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            List<BrandexSuffixStrategicShortModel> brandexSuffixStrategicShortModel = new List<BrandexSuffixStrategicShortModel>();

////                            double dblAverage1Max = 0.0;
////                            double dblAverage2Max = 0.0;

////                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
////                            {
////                                var dblAverage1MaxValue = brandexSuffixStrategicData[i].dblAveragePage1;
////                                if (dblAverage1MaxValue == 0)
////                                {
////                                    dblAverage1Max = 0.0;
////                                }
////                                if (dblAverage1MaxValue > dblAverage1Max)
////                                {
////                                    dblAverage1Max = dblAverage1MaxValue;
////                                }

////                                var dblAverage2MaxValue = brandexSuffixStrategicData[i].dblAveragePage2;
////                                if (dblAverage2MaxValue == 0)
////                                {
////                                    dblAverage2Max = 0;
////                                }
////                                if (dblAverage2MaxValue > dblAverage2Max)
////                                {
////                                    dblAverage2Max = dblAverage2MaxValue;
////                                }
////                            }

////                            // scaling factor 
////                            double scalingFactorForStrategicDistinctiveness = 1.0;

////                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
////                            {
////                                // for the table
////                                BrandexSuffixStrategicShortModel brandexSuffixStrategicModelData = new BrandexSuffixStrategicShortModel();

////                                var dataForBrandexSuffix = brandexSuffixStrategicData[i];

////                                double averagePage1WeightedValue = 0.0;
////                                double averagePage2WeightedValue = 0.0;

////                                if (dblAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValue =
////                                        (dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * dataForBrandexSuffix.dblPage1Weight;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValue = 0;
////                                }

////                                if (dblAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValue =
////                                       (dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * dataForBrandexSuffix.dblPage2Weight;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValue = 0;
////                                }

////                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue) * scalingFactorForStrategicDistinctiveness;

////                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;

////                                brandexSuffixStrategicModelData.dblAveragePage1 = dataForBrandexSuffix.dblAveragePage1;
////                                brandexSuffixStrategicModelData.dblPage1Weight = dataForBrandexSuffix.dblPage1Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

////                                brandexSuffixStrategicModelData.dblAveragePage2 = dataForBrandexSuffix.dblAveragePage2;
////                                brandexSuffixStrategicModelData.dblPage2Weight = dataForBrandexSuffix.dblPage2Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

////                                brandexSuffixStrategicModelData.dblAveragePage3 = dataForBrandexSuffix.dblAveragePage3;
////                                brandexSuffixStrategicModelData.dblPage3Weight = dataForBrandexSuffix.dblPage3Weight;

////                                brandexSuffixStrategicModelData.dblAveragePage4 = dataForBrandexSuffix.dblAveragePage4;
////                                brandexSuffixStrategicModelData.dblPage4Weight = dataForBrandexSuffix.dblPage4Weight;

////                                brandexSuffixStrategicModelData.dblAveragePage5 = dataForBrandexSuffix.dblAveragePage5;
////                                brandexSuffixStrategicModelData.dblPage5Weight = dataForBrandexSuffix.dblPage5Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage5Weighted = dataForBrandexSuffix.dblAveragePage5;

////                                brandexSuffixStrategicModelData.dblAveragePage6 = dataForBrandexSuffix.dblAveragePage6;
////                                brandexSuffixStrategicModelData.dblPage6Weight = dataForBrandexSuffix.dblPage6Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage6Weighted = dataForBrandexSuffix.dblAveragePage6Weighted;

////                                brandexSuffixStrategicModelData.dblAveragePage7 = dataForBrandexSuffix.dblAveragePage7;
////                                brandexSuffixStrategicModelData.dblPage7Weight = dataForBrandexSuffix.dblPage7Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage7Weighted = dataForBrandexSuffix.dblAveragePage7Weighted;

////                                brandexSuffixStrategicModelData.dblAveragePage8 = dataForBrandexSuffix.dblAveragePage8;
////                                brandexSuffixStrategicModelData.dblPage8Weight = dataForBrandexSuffix.dblPage8Weight;
////                                brandexSuffixStrategicModelData.dblAveragePage8Weighted = dataForBrandexSuffix.dblAveragePage8Weighted;

////                                brandexSuffixStrategicModelData.dblIndex = indexSum;
////                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
////                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
////                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
////                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
////                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

////                                // for chart
////                                int memorabilityStrategicWeightPage = 50;
////                                int personalPreferenceStrategicWeightPage = 50;

////                                double averagePage1WeightedValueForChart = 0.0;
////                                double averagePage2WeightedValueForChart = 0.0;

////                                if (dblAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForChart = 0;
////                                }

////                                if (dblAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForChart = 0;
////                                }

////                                double indexSumForChart = averagePage1WeightedValueForChart +
////                                                          averagePage2WeightedValueForChart;

////                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;
////                                brandexSuffixStrategicModelData.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;
////                                brandexSuffixStrategicModelData.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

////                                brandexSuffixStrategicModelData.dblIndexForChart = indexSumForChart;

////                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
////                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
////                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
////                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
////                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

////                                brandexSuffixStrategicShortModel.Add(brandexSuffixStrategicModelData);
////                            }

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSuffixStrategicTemplateModified" + brandexSuffixStrategicShortModel.Count + ".pptx";
////                            dLLClass.BrandexSuffixStrategicMethod(CreateTargetPath(sourcePath), brandexSuffixStrategicShortModel);
////                        }
////                        break;





////                    //case "Initial Recall":
////                    //    List<DLLTemplate.FitToConceptNewModel> initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();
////                    //    initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data);
////                    //    if (initialRecallData.Count == 0)
////                    //    {
////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                    //    }
////                    //    else
////                    //    {
////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + initialRecallData.Count + ".pptx";
////                    //        dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), initialRecallData, this.HistoricalMeanType, this.HistoricalMeanDescription);
////                    //    }
////                    //    break;





////                    case "Brandex Strategic Distinctiveness":

////                        List<BrandexStrategicDistinctivenessModel> brandexStrategicDistinctivenessesData = new List<BrandexStrategicDistinctivenessModel>();
////                        brandexStrategicDistinctivenessesData = JsonConvert.DeserializeObject<List<BrandexStrategicDistinctivenessModel>>(data);

////                        if (brandexStrategicDistinctivenessesData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            List<BrandexStrategicDistinctivenessShortModel> brandexStrategicDistinctivenessesShortData = new List<BrandexStrategicDistinctivenessShortModel>();

////                            double dblAverage1Max = 0.0;
////                            double dblAverage2Max = 0.0;
////                            double dblAverage3Max = 0.0;
////                            double dblAverage4Max = 0.0;

////                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
////                            {
////                                var dblAverage1MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage1;
////                                if (dblAverage1MaxValue == 0)
////                                {
////                                    dblAverage1Max = 0.0;
////                                }
////                                if (dblAverage1MaxValue > dblAverage1Max)
////                                {
////                                    dblAverage1Max = dblAverage1MaxValue;
////                                }

////                                var dblAverage2MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage2;
////                                if (dblAverage2MaxValue == 0)
////                                {
////                                    dblAverage2Max = 0;
////                                }
////                                if (dblAverage2MaxValue > dblAverage2Max)
////                                {
////                                    dblAverage2Max = dblAverage2MaxValue;
////                                }

////                                var dblAverage3MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage3;
////                                if (dblAverage3MaxValue == 0)
////                                {
////                                    dblAverage3Max = 0;
////                                }
////                                if (dblAverage3MaxValue > dblAverage3Max)
////                                {
////                                    dblAverage3Max = dblAverage3MaxValue;
////                                }

////                                var dblAverage4MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage4;
////                                if (dblAverage4MaxValue == 0)
////                                {
////                                    dblAverage4Max = 0;
////                                }
////                                if (dblAverage4MaxValue > dblAverage4Max)
////                                {
////                                    dblAverage4Max = dblAverage4MaxValue;
////                                }
////                            }

////                            // scaling factor 
////                            double scalingFactorForStrategicDistinctiveness = 1.00777;

////                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
////                            {
////                                // for the table
////                                BrandexStrategicDistinctivenessShortModel brandexStrategicDistinctivenessModelData = new BrandexStrategicDistinctivenessShortModel();

////                                var dataForDistinctivessModel = brandexStrategicDistinctivenessesData[i];

////                                double averagePage1WeightedValue = 0.0;
////                                double averagePage2WeightedValue = 0.0;
////                                double averagePage3WeightedValue = 0.0;
////                                double averagePage4WeightedValue = 0.0;

////                                if (dblAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValue =
////                                        (dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * dataForDistinctivessModel.dblPage1Weight;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValue = 0;
////                                }

////                                if (dblAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValue =
////                                       (dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * dataForDistinctivessModel.dblPage2Weight;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValue = 0;
////                                }

////                                if (dblAverage3Max > 0)
////                                {
////                                    averagePage3WeightedValue =
////                                        (dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * dataForDistinctivessModel.dblPage3Weight;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValue = 0;
////                                }

////                                if (dblAverage4Max > 0)
////                                {
////                                    averagePage4WeightedValue =
////                                        (dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * dataForDistinctivessModel.dblPage4Weight;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValue = 0;
////                                }

////                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
////                                             averagePage4WeightedValue) * scalingFactorForStrategicDistinctiveness;

////                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage1 = dataForDistinctivessModel.dblAveragePage1;
////                                brandexStrategicDistinctivenessModelData.dblPage1Weight = dataForDistinctivessModel.dblPage1Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage2 = dataForDistinctivessModel.dblAveragePage2;
////                                brandexStrategicDistinctivenessModelData.dblPage2Weight = dataForDistinctivessModel.dblPage2Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage3 = dataForDistinctivessModel.dblAveragePage3;
////                                brandexStrategicDistinctivenessModelData.dblPage3Weight = dataForDistinctivessModel.dblPage3Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage4 = dataForDistinctivessModel.dblAveragePage4;
////                                brandexStrategicDistinctivenessModelData.dblPage4Weight = dataForDistinctivessModel.dblPage4Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage5 = dataForDistinctivessModel.dblAveragePage5;
////                                brandexStrategicDistinctivenessModelData.dblPage5Weight = dataForDistinctivessModel.dblPage5Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage5Weighted = dataForDistinctivessModel.dblAveragePage5;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage6 = dataForDistinctivessModel.dblAveragePage6;
////                                brandexStrategicDistinctivenessModelData.dblPage6Weight = dataForDistinctivessModel.dblPage6Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage6Weighted = dataForDistinctivessModel.dblAveragePage6Weighted;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage7 = dataForDistinctivessModel.dblAveragePage7;
////                                brandexStrategicDistinctivenessModelData.dblPage7Weight = dataForDistinctivessModel.dblPage7Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage7Weighted = dataForDistinctivessModel.dblAveragePage7Weighted;

////                                brandexStrategicDistinctivenessModelData.dblAveragePage8 = dataForDistinctivessModel.dblAveragePage8;
////                                brandexStrategicDistinctivenessModelData.dblPage8Weight = dataForDistinctivessModel.dblPage8Weight;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage8Weighted = dataForDistinctivessModel.dblAveragePage8Weighted;

////                                brandexStrategicDistinctivenessModelData.dblIndex = indexSum;
////                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
////                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
////                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
////                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
////                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

////                                // for chart distinctiveness

////                                int fitToConceptStrategicWeightPage = 40;
////                                int memorabilityStrategicWeightPage = 15;
////                                int personalPreferenceStrategicWeightPage = 15;
////                                int attributeEvaluationStrategicWeightPage = 30;

////                                double averagePage1WeightedValueForChart = 0.0;
////                                double averagePage2WeightedValueForChart = 0.0;
////                                double averagePage3WeightedValueForChart = 0.0;
////                                double averagePage4WeightedValueForChart = 0.0;

////                                if (dblAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForChart = 0;
////                                }

////                                if (dblAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForChart = 0;
////                                }

////                                if (dblAverage3Max > 0)
////                                {
////                                    averagePage3WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValueForChart = 0;
////                                }

////                                if (dblAverage4Max > 0)
////                                {
////                                    averagePage4WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValueForChart = 0;
////                                }

////                                double indexSumForChartDistinctiveness = averagePage1WeightedValueForChart +
////                                                          averagePage2WeightedValueForChart +
////                                                          averagePage3WeightedValueForChart +
////                                                          averagePage4WeightedValueForChart;

////                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForDistinctivenessChart = averagePage1WeightedValueForChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForDistinctivenessChart = averagePage2WeightedValueForChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForDistinctivenessChart = averagePage3WeightedValueForChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForDistinctivenessChart = averagePage4WeightedValueForChart;

////                                brandexStrategicDistinctivenessModelData.dblIndexForMarketingChart = indexSumForChartDistinctiveness;

////                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
////                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
////                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
////                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
////                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

////                                // for chart marketing

////                                int fitToConceptDistinctivenessWeightPage = 10;
////                                int memorabilityDistinctivenssWeightPage = 30;
////                                int personalPreferenceDistinctivenessWeightPage = 40;
////                                int attributeEvaluationDistinctivenessWeightPage = 20;

////                                double scalingFactorForMarketingChart = 1.00327;

////                                double averagePage1WeightedValueForMarketingChart = 0.0;
////                                double averagePage2WeightedValueForMarketingChart = 0.0;
////                                double averagePage3WeightedValueForMarketingChart = 0.0;
////                                double averagePage4WeightedValueForMarketingChart = 0.0;

////                                if (dblAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptDistinctivenessWeightPage) * scalingFactorForMarketingChart;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForMarketingChart = 0;
////                                }

////                                if (dblAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityDistinctivenssWeightPage) * scalingFactorForMarketingChart;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForMarketingChart = 0;
////                                }

////                                if (dblAverage3Max > 0)
////                                {
////                                    averagePage3WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceDistinctivenessWeightPage) * scalingFactorForMarketingChart;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValueForMarketingChart = 0;
////                                }

////                                if (dblAverage4Max > 0)
////                                {
////                                    averagePage4WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationDistinctivenessWeightPage) * scalingFactorForMarketingChart;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValueForMarketingChart = 0;
////                                }

////                                double indexSumForChartMarketing = averagePage1WeightedValueForMarketingChart +
////                                                          averagePage2WeightedValueForMarketingChart +
////                                                          averagePage3WeightedValueForMarketingChart +
////                                                          averagePage4WeightedValueForMarketingChart;

////                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForMarketingChart = averagePage1WeightedValueForMarketingChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForMarketingChart = averagePage2WeightedValueForMarketingChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForMarketingChart = averagePage3WeightedValueForMarketingChart;
////                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForMarketingChart = averagePage4WeightedValueForMarketingChart;

////                                brandexStrategicDistinctivenessModelData.dblIndexForDistinctivenessChart = indexSumForChartMarketing;
////                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
////                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
////                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
////                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
////                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

////                                brandexStrategicDistinctivenessesShortData.Add(brandexStrategicDistinctivenessModelData);
////                            }

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexStrategicDistinctiveness" + brandexStrategicDistinctivenessesData.Count + ".pptx";
////                            dLLClass.BrandexStrategicDistinctivenessMethod(CreateTargetPath(sourcePath), brandexStrategicDistinctivenessesShortData);
////                        }




////                        break;


////                    //case "Brandex Safety":
////                    //    List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
////                    //    brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

////                    //    List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

////                    //    if (brandexSafetyData.Count == 0)
////                    //    {
////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                    //    }
////                    //    else
////                    //    {
////                    //        double average1Max = 0.0000;
////                    //        double average2Max = 0.0000;
////                    //        double average3Max = 0.0;
////                    //        double average4Max = 0.0000;
////                    //        double average5Max = 0.0000;

////                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
////                    //        {
////                    //            var dblAverage1max = brandexSafetyData[i].dblAveragePage1;
////                    //            if (dblAverage1max == 0.0)
////                    //            {
////                    //                average1Max = 0;
////                    //            }
////                    //            if (average1Max < dblAverage1max)
////                    //            {
////                    //                average1Max = dblAverage1max;
////                    //            }

////                    //            var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
////                    //            if (dblAverage2max == 0.0)
////                    //            {
////                    //                average2Max = 0;
////                    //            }
////                    //            if (average2Max < dblAverage2max)
////                    //            {
////                    //                average2Max = dblAverage2max;
////                    //            }

////                    //            var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
////                    //            if (dblAverage3max == 0.0)
////                    //            {
////                    //                average3Max = 0;
////                    //            }
////                    //            if (average3Max < dblAverage3max)
////                    //            {
////                    //                average3Max = dblAverage3max;
////                    //            }

////                    //            var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
////                    //            if (dblAverage4max == 0.0)
////                    //            {
////                    //                average4Max = 0;
////                    //            }
////                    //            if (average4Max < dblAverage4max)
////                    //            {
////                    //                average4Max = dblAverage4max;
////                    //            }

////                    //            var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
////                    //            if (dblAverage5max == 0.0)
////                    //            {
////                    //                average5Max = 0;
////                    //            }
////                    //            if (average5Max < dblAverage5max)
////                    //            {
////                    //                average5Max = dblAverage5max;
////                    //            }
////                    //        }

////                    //        double scalingFactor = 0.750610000;

////                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
////                    //        {
////                    //            //for the table
////                    //            BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

////                    //            var dataEl = brandexSafetyData[i];

////                    //            double averagePage1WeightedValue = 0.0;
////                    //            double averagePage2WeightedValue = 0.0;
////                    //            double averagePage3WeightedValue = 0.0;
////                    //            double averagePage4WeightedValue = 0.0;
////                    //            double averagePage5WeightedValue = 0.0;

////                    //            if (average1Max > 0)
////                    //            {
////                    //                averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage1WeightedValue = 0;
////                    //            }

////                    //            if (average2Max > 0)
////                    //            {
////                    //                averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage2WeightedValue = 0;
////                    //            }

////                    //            if (average3Max > 0)
////                    //            {
////                    //                averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage3WeightedValue = 0;
////                    //            }

////                    //            if (average4Max > 0)
////                    //            {
////                    //                averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage4WeightedValue = 0;
////                    //            }

////                    //            if (average5Max > 0)
////                    //            {
////                    //                averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage5WeightedValue = 0;
////                    //            }


////                    //            //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
////                    //            //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
////                    //            //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
////                    //            // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
////                    //            //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

////                    //            double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
////                    //                              averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;

////                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                    //            brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
////                    //            brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
////                    //            brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

////                    //            brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
////                    //            brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
////                    //            brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

////                    //            brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
////                    //            brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
////                    //            brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

////                    //            brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
////                    //            brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
////                    //            brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

////                    //            brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
////                    //            brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
////                    //            brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

////                    //            brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
////                    //            brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
////                    //            brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

////                    //            brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
////                    //            brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
////                    //            brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

////                    //            brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
////                    //            brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
////                    //            brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

////                    //            brandexSafetyShortModel.dblIndex = indexSum;
////                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
////                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                    //            //for the chart - 

////                    //            double averagePage1WeightedValueForChart = 0.0;
////                    //            double averagePage2WeightedValueForChart = 0.0;
////                    //            double averagePage3WeightedValueForChart = 0.0;
////                    //            double averagePage4WeightedValueForChart = 0.0;
////                    //            double averagePage5WeightedValueForChart = 0.0;

////                    //            if (average1Max > 0)
////                    //            {
////                    //                averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage1WeightedValueForChart = 0;
////                    //            }

////                    //            if (average2Max > 0)
////                    //            {
////                    //                averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage2WeightedValueForChart = 0;
////                    //            }

////                    //            if (average3Max > 0)
////                    //            {
////                    //                averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage3WeightedValueForChart = 0;
////                    //            }

////                    //            if (average4Max > 0)
////                    //            {
////                    //                averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage4WeightedValueForChart = 0;
////                    //            }

////                    //            if (average5Max > 0)
////                    //            {
////                    //                averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
////                    //            }
////                    //            else
////                    //            {
////                    //                averagePage5WeightedValueForChart = 0;
////                    //            }

////                    //            //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
////                    //            //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
////                    //            //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
////                    //            //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
////                    //            //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

////                    //            double indexSumForChart = averagePage1WeightedValueForChart +
////                    //                                      averagePage2WeightedValueForChart +
////                    //                                      averagePage3WeightedValueForChart +
////                    //                                      averagePage4WeightedValueForChart +
////                    //                                      averagePage5WeightedValueForChart;

////                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                    //            brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

////                    //            brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

////                    //            brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

////                    //            brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

////                    //            brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
////                    //            ;
////                    //            brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
////                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
////                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                    //            brandexData.Add(brandexSafetyShortModel);
////                    //        }

////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
////                    //        dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
////                    //    }
////                    //    break;


////                    case "Brandex Safety":
////                        List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
////                        brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

////                        List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

////                        if (brandexSafetyData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            double average1Max = 0.0000;
////                            double average2Max = 0.0000;
////                            double average3Max = 0.0;
////                            double average4Max = 0.0000;
////                            double average5Max = 0.0000;

////                            for (int i = 0; i < brandexSafetyData.Count; i++)
////                            {
////                                var dblAverage1max = brandexSafetyData[i]?.dblAveragePage1;
////                                if (dblAverage1max == 0.0 || dblAverage1max == null)
////                                {
////                                    average1Max = 0;
////                                }
////                                if (average1Max < dblAverage1max)
////                                {
////                                    average1Max = (double)dblAverage1max;
////                                }

////                                var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
////                                if (dblAverage2max == 0.0)
////                                {
////                                    average2Max = 0;
////                                }
////                                if (average2Max < dblAverage2max)
////                                {
////                                    average2Max = (double)dblAverage2max;
////                                }

////                                var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
////                                if (dblAverage3max == 0.0)
////                                {
////                                    average3Max = 0;
////                                }
////                                if (average3Max < dblAverage3max)
////                                {
////                                    average3Max = (double)dblAverage3max;
////                                }

////                                var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
////                                if (dblAverage4max == 0.0)
////                                {
////                                    average4Max = 0;
////                                }
////                                if (average4Max < dblAverage4max)
////                                {
////                                    average4Max = (double)dblAverage4max;
////                                }

////                                var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
////                                if (dblAverage5max == 0.0)
////                                {
////                                    average5Max = 0;
////                                }
////                                if (average5Max < dblAverage5max)
////                                {
////                                    average5Max = (double)dblAverage5max;
////                                }
////                            }

////                            double scalingFactor = 0.750610000;

////                            for (int i = 0; i < brandexSafetyData.Count; i++)
////                            {
////                                //for the table
////                                BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

////                                var dataEl = brandexSafetyData[i];

////                                double averagePage1WeightedValue = 0.0;
////                                double averagePage2WeightedValue = 0.0;
////                                double averagePage3WeightedValue = 0.0;
////                                double averagePage4WeightedValue = 0.0;
////                                double averagePage5WeightedValue = 0.0;

////                                if (average1Max > 0)
////                                {
////                                    // averagePage1WeightedValue =(dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) ;

////                                    averagePage1WeightedValue = (double)((double)(dataEl.dblAveragePage1 / average1Max) * (double)(dataEl.dblPage1Weight));
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValue = 0;
////                                }

////                                if (average2Max > 0)
////                                {
////                                    // averagePage2WeightedValue = (double)(dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;

////                                    averagePage2WeightedValue = (double)((double)(dataEl.dblAveragePage2 / average2Max) * (double)(dataEl.dblPage2Weight));
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValue = 0;
////                                }

////                                if (average3Max > 0)
////                                {
////                                    averagePage3WeightedValue = (double)((double)(dataEl.dblAveragePage3 / average3Max) * (double)(dataEl.dblPage3Weight));
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValue = 0;
////                                }

////                                if (average4Max > 0)
////                                {
////                                    //averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;

////                                    averagePage4WeightedValue = (double)((double)(dataEl.dblAveragePage4 / average4Max) * (double)(dataEl.dblPage4Weight));
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValue = 0;
////                                }

////                                if (average5Max > 0)
////                                {
////                                    // averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
////                                    averagePage5WeightedValue = (double)((double)(dataEl.dblAveragePage5 / average5Max) * (double)(dataEl.dblPage5Weight));


////                                }
////                                else
////                                {
////                                    averagePage5WeightedValue = 0;
////                                }


////                                //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
////                                //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
////                                //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
////                                // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
////                                //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

////                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
////                                                  averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;

////                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1 ?? 0.0;
////                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2 ?? 0.0;
////                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3 ?? 0.0;
////                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4 ?? 0.0;
////                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5 ?? 0.0;
////                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6 ?? 0.0;
////                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

////                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7 ?? 0.0;
////                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

////                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8 ?? 0.0;
////                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight ?? 0.0;
////                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

////                                brandexSafetyShortModel.dblIndex = indexSum;
////                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                                brandexSafetyShortModel.intRed = dataEl.intRed;
////                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                                //for the chart - 

////                                double averagePage1WeightedValueForChart = 0.0;
////                                double averagePage2WeightedValueForChart = 0.0;
////                                double averagePage3WeightedValueForChart = 0.0;
////                                double averagePage4WeightedValueForChart = 0.0;
////                                double averagePage5WeightedValueForChart = 0.0;

////                                if (average1Max > 0)
////                                {
////                                    averagePage1WeightedValueForChart = (double)((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForChart = 0;
////                                }

////                                if (average2Max > 0)
////                                {
////                                    averagePage2WeightedValueForChart = (double)((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForChart = 0;
////                                }

////                                if (average3Max > 0)
////                                {
////                                    averagePage3WeightedValueForChart = (double)((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValueForChart = 0;
////                                }

////                                if (average4Max > 0)
////                                {
////                                    averagePage4WeightedValueForChart = (double)((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValueForChart = 0;
////                                }

////                                if (average5Max > 0)
////                                {
////                                    averagePage5WeightedValueForChart = (double)((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage5WeightedValueForChart = 0;
////                                }

////                                //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
////                                //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
////                                //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
////                                //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
////                                //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

////                                double indexSumForChart = averagePage1WeightedValueForChart +
////                                                          averagePage2WeightedValueForChart +
////                                                          averagePage3WeightedValueForChart +
////                                                          averagePage4WeightedValueForChart +
////                                                          averagePage5WeightedValueForChart;

////                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                                brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
////                                ;
////                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
////                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                                brandexSafetyShortModel.intRed = dataEl.intRed;
////                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                                brandexData.Add(brandexSafetyShortModel);
////                            }

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
////                            dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
////                        }
////                        break;

////                    case "Potential For Error - Bar":
////                    case "Potential For Error":


////                        List<DLLTemplate.FitToConceptModel> potentialForError = new List<DLLTemplate.FitToConceptModel>();

////                        potentialForError = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


////                        if (potentialForError.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PotentialForError-Bar" + potentialForError.Count + ".pptx";
////                            dLLClass.PotentialForErrorBarMethod(CreateTargetPath(sourcePath), potentialForError);
////                        }
////                        break;

////                    case "Personal Preferences":


////                        List<DLLTemplate.FitToConceptModel> personalPref = new List<DLLTemplate.FitToConceptModel>();


////                        personalPref = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

////                        if (personalPref.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            //lok needs a fix RADIANT : error (Personal Preference testname null)

////                            // notAvailable(this.template);
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PersonalPreferences" + personalPref.Count + ".pptx";
////                            dLLClass.PersonalPreferencesMethod(CreateTargetPath(sourcePath), personalPref, HistoricalMeanType, HistoricalMeanDescription);
////                        }
////                        break;

////                    case "Likeability":


////                        List<DLLTemplate.OverallImpressionModel> likeability = new List<DLLTemplate.OverallImpressionModel>();
////                        likeability = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

////                        if (likeability.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Likeability" + likeability.Count + ".pptx";
////                            dLLClass.LikeabilityMethod(CreateTargetPath(sourcePath), likeability);
////                        }
////                        break;



////                    case "Likeability Rationale":
////                        List<LikeabilityRationaleModel> likeabilyRationalesData = new List<LikeabilityRationaleModel>();
////                        likeabilyRationalesData = JsonConvert.DeserializeObject<List<LikeabilityRationaleModel>>(data);

////                        if (likeabilyRationalesData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationale.pptx";
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationaleOrdered.pptx";
////                            dLLClass.LikeabilityRationaleMethod(CreateTargetPath(sourcePath), likeabilyRationalesData, breakdown);
////                        }
////                        break;

////                    case "Memorability":


////                        List<DLLTemplate.FitToConceptModel> memorability = new List<DLLTemplate.FitToConceptModel>();

////                        memorability = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

////                        if (memorability.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Memorability" + memorability.Count + ".pptx";
////                            dLLClass.MemorabilityMethod(CreateTargetPath(sourcePath), memorability, HistoricalMeanType, HistoricalMeanDescription);
////                        }
////                        break;

////                    case "Verbal Understanding - Bar":
////                    case "Verbal Understanding":


////                        List<DLLTemplate.FitToConceptModel> verbalUnde = new List<DLLTemplate.FitToConceptModel>();


////                        verbalUnde = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

////                        if (verbalUnde.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\VerbalUnderstanding-Bar" + verbalUnde.Count + ".pptx";
////                            dLLClass.VerbalUnderstandingBarMethod(CreateTargetPath(sourcePath), verbalUnde);
////                        }
////                        break;

////                    case "Overall Impressions":


////                        List<DLLTemplate.OverallImpressionNewModel> overallImpression = new List<DLLTemplate.OverallImpressionNewModel>();



////                        overallImpression = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

////                        if (overallImpression.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\OverallImpressions" + overallImpression.Count + ".pptx";
////                            dLLClass.OverallImpressionsMethod(CreateTargetPath(sourcePath), overallImpression);
////                        }
////                        break;

////                    case "Suffix":
////                    case "Suffix Meaning":
////                    case "Existing Abbreviation ID":
////                    case "Existing Suffix ID":


////                        List<DLLTemplate.OverallImpressionModel> suffixOverallImpressionData = new List<DLLTemplate.OverallImpressionModel>();

////                        suffixOverallImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

////                        if (suffixOverallImpressionData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Suffix" + suffixOverallImpressionData.Count + ".pptx";
////                            dLLClass.SuffixMethod(CreateTargetPath(sourcePath), suffixOverallImpressionData, breakdown);
////                        }
////                        break;

////                    case "PromotionalReview":

////                        List<DLLTemplate.OverallImpressionModel> promotionalImpressionData = new List<DLLTemplate.OverallImpressionModel>();

////                        promotionalImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

////                        if (promotionalImpressionData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PromotionalReview" + promotionalImpressionData.Count + ".pptx";
////                            dLLClass.PromotionalReviewMethod(CreateTargetPath(sourcePath), promotionalImpressionData);
////                        }
////                        break;

////                    case "Ease Of Pronounciation":
////                    case "Ease of Pronunciation":



////                        List<DLLTemplate.FitToConceptModel> easeOfPronounicationData = new List<DLLTemplate.FitToConceptModel>();
////                        easeOfPronounicationData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

////                        if (easeOfPronounicationData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfPronunciation" + easeOfPronounicationData.Count + ".pptx";
////                            dLLClass.EaseOfPronounicationMethod(CreateTargetPath(sourcePath), easeOfPronounicationData, HistoricalMeanType, HistoricalMeanDescription);
////                        }
////                        break;

////                    case "Ease Of Spelling":

////                        List<DLLTemplate.FitToConceptNewModel> easeOfSpellingData = new List<DLLTemplate.FitToConceptNewModel>();

////                        easeOfSpellingData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

////                        if (easeOfSpellingData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfSpelling" + easeOfSpellingData.Count + ".pptx";
////                            dLLClass.EaseOfSpellingMethod(CreateTargetPath(sourcePath), easeOfSpellingData);
////                        }

////                        break;

////                    case "03 Exagg":
////                    case "01 Untrue":
////                    case "02 Misleading":
////                    case "02 Mislead":
////                    case "03 Exaggerative":
////                    case "Exaggerative-Inappropriate":


////                        List<DLLTemplate.OverallImpressionNewModel> ExaggData = new List<DLLTemplate.OverallImpressionNewModel>();
////                        ExaggData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);
////                        if (ExaggData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            Console.WriteLine(ExaggData);

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative" + ExaggData.Count + ".pptx";
////                            dLLClass.ExaggerativeMethod(CreateTargetPath(sourcePath), ExaggData, this.template);
////                        }
////                        break;


////                    case "Initial Recall":

////                        List<DLLTemplate.FitToConceptNewModel> Vaynsmr_heme_initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();

////                        Vaynsmr_heme_initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

////                        if (Vaynsmr_heme_initialRecallData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + Vaynsmr_heme_initialRecallData.Count + ".pptx";
////                            dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), Vaynsmr_heme_initialRecallData, HistoricalMeanType, HistoricalMeanDescription);
////                        }
////                        break;

////                    case "Innovation":
////                        List<DLLTemplate.InnovationModel> doraemonInnovationData = new List<DLLTemplate.InnovationModel>();
////                        doraemonInnovationData = JsonConvert.DeserializeObject<List<DLLTemplate.InnovationModel>>(data);
////                        if (doraemonInnovationData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Innovation" + doraemonInnovationData.Count + ".pptx";
////                            dLLClass.InnovationMethod(CreateTargetPath(sourcePath), doraemonInnovationData, this.HistoricalMeanType, this.HistoricalMeanDescription);
////                        }
////                        break;


////                    case "Modifier":

////                        List<DLLTemplate.OverallImpressionNewModel> atemporalModifierData = new List<DLLTemplate.OverallImpressionNewModel>();

////                        atemporalModifierData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

////                        if (atemporalModifierData.Count == 0)
////                        {

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Modifier" + atemporalModifierData.Count + ".pptx";
////                            dLLClass.ModifierMethod(CreateTargetPath(sourcePath), atemporalModifierData, group);
////                        }
////                        break;

////                    case "Written Understanding - Bar":
////                    case "Written Understanding":


////                        List<DLLTemplate.WrittenUnderstadingModel> writtenUnde = new List<DLLTemplate.WrittenUnderstadingModel>();

////                        writtenUnde = JsonConvert.DeserializeObject<List<DLLTemplate.WrittenUnderstadingModel>>(data, jsonSettings);

////                        if (writtenUnde.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\WrittenUnderstanding" + writtenUnde.Count + ".pptx";
////                            dLLClass.WrittenUnderstandingMethod(CreateTargetPath(sourcePath), writtenUnde);
////                        }
////                        break;




////                    case "Preference Ranking":
////                        List<PreferenceRankingModel> preferenceRankingData = new List<PreferenceRankingModel>();
////                        preferenceRankingData = JsonConvert.DeserializeObject<List<PreferenceRankingModel>>(data);

////                        if (preferenceRankingData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            List<int> counterList = new List<int>();

////                            int count = 0;

////                            for (int j = 0; j < preferenceRankingData.Count; j++)
////                            {
////                                var apiDataEl = preferenceRankingData[j];
////                                if (apiDataEl.intRankCount1 != 0)
////                                {
////                                    count = 1;
////                                }
////                                if (apiDataEl.intRankCount2 != 0)
////                                {
////                                    count = 2;
////                                }
////                                if (apiDataEl.intRankCount3 != 0)
////                                {
////                                    count = 3;
////                                }
////                                if (apiDataEl.intRankCount4 != 0)
////                                {
////                                    count = 4;
////                                }
////                                if (apiDataEl.intRankCount5 != 0)
////                                {
////                                    count = 5;
////                                }
////                                if (apiDataEl.intRankCount6 != 0)
////                                {
////                                    count = 6;
////                                }
////                                if (apiDataEl.intRankCount7 != 0)
////                                {
////                                    count = 7;
////                                }
////                                if (apiDataEl.intRankCount8 != 0)
////                                {
////                                    count = 8;
////                                }
////                                if (apiDataEl.intRankCount9 != 0)
////                                {
////                                    count = 9;
////                                }
////                                counterList.Add(count);
////                            }

////                            List<PreferenceRankingSortedDataModel> apiDataWithWeightedList = new List<PreferenceRankingSortedDataModel>();

////                            List<int> lstWeightedData = new List<int>();

////                            int weightedSum = 0;

////                            for (int j = 0; j < preferenceRankingData.Count; j++)
////                            {
////                                PreferenceRankingSortedDataModel preferenceRankingSortedData = new PreferenceRankingSortedDataModel();

////                                var counterVal = counterList[j];

////                                weightedSum = 0;
////                                var apiDataEl = preferenceRankingData[j];

////                                preferenceRankingSortedData.strTestName = apiDataEl.strTestName;

////                                if (apiDataEl.intRankCount1 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount1 * counterVal;
////                                    preferenceRankingSortedData.intRankCount1 = apiDataEl.intRankCount1;

////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount2 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount2 * counterVal;
////                                    preferenceRankingSortedData.intRankCount2 = apiDataEl.intRankCount2;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount3 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount3 * counterVal;
////                                    preferenceRankingSortedData.intRankCount3 = apiDataEl.intRankCount3;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount4 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount4 * counterVal;
////                                    preferenceRankingSortedData.intRankCount4 = apiDataEl.intRankCount4;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount5 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount5 * counterVal;
////                                    preferenceRankingSortedData.intRankCount5 = apiDataEl.intRankCount5;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount6 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount6 * counterVal;
////                                    preferenceRankingSortedData.intRankCount6 = apiDataEl.intRankCount6;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount7 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount7 * counterVal;
////                                    preferenceRankingSortedData.intRankCount7 = apiDataEl.intRankCount7;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount8 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount8 * counterVal;
////                                    preferenceRankingSortedData.intRankCount8 = apiDataEl.intRankCount8;
////                                    counterVal--;
////                                }
////                                if (apiDataEl.intRankCount9 != 0)
////                                {
////                                    weightedSum += apiDataEl.intRankCount9 * counterVal;
////                                    preferenceRankingSortedData.intRankCount9 = apiDataEl.intRankCount9;
////                                    counterVal--;
////                                }

////                                preferenceRankingSortedData.weightedScore = weightedSum;
////                                lstWeightedData.Add(weightedSum);
////                                apiDataWithWeightedList.Add(preferenceRankingSortedData);
////                            }

////                            var max = apiDataWithWeightedList.Max(m => m.weightedScore);

////                            if (max > 100 && max < 200)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking200Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 200 && max < 300)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking300Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 300 && max < 400)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking400Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 400 && max < 500)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking500Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 500 && max < 600)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking600Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 600 && max < 700)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking700Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 700 && max < 800)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking800Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 800 && max < 900)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking900Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                            if (max > 900 && max < 1000)
////                            {
////                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking1000Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
////                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
////                            }
////                        }

////                        break;




////                    case "Chemical Structure Appropriateness":
////                    case "Chemical Structure":
////                        List<ChemicalModel> chemicalData = new List<ChemicalModel>();
////                        chemicalData = JsonConvert.DeserializeObject<List<ChemicalModel>>(data);

////                        if (chemicalData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ChemicalStructure" + chemicalData.Count + ".pptx";
////                            dLLClass.ChemicalStructureMethod(CreateTargetPath(sourcePath), chemicalData);
////                        }
////                        break;









////                    case "Fit to Therapeutic Class":

////                        List<FitToTherapeuticClassModel> therapeuticClassData = new List<FitToTherapeuticClassModel>();
////                        therapeuticClassData = JsonConvert.DeserializeObject<List<FitToTherapeuticClassModel>>(data);

////                        if (therapeuticClassData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticClass" + therapeuticClassData.Count + ".pptx";
////                            dLLClass.FitToTherapeuticClassMethod(CreateTargetPath(sourcePath), therapeuticClassData, this.breakdown);
////                        }
////                        break;



////                    case "Associations":
////                        List<AssociationsModel> associationsData = new List<AssociationsModel>();
////                        associationsData = JsonConvert.DeserializeObject<List<AssociationsModel>>(data, jsonSettings);

////                        if (associationsData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Associations.pptx";
////                            dLLClass.AssociationsMethod(CreateTargetPath(sourcePath), associationsData, breakdown);
////                        }
////                        break;




////                    case "QTC":
////                        List<QTCModel> qtcData = new List<QTCModel>();
////                        qtcData = JsonConvert.DeserializeObject<List<QTCModel>>(data, jsonSettings);

////                        if (qtcData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTC.pptx";
////                            dLLClass.QTCMethod(CreateTargetPath(sourcePath), qtcData, breakdown);
////                        }
////                        break;


////                    case "QTCCustom":
////                        List<QTCCustomModel> qtcCustomData = new List<QTCCustomModel>();
////                        qtcCustomData = JsonConvert.DeserializeObject<List<QTCCustomModel>>(data);

////                        int counter = 0;

////                        for (int i = 0; i < qtcCustomData.Count; i++)
////                        {
////                            var dataEl = qtcCustomData[i];
////                            if (dataEl.strPageType1 != null)
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType2 != null)
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType3 != null)
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType4 != null)
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType5 != null)
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType6 != "")
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType7 != "")
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType8 != "")
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageType9 != "")
////                            {
////                                counter++;
////                            }
////                            if (dataEl.strPageName10 != "")
////                            {
////                                counter++;
////                            }
////                            break;
////                        }

////                        if (qtcCustomData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTCCustom" + (counter * 2) + ".pptx";
////                            dLLClass.QTCCustomMethod(CreateTargetPath(sourcePath), qtcCustomData, this.breakdown);
////                        }
////                        break;



////                    case "Reflective of Mechanism of Action":
////                        List<ReflectiveOfMechanismOfActionModel> reflectiveOfMechanismsData = new List<ReflectiveOfMechanismOfActionModel>();
////                        reflectiveOfMechanismsData = JsonConvert.DeserializeObject<List<ReflectiveOfMechanismOfActionModel>>(data, jsonSettings);

////                        if (reflectiveOfMechanismsData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ReflectiveofMechanismofAction" + reflectiveOfMechanismsData.Count + ".pptx";
////                            dLLClass.ReflectiveofMechanismofActionMethod(CreateTargetPath(sourcePath), reflectiveOfMechanismsData);
////                        }
////                        break;


////                    case "PhoneticTesting":
////                        List<PhoneticTestingModel> phoneticData = new List<PhoneticTestingModel>();
////                        phoneticData = JsonConvert.DeserializeObject<List<PhoneticTestingModel>>(data, jsonSettings);

////                        if (phoneticData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PhoneticTesting" + phoneticData.Count + ".pptx";
////                            dLLClass.PhoneticTestingMethod(CreateTargetPath(sourcePath), phoneticData);
////                        }
////                        break;



////                    case "Modifier Confusion":
////                        List<OverallImpressionNewModel> modifierConfusionData = new List<OverallImpressionNewModel>();
////                        modifierConfusionData = JsonConvert.DeserializeObject<List<OverallImpressionNewModel>>(data, jsonSettings);

////                        if (modifierConfusionData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ModifierConfusion" + modifierConfusionData.Count + ".pptx";
////                            dLLClass.ModifierConfusionMethod(CreateTargetPath(sourcePath), modifierConfusionData, group);
////                        }
////                        break;


////                    case "Medical Terms":
////                    case "Medical Terms Similarity":

////                        List<MedicalTermsModel> medicalData = new List<MedicalTermsModel>();
////                        medicalData = JsonConvert.DeserializeObject<List<MedicalTermsModel>>(data);

////                        if (medicalData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);

////                        }
////                        else
////                        {
////                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholder.pptx";
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholderRow.pptx";
////                            dLLClass.MedicalTermsMethod(CreateTargetPath(sourcePath), medicalData);
////                        }
////                        break;


////                    case "Non-Medical Terms":
////                    case "Non - Medical Terms Similarity":
////                    case "Non-Medical Terms Similarity":

////                        List<NonMedicalTermsModel> nonMedicalData = new List<NonMedicalTermsModel>();
////                        nonMedicalData = JsonConvert.DeserializeObject<List<NonMedicalTermsModel>>(data);

////                        if (nonMedicalData.Count == 0)
////                        {
////                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholderRow.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholder.pptx";
////                            dLLClass.NonMedicalTermsMethod(CreateTargetPath(sourcePath), nonMedicalData);
////                        }
////                        break;





////                    case "BRANDEX LOGO":
////                        List<BrandexLogoModel> brandexLogoData = new List<BrandexLogoModel>();
////                        brandexLogoData = JsonConvert.DeserializeObject<List<BrandexLogoModel>>(data);

////                        if (brandexLogoData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            List<BrandexLogoShortModel> brandexLogoShortData = new List<BrandexLogoShortModel>();

////                            double average1MaxLogo = 0.0;
////                            double average2MaxLogo = 0.0;
////                            double average3MaxLogo = 0.0;
////                            double average4MaxLogo = 0.0;
////                            double average5MaxLogo = 0.0;

////                            for (int i = 0; i < brandexLogoData.Count; i++)
////                            {
////                                var dblAverage1max = brandexLogoData[i].dblAveragePage1;
////                                if (dblAverage1max == 0.0)
////                                {
////                                    average1MaxLogo = 0;
////                                }
////                                if (average1MaxLogo < dblAverage1max)
////                                {
////                                    average1MaxLogo = dblAverage1max;
////                                }

////                                var dblAverage2max = brandexLogoData[i].dblAveragePage2;
////                                if (dblAverage2max == 0.0)
////                                {
////                                    average2MaxLogo = 0;
////                                }
////                                if (average2MaxLogo < dblAverage2max)
////                                {
////                                    average2MaxLogo = dblAverage2max;
////                                }

////                                var dblAverage3max = brandexLogoData[i].dblAveragePage3;
////                                if (dblAverage3max == 0.0)
////                                {
////                                    average3MaxLogo = 0;
////                                }
////                                if (average3MaxLogo < dblAverage3max)
////                                {
////                                    average3MaxLogo = dblAverage3max;
////                                }

////                                var dblAverage4max = brandexLogoData[i].dblAveragePage4;
////                                if (dblAverage4max == 0.0)
////                                {
////                                    average4MaxLogo = 0;
////                                }
////                                if (average4MaxLogo < dblAverage4max)
////                                {
////                                    average4MaxLogo = dblAverage4max;
////                                }

////                                var dblAverage5max = brandexLogoData[i].dblAveragePage5;
////                                if (dblAverage5max == 0.0)
////                                {
////                                    average5MaxLogo = 0;
////                                }
////                                if (average5MaxLogo < dblAverage5max)
////                                {
////                                    average5MaxLogo = dblAverage5max;
////                                }
////                            }

////                            double brandexLogoscalingFactor = 1.105380082;

////                            for (int i = 0; i < brandexLogoData.Count; i++)
////                            {
////                                //for the table
////                                BrandexLogoShortModel brandexLogoShortModelData = new BrandexLogoShortModel();

////                                var dataEl = brandexLogoData[i];

////                                double averagePage1WeightedValue = 0.0;
////                                double averagePage2WeightedValue = 0.0;
////                                double averagePage3WeightedValue = 0.0;
////                                double averagePage4WeightedValue = 0.0;
////                                double averagePage5WeightedValue = 0.0;

////                                if (average1MaxLogo > 0)
////                                {
////                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValue = 0;
////                                }

////                                if (average2MaxLogo > 0)
////                                {
////                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValue = 0;
////                                }

////                                if (average3MaxLogo > 0)
////                                {
////                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValue = 0;
////                                }

////                                if (average4MaxLogo > 0)
////                                {
////                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValue = 0;
////                                }

////                                if (average5MaxLogo > 0)
////                                {
////                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight;
////                                }
////                                else
////                                {
////                                    averagePage5WeightedValue = 0;
////                                }

////                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
////                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexLogoscalingFactor;

////                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

////                                brandexLogoShortModelData.dblAveragePage1 = dataEl.dblAveragePage1;
////                                brandexLogoShortModelData.dblPage1Weight = dataEl.dblPage1Weight;
////                                brandexLogoShortModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

////                                brandexLogoShortModelData.dblAveragePage2 = dataEl.dblAveragePage2;
////                                brandexLogoShortModelData.dblPage2Weight = dataEl.dblPage2Weight;
////                                brandexLogoShortModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

////                                brandexLogoShortModelData.dblAveragePage3 = dataEl.dblAveragePage3;
////                                brandexLogoShortModelData.dblPage3Weight = dataEl.dblPage3Weight;
////                                brandexLogoShortModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

////                                brandexLogoShortModelData.dblAveragePage4 = dataEl.dblAveragePage4;
////                                brandexLogoShortModelData.dblPage4Weight = dataEl.dblPage4Weight;
////                                brandexLogoShortModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

////                                brandexLogoShortModelData.dblAveragePage5 = dataEl.dblAveragePage5;
////                                brandexLogoShortModelData.dblPage5Weight = dataEl.dblPage5Weight;
////                                brandexLogoShortModelData.dblAveragePage5Weighted = averagePage5WeightedValue;

////                                brandexLogoShortModelData.dblAveragePage6 = dataEl.dblAveragePage6;
////                                brandexLogoShortModelData.dblPage6Weight = dataEl.dblPage6Weight;
////                                brandexLogoShortModelData.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

////                                brandexLogoShortModelData.dblAveragePage7 = dataEl.dblAveragePage7;
////                                brandexLogoShortModelData.dblPage7Weight = dataEl.dblPage7Weight;
////                                brandexLogoShortModelData.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

////                                brandexLogoShortModelData.dblAveragePage8 = dataEl.dblAveragePage8;
////                                brandexLogoShortModelData.dblPage8Weight = dataEl.dblPage8Weight;
////                                brandexLogoShortModelData.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

////                                brandexLogoShortModelData.dblIndex = indexSum;
////                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
////                                brandexLogoShortModelData.intRed = dataEl.intRed;
////                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
////                                brandexLogoShortModelData.intBlue = dataEl.intBlue;


////                                //for the chart - 
////                                double averagePage1WeightedValueForChart = 0.0;
////                                double averagePage2WeightedValueForChart = 0.0;
////                                double averagePage3WeightedValueForChart = 0.0;
////                                double averagePage4WeightedValueForChart = 0.0;
////                                double averagePage5WeightedValueForChart = 0.0;

////                                if (average1MaxLogo > 0)
////                                {
////                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight) * brandexLogoscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForChart = 0;
////                                }

////                                if (average2MaxLogo > 0)
////                                {
////                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight) * brandexLogoscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForChart = 0;
////                                }

////                                if (average3MaxLogo > 0)
////                                {
////                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight) * brandexLogoscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValueForChart = 0;
////                                }

////                                if (average4MaxLogo > 0)
////                                {
////                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight) * brandexLogoscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValueForChart = 0;
////                                }

////                                if (average5MaxLogo > 0)
////                                {
////                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight) * brandexLogoscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage5WeightedValueForChart = 0;
////                                }

////                                double indexSumForChart = averagePage1WeightedValueForChart +
////                                                          averagePage2WeightedValueForChart +
////                                                          averagePage3WeightedValueForChart +
////                                                          averagePage4WeightedValueForChart +
////                                                          averagePage5WeightedValueForChart;

////                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

////                                brandexLogoShortModelData.dblAveragePage1WeightedForChart = Math.Round(averagePage1WeightedValueForChart, 1);

////                                brandexLogoShortModelData.dblAveragePage2WeightedForChart = Math.Round(averagePage2WeightedValueForChart, 1);

////                                brandexLogoShortModelData.dblAveragePage3WeightedForChart = Math.Round(averagePage3WeightedValueForChart, 1);

////                                brandexLogoShortModelData.dblAveragePage4WeightedForChart = Math.Round(averagePage4WeightedValueForChart, 1); ;

////                                brandexLogoShortModelData.dblAveragePage5WeightedForChart = Math.Round(averagePage5WeightedValueForChart, 1); ;

////                                brandexLogoShortModelData.dblIndexForChart = indexSumForChart;
////                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
////                                brandexLogoShortModelData.intRed = dataEl.intRed;
////                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
////                                brandexLogoShortModelData.intBlue = dataEl.intBlue;

////                                brandexLogoShortData.Add(brandexLogoShortModelData);
////                            }

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexLogoTemplate" + brandexLogoShortData.Count + ".pptx";

////                            dLLClass.BrandexLogoMethod(CreateTargetPath(sourcePath), brandexLogoShortData, this.breakdown);
////                        }
////                        break;








////                    case "SALA":
////                    case "Sound Alike-Look Alike":
////                        List<SALANewModel> salaData = new List<SALANewModel>();
////                        salaData = JsonConvert.DeserializeObject<List<SALANewModel>>(data, jsonSettings);

////                        if (salaData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            CreateTargetPath(sourcePath);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";
////                            dLLClass.SALANewMethod(CreateTargetPath(sourcePath), salaData);
////                        }
////                        break;





////                    case "Negative Communication":
////                    case "Negative or Offensive Communication":


////                        List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
////                        negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data);

////                        if (negCommData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + negCommData.Count + ".pptx";
////                            dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), negCommData);
////                        }
////                        break;


////                    case "Negative Communication Rationale":


////                        List<NegativeCommunicationRationaleModel> negCommRationaleData = new List<NegativeCommunicationRationaleModel>();
////                        negCommRationaleData = JsonConvert.DeserializeObject<List<NegativeCommunicationRationaleModel>>(data);

////                        if (negCommRationaleData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunicationRationalePlaceholder.pptx";
////                            dLLClass.NegativeCommRationaleMethod(CreateTargetPath(sourcePath), negCommRationaleData, this.breakdown);
////                        }
////                        break;


////                    case "Fit to Theraputic Category":
////                    case "Fit to Therapeutic Category":
////                        List<FitToTheraputicCategoryModel> theraputicData = new List<FitToTheraputicCategoryModel>();
////                        theraputicData = JsonConvert.DeserializeObject<List<FitToTheraputicCategoryModel>>(data);

////                        List<FitToTheraputicCategoryModel> testnameCountList = new List<FitToTheraputicCategoryModel>();

////                        for (int i = 0; i < theraputicData.Count; i++)
////                        {
////                            var groupData = theraputicData.GroupBy(x => x.strTestName).ToList();
////                            if (groupData != null)
////                            {
////                                var firstTestnameList = groupData[0].ToList();
////                                testnameCountList = firstTestnameList;
////                            }
////                            break;
////                        }

////                        int tableHeaderCounter = testnameCountList.Count;

////                        if (theraputicData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticCategory" + tableHeaderCounter + ".pptx";
////                            dLLClass.FitToTheraputicMethod(CreateTargetPath(sourcePath), theraputicData);
////                        }
////                        break;






////                    //    List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
////                    //    negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data, jsonSettings);


////                    //    if (negCommData.Count == 0)
////                    //    {
////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                    //        CreateTargetPath(sourcePath);
////                    //    }
////                    //    else
////                    //    {
////                    //        if (this.template == "Sound Alike-Look Alike")
////                    //        {
////                    //            this.template = "SALA";
////                    //        }

////                    //        //sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";


////                    //        int dataIndex = 0;
////                    //        double sum = 0;
////                    //        List<NegativeCommunicationModelShort> lstNegCommValue = new List<NegativeCommunicationModelShort>();

////                    //        var groupApiData = negCommData.GroupBy(item => item.strTestName).OrderBy(group => group.Key).ToList();


////                    //        while (dataIndex < groupApiData.Count)
////                    //        {
////                    //            sum = 0;
////                    //            var testNameGroup = groupApiData[dataIndex];
////                    //            var testNameData = testNameGroup.ToList();

////                    //            foreach (var Tdata in testNameData)
////                    //            {
////                    //                sum += Tdata.intSum;
////                    //            }

////                    //            double total = testNameData.FirstOrDefault().intTotal;

////                    //            double percentageVal = (double)(((total - sum) / total) * 100);

////                    //            double remainingPercentageVal = 100 - percentageVal;

////                    //            // lstNegCommValue.FirstOrDefault().percentage = percentageVal.ToString();

////                    //            NegativeCommunicationModelShort ncomShort = new NegativeCommunicationModelShort();

////                    //            ncomShort.percentage = percentageVal.ToString();
////                    //            ncomShort.strTestName = testNameData.FirstOrDefault().strTestName;
////                    //            ncomShort.intBlue = testNameData.Max(x => x.intBlue);
////                    //            ncomShort.intGreen = testNameData.Max(x => x.intGreen);
////                    //            ncomShort.intRed = testNameData.Max(x => x.intRed);
////                    //            ncomShort.remainingPercentage = remainingPercentageVal.ToString();

////                    //            lstNegCommValue.Add(ncomShort);

////                    //            dataIndex++;
////                    //        }


////                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + lstNegCommValue.Count + ".pptx";
////                    //        dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), lstNegCommValue);




////                    //    }
////                    //    break;


////                    case "Distinctiveness":
////                        List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
////                        distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

////                        if (distinctData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            notAvailable(sourcePath, template, breakdown);
////                        }
////                        else
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
////                            dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
////                        }
////                        break;



////                    case "BRANDEX MEDICAL DEVICE 1":
////                    case "Brandex Medical Device 1":

////                        List<BrandexMedicalDevice1Model> brandexMedicalData = new List<BrandexMedicalDevice1Model>();
////                        brandexMedicalData = JsonConvert.DeserializeObject<List<BrandexMedicalDevice1Model>>(data);

////                        if (brandexMedicalData.Count == 0)
////                        {
////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
////                        }
////                        else
////                        {
////                            List<BrandexMedicalDevice1ShortModel> brandexMedicalShortData = new List<BrandexMedicalDevice1ShortModel>();

////                            double brandexMedicalAverage1Max = 0.0;
////                            double brandexMedicalAverage2Max = 0.0;
////                            double brandexMedicalAverage3Max = 0.0;
////                            double brandexMedicalAverage4Max = 0.0;
////                            double brandexMedicalAverage5Max = 0.0;

////                            for (int i = 0; i < brandexMedicalData.Count; i++)
////                            {
////                                var dblAverage1max = brandexMedicalData[i].dblAveragePage1;
////                                if (dblAverage1max == 0.0)
////                                {
////                                    brandexMedicalAverage1Max = 0;
////                                }
////                                if (brandexMedicalAverage1Max < dblAverage1max)
////                                {
////                                    brandexMedicalAverage1Max = dblAverage1max;
////                                }

////                                var dblAverage2max = brandexMedicalData[i].dblAveragePage2;
////                                if (dblAverage2max == 0.0)
////                                {
////                                    brandexMedicalAverage2Max = 0;
////                                }
////                                if (brandexMedicalAverage2Max < dblAverage2max)
////                                {
////                                    brandexMedicalAverage2Max = dblAverage2max;
////                                }

////                                var dblAverage3max = brandexMedicalData[i].dblAveragePage3;
////                                if (dblAverage3max == 0.0)
////                                {
////                                    brandexMedicalAverage3Max = 0;
////                                }
////                                if (brandexMedicalAverage3Max < dblAverage3max)
////                                {
////                                    brandexMedicalAverage3Max = dblAverage3max;
////                                }

////                                var dblAverage4max = brandexMedicalData[i].dblAveragePage4;
////                                if (dblAverage4max == 0.0)
////                                {
////                                    brandexMedicalAverage4Max = 0;
////                                }
////                                if (brandexMedicalAverage4Max < dblAverage4max)
////                                {
////                                    brandexMedicalAverage4Max = dblAverage4max;
////                                }

////                                var dblAverage5max = brandexMedicalData[i].dblAveragePage5;
////                                if (dblAverage5max == 0.0)
////                                {
////                                    brandexMedicalAverage5Max = 0;
////                                }
////                                if (brandexMedicalAverage5Max < dblAverage5max)
////                                {
////                                    brandexMedicalAverage5Max = dblAverage5max;
////                                }
////                            }

////                            double brandexMedicalscalingFactor = 1.0023;

////                            for (int i = 0; i < brandexMedicalData.Count; i++)
////                            {
////                                //for the table
////                                BrandexMedicalDevice1ShortModel brandexSafetyShortModel = new BrandexMedicalDevice1ShortModel();

////                                var dataEl = brandexMedicalData[i];

////                                double averagePage1WeightedValue = 0.0;
////                                double averagePage2WeightedValue = 0.0;
////                                double averagePage3WeightedValue = 0.0;
////                                double averagePage4WeightedValue = 0.0;
////                                double averagePage5WeightedValue = 0.0;

////                                if (brandexMedicalAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValue = 0;
////                                }

////                                if (brandexMedicalAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValue = 0;
////                                }

////                                if (brandexMedicalAverage3Max > 0)
////                                {
////                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValue = 0;
////                                }

////                                if (brandexMedicalAverage4Max > 0)
////                                {
////                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValue = 0;
////                                }

////                                if (brandexMedicalAverage5Max > 0)
////                                {
////                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight;
////                                }
////                                else
////                                {
////                                    averagePage5WeightedValue = 0;
////                                }

////                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
////                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexMedicalscalingFactor;

////                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
////                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
////                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
////                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
////                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
////                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
////                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
////                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
////                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
////                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
////                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

////                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
////                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
////                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

////                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
////                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
////                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

////                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
////                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
////                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

////                                brandexSafetyShortModel.dblIndex = indexSum;
////                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                                brandexSafetyShortModel.intRed = dataEl.intRed;
////                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                                //for the chart - 

////                                double averagePage1WeightedValueForChart = 0.0;
////                                double averagePage2WeightedValueForChart = 0.0;
////                                double averagePage3WeightedValueForChart = 0.0;
////                                double averagePage4WeightedValueForChart = 0.0;
////                                double averagePage5WeightedValueForChart = 0.0;

////                                if (brandexMedicalAverage1Max > 0)
////                                {
////                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight) * brandexMedicalscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage1WeightedValueForChart = 0;
////                                }

////                                if (brandexMedicalAverage2Max > 0)
////                                {
////                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight) * brandexMedicalscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage2WeightedValueForChart = 0;
////                                }

////                                if (brandexMedicalAverage3Max > 0)
////                                {
////                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight) * brandexMedicalscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage3WeightedValueForChart = 0;
////                                }

////                                if (brandexMedicalAverage4Max > 0)
////                                {
////                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight) * brandexMedicalscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage4WeightedValueForChart = 0;
////                                }

////                                if (brandexMedicalAverage5Max > 0)
////                                {
////                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight) * brandexMedicalscalingFactor;
////                                }
////                                else
////                                {
////                                    averagePage5WeightedValueForChart = 0;
////                                }

////                                double indexSumForChart = averagePage1WeightedValueForChart +
////                                                          averagePage2WeightedValueForChart +
////                                                          averagePage3WeightedValueForChart +
////                                                          averagePage4WeightedValueForChart +
////                                                          averagePage5WeightedValueForChart;

////                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

////                                brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

////                                brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
////                                ;
////                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
////                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
////                                brandexSafetyShortModel.intRed = dataEl.intRed;
////                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
////                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
////                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


////                                brandexMedicalShortData.Add(brandexSafetyShortModel);
////                            }

////                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexMedicalDevice1Template" + brandexMedicalShortData.Count + ".pptx";
////                            dLLClass.BrandexMedicalDevice1Method(CreateTargetPath(sourcePath), brandexMedicalShortData);
////                        }
////                        break;











////                    default:
////                        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                        notAvailable(CreateTargetPath(sourcePath), template, breakdown);
////                        break;


////                        //case "":
////                        //    List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
////                        //    distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

////                        //    if (distinctData.Count == 0)
////                        //    {
////                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                        //        notAvailable(sourcePath, template, breakdown);
////                        //    }
////                        //    else
////                        //    {
////                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
////                        //        dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
////                        //    }
////                        //    break;


////                        //default:
////                        //    sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
////                        //    notAvailable(CreateTargetPath(sourcePath), template, breakdown);
////                        //    break;



////                }
////            }




////        }


////        public async Task<string> addChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string BreakDown)
////        {
////            //hello
////            await Task.Run(() => fnaddChartsToFinalTemplate1(project, charts, finalTemplate, BreakDown));
////            return "Process sucessful";
////        }


////        public string getAttributeTitle(string project, string chart)
////        {
////            string repText = "";
////            System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

////            foreach (DataRow row in dt1.Rows)
////            {
////                repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
////            }

////            return repText;


////        }


////        public async Task fnaddChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string breakDown)
////        {
////            try
////            {
////                if (!Directory.Exists($"C:\\excelfiles\\{project}\\Final"))
////                {
////                    Directory.CreateDirectory($"C:\\excelfiles\\{project}\\Final");
////                }

////                //temp copy the final template file to the path 

////                //commented
////                // File.Copy("\\\\miafs02\\Market Research\\MR Programs\\ExcelCharts_Chartsdll\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", true);

////                //local path

////                File.Copy("C:\\ExcelChartFiles\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", true);

////                //get all the pagegroup names for the chart 

////                List<clschartPageGroup> lstPg = getProjectPagegroupNames(project);


////                //get the display names of the charts in excel sheet app

////                List<clschartPageDisplayName> lstChartDispNames = getProjectChartDisplayNames(project);

////                string op = "";

////                //get the page settings for the template 

////                //System.Data.DataTable dt = clsData.MRData.getDataTable("ExcelChartsPrc_getPPTFinalSettings " + "'" + "MR-Rx Naming" + "'," + "'" + "BI - 2024" + "'");
////                System.Data.DataTable dt = clsData.MRData.getDataTable("ExcelChartsPrc_getPPTFinalSettings " + "'" + $"{finalTemplate}" + "'," + "'" + "BI - 2024" + "'");

////                Console.WriteLine($"{finalTemplate}");

////                Console.WriteLine(dt.Rows);

////                List<clsPPTFinalSettings> lstPPTFinalSettings = new List<clsPPTFinalSettings>();

////                int DelLastPage = 0;
////                int DelFirstPage = 0;
////                List<int> intChartPages = new List<int>();


////                foreach (DataRow row in dt.Rows)
////                {
////                    // Create a new object for each row and add it to the list
////                    clsPPTFinalSettings data1 = new clsPPTFinalSettings
////                    {
////                        intPPTSlideIndexFirst = Convert.ToInt32(row["intPPTSlideIndexFirst"]),
////                        intPPTSlideIndexLast = Convert.ToInt32(row["intPPTSlideIndexLast"]),
////                        strTemplateName = Convert.ToString(row["strTemplateName"]),
////                        strTemplateSourcePath = Convert.ToString(row["strTemplateSourcePath"]),
////                        strPageGroupName = Convert.ToString(row["strPageGroupName"]),
////                        strPageGroupType = Convert.ToString(row["strPageGroupType"]),
////                    };
////                    lstPPTFinalSettings.Add(data1);
////                }



////                int specialChartTypeCount = 0;

////                //copy a final in the path

////                createFolder($"C:\\excelfiles\\{project}\\Final");
////                //copyFile("C:\\ExcelChartsTemplatesNew\\ExcelCharts_ChartTemplates\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx", $"C:\\excelfiles\\{project}\\Final\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx");

////                copyFile("C:\\ExcelChartFiles\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx");

////                lstPPTFinalSettings = lstPPTFinalSettings.OrderByDescending(p => p.intPPTSlideIndexFirst).ToList();
////                int chartsCompCount = 0;
////                List<string> chartsCompleted = new List<string>();
////                List<string> pageGroupNameCompleted = new List<string>();
////                List<clsChartDislayNamePageGroupName> lstchartDispPageGroupName = new List<clsChartDislayNamePageGroupName>();


////                //
////                foreach (clschartPageDisplayName obj in lstChartDispNames)
////                {
////                    foreach (string chart in charts)
////                    {
////                        if (obj?.strPageName == chart)
////                        {
////                            obj.isReportSelectedByUser = true;


////                            //updating Attribute evaluation cover page :


////                            if (obj.strPageType.ToLower() == "attribute evaluation")
////                            {
////                                if (obj.strPageName.Contains("1"))
////                                {
////                                    //update the slide 38

////                                    string repText = getAttributeTitle(project, chart);


////                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);



////                                    //await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);


////                                }

////                                if (obj.strPageName.Contains("2"))
////                                {
////                                    //update the slide 38

////                                    string repText = getAttributeTitle(project, chart);

////                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", repText, 38);

////                                    // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

////                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle2", repText, 38);


////                                }

////                                if (obj.strPageName.Contains("3"))
////                                {
////                                    //update the slide 38

////                                    string repText = getAttributeTitle(project, chart);


////                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", repText, 38);

////                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

////                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle3", repText, 38);


////                                }

////                                if (obj.strPageName.Contains("4"))
////                                {
////                                    //update the slide 38

////                                    string repText = getAttributeTitle(project, chart);

////                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", repText, 38);

////                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

////                                    // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle4", repText, 38);


////                                }


////                            }



////                        }

////                    }

////                }


////                //in case if the attribute evaluations are not there 

////                List<clschartPageDisplayName> lstAtts = lstChartDispNames.Where(p => p.strPageType.ToLower() == "attribute evaluation").ToList();


////                if (!lstAtts.Any(obj => obj.strPageName.Contains("1")) || lstAtts.Any(obj => obj.strPageName.Contains("1") && obj.isReportSelectedByUser == false))
////                {

////                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);

////                }

////                if (!lstAtts.Any(obj => obj.strPageName.Contains("2")) || lstAtts.Any(obj => obj.strPageName.Contains("2") && obj.isReportSelectedByUser == false))
////                {

////                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);

////                }


////                if (!lstAtts.Any(obj => obj.strPageName.Contains("3")) || lstAtts.Any(obj => obj.strPageName.Contains("3") && obj.isReportSelectedByUser == false))
////                {

////                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);

////                }

////                if (!lstAtts.Any(obj => obj.strPageName.Contains("4")) || lstAtts.Any(obj => obj.strPageName.Contains("4") && obj.isReportSelectedByUser == false))
////                {

////                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);

////                }






////                //merge


////                //bool first chart name replaced

////                bool firstChartTextReplaced = false;
////                foreach (clsPPTFinalSettings objclsPPT in lstPPTFinalSettings)
////                {
////                    if (chartsCompCount == charts.Count())
////                    {
////                        break;
////                    }

////                    specialChartTypeCount = 0;
////                    //check if chart type is selected  

////                    bool ifThePageTypeIsSelected = true;


////                    //first chart project name and details change 

////                    if (!firstChartTextReplaced)
////                    {
////                        DateTime currentDate = DateTime.Now;

////                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "ProjectName", project, 0);

////                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<", "", 0);

////                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", ">>", "", 0);

////                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Month", currentDate.ToString("MMMM"), 0);

////                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Year", DateTime.Now.Year.ToString(), 0);


////                        //replacing the attribute evaluation cover page :

////                        firstChartTextReplaced = true;

////                    }



////                    //delete the slides that are not part of the report 


////                    if (!lstChartDispNames.Any(p => p.strPageType == objclsPPT.strPageGroupType))
////                    {

////                        // skip the Attribute Evaluation Cover if the report has Attribute eavluation

////                        if (pageGroupNameCompleted.Any(z => z.Contains("Attribute")) && objclsPPT.strPageGroupType == "Attribute Evaluation Cover")
////                        {
////                            continue;
////                        }

////                        else
////                        {
////                            DelLastPage = objclsPPT.intPPTSlideIndexLast;
////                            DelFirstPage = objclsPPT.intPPTSlideIndexFirst;


////                            if (DelLastPage - DelFirstPage == 0)
////                            {
////                                for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
////                                {
////                                    //if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
////                                    //{
////                                    //    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                    //}

////                                    if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
////                                    {
////                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                    }
////                                }

////                            }

////                            else
////                            {
////                                for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
////                                {
////                                    if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
////                                    {
////                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                    }

////                                    //if (DelLastPage - k - 1 < CountSlides($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
////                                    //{
////                                    //    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                    //}

////                                }

////                            }


////                            continue;

////                        }





////                    }



////                    foreach (string chart in charts)
////                    {

////                        //if the chart is completed break

////                        if (pageGroupNameCompleted.Contains(objclsPPT.strPageGroupName))
////                        {
////                            break;
////                        }





////                        //get the chart group type

////                        List<clschartPageDisplayName> lstCheck = lstChartDispNames.Where(p => p.strPageName == chart).ToList();

////                        bool noproceed = false;

////                        //delete the slides from the template if a report is not selected .

////                        bool isReportSelected = false;


////                        if (lstChartDispNames.Any(item => item.strPageType == objclsPPT?.strPageGroupType))
////                        {
////                            if (objclsPPT?.strPageGroupType.ToLower() == "exaggerative-inappropriate")
////                            {

////                                //if( lstChartDispNames.Any(item =>item.strPageType.ToLower() == "exaggerative-inappropriate" && getNumbersFromString(item.strPageName)==getNumbersFromString(chart)))
////                                //{
////                                //     isReportSelected = true;
////                                //}



////                                List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "exaggerative-inappropriate").ToList();

////                                foreach (clschartPageDisplayName obj in lstFiltered)
////                                {
////                                    if (getNumbersFromString(obj.strPageName) == getNumbersFromString(objclsPPT?.strPageGroupName))
////                                    {

////                                        if ((bool)(obj?.isReportSelectedByUser))
////                                        {
////                                            isReportSelected = true;
////                                            break;

////                                        }

////                                    }
////                                }
////                            }

////                            else if (objclsPPT?.strPageGroupType.ToLower() == "attribute evaluation")
////                            {
////                                List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "attribute evaluation" || item.strPageType.ToLower() == "attribute evaluation aggregate").ToList();

////                                foreach (clschartPageDisplayName obj in lstFiltered)
////                                {
////                                    if (getNumbersFromString(obj?.strPageName) == getNumbersFromString(objclsPPT?.strPageGroupName))
////                                    {


////                                        if ((bool)(obj?.isReportSelectedByUser))
////                                        {
////                                            isReportSelected = true;
////                                            break;

////                                        }
////                                    }
////                                }
////                            }

////                            else
////                            {
////                                List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == objclsPPT?.strPageGroupName.ToLower()).ToList();

////                                foreach (clschartPageDisplayName obj in lstFiltered)
////                                {
////                                    if ((bool)(obj?.isReportSelectedByUser))
////                                    {
////                                        isReportSelected = true;
////                                        break;
////                                    }
////                                }
////                            }
////                        }


////                        //if (lstChartDispNames.Any(item => item.strPageType == objclsPPT.strPageGroupType))
////                        //{
////                        //    foreach (clschartPageDisplayName item in lstChartDispNames)
////                        //    {
////                        //        if (item.strPageType == objclsPPT.strPageGroupType && item.strPageName==chart)
////                        //        {
////                        //            if (!item.isReportSelectedByUser)
////                        //            {
////                        //                isReportSelected = false;
////                        //                break;
////                        //            }
////                        //        }
////                        //    }


////                        //}

////                        //else
////                        //{
////                        //    isReportSelected = false;


////                        //}


////                        if (!isReportSelected)
////                        {
////                            //delete if the project is not checked 
////                            DelLastPage = (int)(objclsPPT?.intPPTSlideIndexLast);
////                            DelFirstPage = (int)(objclsPPT?.intPPTSlideIndexFirst);

////                            if (DelLastPage - DelFirstPage == 1 && objclsPPT?.strPageGroupName.ToLower() == "attribute evaluation cover")
////                            {
////                                List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "attribute evaluation" && item.isReportSelectedByUser == true).ToList();

////                                if (lstFiltered.Count > 0)
////                                {
////                                    break;
////                                }

////                                else
////                                {
////                                    DelLastPage = DelFirstPage;
////                                }


////                            }

////                            if (!pageGroupNameCompleted.Contains(objclsPPT?.strPageGroupName) && (DelLastPage != 0 && DelFirstPage != 0))
////                            {
////                                if (DelLastPage - DelFirstPage == 0)
////                                {
////                                    for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
////                                    {
////                                        if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
////                                        {
////                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                        }

////                                        //if (DelLastPage - k - 1 < CountSlides($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
////                                        //{
////                                        //    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                        //}

////                                    }

////                                }

////                                else
////                                {
////                                    for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
////                                    {
////                                        if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
////                                        {
////                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                        }

////                                        //if (DelLastPage - k - 1 < CountSlides($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
////                                        //{
////                                        //    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                        //}

////                                    }
////                                }
////                            }
////                            break;
////                        }

////                        foreach (clschartPageDisplayName oj in lstCheck)
////                        {

////                            if (objclsPPT?.strPageGroupType.ToLower() == "exaggerative-inappropriate" || objclsPPT?.strPageGroupType.ToLower() == "attribute evaluation")
////                            {

////                                if (oj.strPageName.ToLower() == "attribute evaluation aggregate")
////                                {
////                                    oj.strPageType = "Attribute Evaluation";
////                                }

////                                if ((oj.strPageType == objclsPPT.strPageGroupType) && (oj.strPageName == chart) && (getNumbersFromString(objclsPPT.strPageGroupName) == getNumbersFromString(chart)))
////                                {
////                                    noproceed = false;
////                                    break;

////                                }

////                                else
////                                {
////                                    noproceed = true;
////                                    break;

////                                }

////                            }

////                            else
////                            {
////                                if ((oj.strPageType == objclsPPT.strPageGroupType) && (oj.strPageName == chart))
////                                {
////                                    noproceed = false;
////                                    break;

////                                }

////                                else
////                                {
////                                    noproceed = true;
////                                    break;

////                                }
////                            }
////                        }

////                        if (noproceed)
////                        {
////                            continue;
////                        }




////                        if (!chartsCompleted.Contains(chart))
////                        {

////                            if (objclsPPT?.strPageGroupName != "Main Cover")
////                            {
////                                DelLastPage = (int)(objclsPPT?.intPPTSlideIndexLast);
////                                DelFirstPage = (int)(objclsPPT?.intPPTSlideIndexFirst);


////                                //add

////                                string firstMatchingName = lstPg
////                               .Where(p => p.strPageGroup == chart)
////                               .Select(p => p.strPageGroupType)
////                               .FirstOrDefault();


////                                if (objclsPPT?.strPageGroupType != "Attribute Evaluation" && objclsPPT?.strPageGroupType != "Exaggerative-Inappropriate")
////                                {
////                                    //item array of pages with slide numbers 
////                                    intChartPages = new List<int>();

////                                    if (DelLastPage > DelFirstPage)
////                                    {
////                                        if (objclsPPT?.strPageGroupType == "Phonetic Testing" || objclsPPT?.strPageGroupType == "JSCAN")
////                                        {
////                                            intChartPages.Add(DelLastPage);

////                                        }

////                                        else if (objclsPPT?.strPageGroupType == "Sound Alike-Look Alike" || objclsPPT?.strPageGroupType == "Medical Terms Similarity" || objclsPPT.strPageGroupType == "Non-Medical Terms Similarity" || objclsPPT.strPageGroupType == "Brandex Strategic Distinctiveness" || objclsPPT.strPageGroupType == "Brandex Safety")
////                                        {
////                                            for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
////                                            {
////                                                intChartPages.Add(DelFirstPage + l);
////                                            }

////                                        }
////                                        else
////                                        {
////                                            for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
////                                            {
////                                                intChartPages.Add(DelFirstPage + l - 1);
////                                            }


////                                        }



////                                    }
////                                    else
////                                    {
////                                        intChartPages.Add(DelLastPage);
////                                    }



////                                    if (objclsPPT?.strPageGroupName.ToLower() == chart.ToLower() || objclsPPT?.strPageGroupName.ToLower() == firstMatchingName.ToLower())
////                                    {
////                                        for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
////                                        {
////                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
////                                            //op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - k - 1);

////                                            //Thread.Sleep(100);
////                                        }

////                                        if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
////                                        {
////                                            for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
////                                            {
////                                                intChartPages.RemoveAt(i);
////                                            }

////                                        }

////                                        if (File.Exists(getIndividualChartPath(chart, project, breakDown)))
////                                        {
////                                            try
////                                            {
////                                                op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);
////                                            }
////                                            catch (Exception)
////                                            {
////                                                throw;
////                                            }
////                                        }


////                                        chartsCompCount = chartsCompCount + 1;

////                                        lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT.strPageGroupName));

////                                        chartsCompleted.Add(chart);

////                                        pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);
////                                    }


////                                }

////                                else
////                                {
////                                    //item array of pages with slide numbers 
////                                    intChartPages = new List<int>();

////                                    if (DelLastPage > DelFirstPage && DelLastPage - DelFirstPage > 1)
////                                    {
////                                        for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
////                                        {
////                                            intChartPages.Add(DelFirstPage + l - 1);
////                                        }

////                                    }
////                                    else
////                                    {
////                                        intChartPages.Add(DelLastPage);
////                                    }

////                                    //get the attribute evaluation category and Exaggerative 

////                                    string PageType = lstChartDispNames.FirstOrDefault(item => item.strPageName == chart)?.strPageType;

////                                    string resPageType = "";

////                                    op = "";

////                                    if (PageType == "Attribute Evaluation" && chart.Contains("1"))
////                                    {

////                                        resPageType = "Attribute Evaluation 1";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");





////                                    }

////                                    else if (PageType == "Attribute Evaluation" && chart.Contains("2"))
////                                    {
////                                        resPageType = "Attribute Evaluation 2";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "FirstPage");
////                                        //await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "Attribute #2", 38);

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


////                                        //}



////                                    }

////                                    else if (PageType == "Attribute Evaluation" && chart.Contains("3"))
////                                    {

////                                        resPageType = "Attribute Evaluation 3";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "FirstPage");
////                                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "Attribute #3", 38);

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


////                                        //}



////                                    }

////                                    else if (PageType == "Attribute Evaluation" && chart.Contains("4"))
////                                    {
////                                        resPageType = "Attribute Evaluation 4";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "FirstPage");
////                                        await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "Attribute #4", 38);

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "", "", 38);


////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


////                                        //}

////                                        //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
////                                        //{
////                                        //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


////                                        //}


////                                    }


////                                    else if (PageType.Contains("Attribute Evaluation") && chart.Contains("Aggregate") && chart.Contains("Attribute"))
////                                    {
////                                        resPageType = "Attribute Evaluation Aggregate";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "FirstPage");


////                                    }


////                                    else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "1" || getNumbersFromString(chart) == "01"))
////                                    {
////                                        resPageType = "01 Untrue";

////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "LastPage");

////                                        // DelFirstPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "FirstPage");

////                                        DelFirstPage = DelLastPage;


////                                    }


////                                    else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "2" || getNumbersFromString(chart) == "02"))
////                                    {

////                                        resPageType = "02 Mislead";

////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "LastPage");

////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "FirstPage");


////                                    }


////                                    else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "3" || getNumbersFromString(chart) == "03"))
////                                    {

////                                        resPageType = "03 Exagg";

////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "LastPage");

////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "FirstPage");


////                                    }

////                                    //if (objclsPPT.strPageGroupName.ToLower() == resPageType.ToLower())
////                                    //{


////                                    if (resPageType == "01 Untrue")
////                                    {

////                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - 1);
////                                        //op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - 1);
////                                    }

////                                    else
////                                    {
////                                        for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
////                                        {
////                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
////                                            //op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", DelLastPage - k - 1);
////                                            //Thread.Sleep(100);
////                                        }
////                                    }




////                                    if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
////                                    {
////                                        for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
////                                        {
////                                            intChartPages.RemoveAt(i);
////                                        }

////                                    }

////                                    op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);



////                                    if (PageType == "Attribute Evaluation" && chart.Contains("1"))
////                                    {

////                                        resPageType = "Attribute Evaluation 1";
////                                        DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
////                                        DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");

////                                        //update the slide 38

////                                        //update the slide 38

////                                        string repText = "";
////                                        System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

////                                        foreach (DataRow row in dt1.Rows)
////                                        {
////                                            repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
////                                        }






////                                        await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);

////                                        await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

////                                        await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);



////                                        chartsCompCount = chartsCompCount + 1;

////                                        lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT.strPageGroupName));

////                                        chartsCompleted.Add(chart);

////                                        pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);

////                                    }
////                                    // }

////                                }

////                            }

////                        }
////                    }
////                }
////                //get the final file in  the folder

////                createFolder($"C:\\excelfiles\\{project}\\PPT");
////                createFolder($"C:\\excelfiles\\{project}\\PPT\\01 Final");


////                while (true)
////                {


////                    if (File.Exists($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
////                    {
////                        copyFile($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\PPT\\01 Final\\ {project}_{finalTemplate}_{breakDown}.pptx");


////                        break;

////                    }

////                }

////            }
////            catch (Exception)
////            {
////                throw;
////            }
////        }



////        private int getPageIndex(List<clsPPTFinalSettings> lst, string pageGroupName, string pageIndex)
////        {
////            try
////            {

////                int result = 0;

////                if (pageIndex == "LastPage")
////                {
////                    result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexLast;
////                }

////                else if (pageIndex == "FirstPage")
////                {

////                    result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexFirst;
////                }

////                return result;
////            }
////            catch (Exception)
////            {

////                throw;
////            }
////        }


////        private bool checkAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName)
////        {

////            bool result = false;
////            if (lstPageDispName.Any(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true))
////            {

////                result = true;

////            }

////            return result;

////        }

////        private bool checkParticularAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
////        {

////            bool result = false;

////            var lst = lstPageDispName.Where(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true);

////            foreach (clschartPageDisplayName obj in lst)
////            {
////                if (obj.strPageName.Contains(attValue.ToString()))
////                {

////                    result = true;

////                }

////            }

////            return result;

////        }


////        private bool checkParticularExaggIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
////        {

////            bool result = false;

////            var lst = lstPageDispName.Where(m => m.strPageType == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true && m.strPageName.Contains(attValue.ToString()));

////            foreach (clschartPageDisplayName obj in lst)
////            {
////                if (obj.strPageName.Contains(attValue.ToString()))
////                {
////                    result = true;

////                }

////            }


////            return result;

////        }


////        public static string getNumbersFromString(string input)
////        {
////            // Use a regular expression to match all digits in the string
////            string numbers = Regex.Replace(input, @"[^0-9]", "");
////            return numbers;
////        }


////        private bool checkParticularChartIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue, string chart)
////        {

////            bool result = false;

////            var lst = lstPageDispName.Where(m => m.strPageName == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true);

////            foreach (clschartPageDisplayName obj in lst)
////            {
////                if (obj.strPageName.Contains(attValue.ToString()))
////                {
////                    result = true;

////                }

////            }


////            return result;

////        }




////        private string getReverseChartName(string chartName)
////        {
////            if (chartName == "Attribute Evaluation 1")
////            {
////                return "Attribute 1";
////            }
////            else if (chartName == "Attribute Evaluation 2")
////            {
////                return "Attribute 2";
////            }

////            else if (chartName == "Attribute Evaluation 3")
////            {
////                return "Attribute 3";
////            }

////            else if (chartName == "Attribute Evaluation 4")
////            {
////                return "Attribute 4";
////            }


////            else
////            {
////                return chartName;
////            }

////        }



////        private int getmaxSlideid(string fileName)
////        {
////            int maxSlideNumber = 0;
////            if (!File.Exists(fileName))
////            {
////                return 0;
////            }

////            else
////            {
////                maxSlideNumber = 1;
////                using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
////                {

////                    PresentationPart presentationPart = doc.PresentationPart;

////                    maxSlideNumber = presentationPart.SlideParts.Count();
////                }

////                return Convert.ToInt32(maxSlideNumber);


////            }




////        }



////        private string getIndividualChartPath(string chart, string projectName, string breakDown)
////        {



////            string outputpath = $"C:\\excelfiles\\{projectName}\\{chart}_{breakDown}.pptx";


////            //get the right chart name 

////            return outputpath;

////        }

////        public int CountSlides(string presentationFile)
////        {

////            if (File.Exists(presentationFile))
////            {

////                using (PresentationDocument ppt = PresentationDocument.Open(presentationFile, false))
////                {
////                    // Get the presentation part of the presentation document
////                    PresentationPart presentationPart = ppt.PresentationPart;
////                    if (presentationPart == null || presentationPart.Presentation == null)
////                    {
////                        return 0;
////                    }

////                    // Get the slide ID list
////                    SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
////                    if (slideIdList == null)
////                    {
////                        return 0;
////                    }

////                    // Return the count of slide IDs
////                    return slideIdList.ChildElements.Count;
////                }




////            }

////            else
////            {
////                return 0;

////            }




////        }



////        private string getRightChartName(string pageName, int index)
////        {
////            if (pageName == "Attribute Evaluation" && index == 1)
////            {
////                return "Attribute Evaluation " + index;
////            }
////            else if (pageName == "Attribute Evaluation" && index == 2)
////            {
////                return "Attribute Evaluation " + index;
////            }
////            else if (pageName == "Attribute Evaluation" && index == 3)
////            {
////                return "Attribute Evaluation " + index;
////            }

////            else if (pageName == "Attribute Evaluation" && index == 3)
////            {
////                return "Attribute Evaluation " + index;
////            }

////            else if (pageName == "Exaggerative-Inappropriate" && index == 1)
////            {
////                return "01 Untrue";
////            }

////            else if (pageName == "Exaggerative-Inappropriate" && index == 2)
////            {
////                return "02 Misleading";
////            }

////            else if (pageName == "Exaggerative-Inappropriate" && index == 3)
////            {
////                return "03 Exagg";
////            }

////            else
////            {
////                return pageName;
////            }

////        }


////        private string getChartsPagegroupName(string projectName, string chartName)
////        {
////            string realchartName = "";

////            System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");



////            List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
////            {
////                strPageGroup = row.Field<string>("strPageGroup"),
////                strPageGroupType = row.Field<string>("strPageGroupType")
////            }).ToList();


////            List<clschartPageGroup> lstPgChart = lstPg.Where(p => p.strPageGroup == chartName).ToList();

////            foreach (clschartPageGroup obj in lstPgChart)
////            {

////                realchartName = obj.strPageGroupType;

////            }

////            return realchartName;

////        }

////        private List<clschartPageGroup> getProjectPagegroupNames(string projectName)
////        {
////            try
////            {
////                string realchartName = "";

////                System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");


////                List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
////                {
////                    strPageGroup = row.Field<string>("strPageGroup"),
////                    strPageGroupType = row.Field<string>("strPageGroupType")
////                }).ToList();


////                return lstPg;
////            }
////            catch (Exception)
////            {

////                throw;
////            }



////        }


////        private List<clschartPageDisplayName> getProjectChartDisplayNames(string projectName)
////        {
////            try
////            {

////                string realchartName = "";

////                System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPages]  " + 1 + ",'" + projectName + "'");

////                List<clschartPageDisplayName> lstPg = dt.AsEnumerable().Select(row => new clschartPageDisplayName
////                {
////                    strPageName = row.Field<string>("strPageName"),
////                    strPageType = row.Field<string>("strPageType"),
////                    intPageID = row.Field<Int32>("intPageID").ToString(),
////                    intPageIndex = row.Field<Int32>("intPageIndex").ToString(),
////                    intSurveyID = row.Field<Int32>("intSurveyID").ToString(),

////                }).ToList();


////                return lstPg;

////            }
////            catch (Exception)
////            {

////                throw;
////            }
////        }


////        public static void createFolder(string path)
////        {
////            try
////            {
////                if (!Directory.Exists(path))
////                {
////                    // Try to create the directory.
////                    DirectoryInfo di = Directory.CreateDirectory(path);
////                }
////            }
////            catch (IOException ioex)
////            {
////                Console.WriteLine(ioex.Message);
////            }


////        }


////        public static void copyFile(string source, string destination)
////        {
////            try
////            {
////                File.Copy(source, destination, true);

////            }
////            catch (Exception ex)
////            {
////                Console.WriteLine(ex.Message);
////            }


////        }


////        public static void notAvailable(string sourcePath, string template, string breakdown)
////        {
////            DLLClass dLLClass = new DLLClass();
////            APIWrapper obj = new APIWrapper();
////            dLLClass.NotAvailableMethod(sourcePath, template, breakdown);

////        }


////        public static void copyTemplates()
////        {

////            //lok please check currently doing only for the first time

////            string sourceDirectory = @"\\miafs02\Market Research\MR Programs\ExcelCharts_Chartsdll\Templates";
////            string destinationDirectory = "C:\\ExcelChartFiles\\Templates\\";

////            if (!Directory.Exists("C:\\ExcelChartFiles\\"))
////            {

////                if (!Directory.Exists("C:\\ExcelChartFiles\\"))
////                {
////                    // Create the directory if it doesn't exist
////                    Directory.CreateDirectory("C:\\ExcelChartFiles\\");
////                }

////                if (!Directory.Exists(destinationDirectory))
////                {
////                    // Create the directory if it doesn't exist
////                    Directory.CreateDirectory(destinationDirectory);

////                }

////                foreach (var file in Directory.GetFiles(sourceDirectory))
////                {
////                    File.Copy(file, Path.Combine(destinationDirectory, Path.GetFileName(file)), true);
////                }


////            }


////        }


////        public void notAvailable(string template)
////        {

////            //DLLClass obj = new DLLClass();
////            //obj.NotAvailableMethod(CreateTargetPath($"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx"), template);

////            //template.NotAvailableTemplate(destination, templateName, breakdown);

////        }


////        public void copyInsertSlide(string source, string destination, int pos)
////        {
////            clsMisc.CopySlide(source, destination, pos);

////        }








////    }
////}



//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Presentation;
//using Newtonsoft.Json;
//using OpenXmlDll;
//using prjData;
//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Net.Http;
//using System.Text.RegularExpressions;
//using System.Threading.Tasks;
//using System.Xml.Linq;
//using static OpenXmlDLLDotnetFramework.DLLTemplate;


//namespace OpenXmlDLLDotnetFramework
//{
//    public class APIWrapper
//    {
//        string project = "";
//        string template = "";
//        string templateType = "";
//        string finalTemplate = "";
//        string group = "";
//        string breakdown = "";
//        string HistoricalMeanType = "";
//        string HistoricalMeanDescription = "";
//        //#pragma warning disable CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used
//        string genPath = "";
//        //#pragma warning restore CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used

//        public APIWrapper() { }

//        public APIWrapper(string project,
//                            string template,
//                            string group,
//                            string breakdown,
//                            string HistoricalMeanType,
//                            string HistoricalMeanDescription, string finalTemplate)
//        {
//            this.project = project;
//            this.template = template;
//            this.group = group;
//            this.breakdown = breakdown;
//            this.HistoricalMeanDescription = HistoricalMeanDescription;
//            this.HistoricalMeanType = HistoricalMeanType;
//            this.finalTemplate = finalTemplate;
//        }

//        public async Task<string> CallAPI(string project,
//                                        string template,
//                                        string group,
//                                        string breakdown,
//                                        string HistoricalMeanType,
//                                        string HistoricalMeanDescription)
//        {
//            using (HttpClient http = new HttpClient())
//            {
//                const string URL = "https://tools.brandinstitute.com/wsXlCharts/wsExcelCharts.asmx/GetChartBreakDownData";

//                //this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");



//                var payload = new[]
//                {
//                new KeyValuePair<string,string>("token","2BF27A11-E318-447A-98FD-70AFE3871AA9"),
//                new KeyValuePair<string,string>("project",project),
//               new KeyValuePair<string,string>("template",template),
//               //new KeyValuePair<string,string>("template",this.templateType),
//                new KeyValuePair<string,string>("group",group),
//                new KeyValuePair<string,string>("breakdown",breakdown),
//                new KeyValuePair<string,string>("historicalMeanType",HistoricalMeanType),
//                new KeyValuePair<string,string>("historicalMeanDescription",HistoricalMeanDescription)
//            };

//                var content = new FormUrlEncodedContent(payload);

//                try
//                {
//                    var response = await http.PostAsync(URL, content);

//                    if (response.IsSuccessStatusCode)
//                    {
//                        return await response.Content.ReadAsStringAsync();
//                    }
//                    else
//                    {
//                        return $"Error: {response.StatusCode}, {response.ReasonPhrase}";
//                    }
//                }
//                catch (Exception ex)
//                {
//                    return ex.Message;
//                }
//            }
//        }

//        public string CreateTargetPath(string myTemplate)
//        {
//            string path = $"C:\\excelfiles\\{this.project}";
//            if (!Directory.Exists(path))
//            {
//                Directory.CreateDirectory(path);
//            }
//            path = $"C:\\excelfiles\\{this.project}\\{this.template}_{this.breakdown}.pptx";
//            File.Copy(myTemplate, path, true);
//            return path;
//        }

//        public string getChartPathToCopyToFinal(string myTemplate, string project, string breakdown)
//        {
//            string path = $"C:\\excelfiles\\{project}";
//            if (!Directory.Exists(path))
//            {
//                Directory.CreateDirectory(path);
//            }
//            path = $"C:\\excelfiles\\{project}\\{myTemplate}_{breakdown}.pptx";
//            return path;
//        }

//        public async Task OpenXMLParallelProcess(string project,
//                                            List<string> templates,
//                                            List<string> breakdowns,
//                                            string HistoricalMeanType,
//                                            string HistoricalMeanDescription, string finalTemplateName)
//        {

//            //copy the template folder to the local system

//            //copyTemplates();


//            List<Task> taskArr = new List<Task>();

//            foreach (var breakdown in breakdowns)
//            {
//                foreach (var template in templates)
//                {

//                    APIWrapper wrapper = new APIWrapper(project, template, template, breakdown, HistoricalMeanType, HistoricalMeanDescription, finalTemplateName);
//                    taskArr.Add(Task.Run(() => wrapper.Process()));
//                }
//            }
//            await Task.WhenAll(taskArr);



//            //adding the template to the final



//        }



//        public Task<clsFinalPageNumberRange> getPPTFinalPageSettings(string strFinalReport, string year, string chartName)
//        {
//            return Task.Run(() =>
//            {

//                System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPPTFinalSettings] " + "'" + strFinalReport + "'," + "'" + year + "'");


//                List<clsFinalPageModel> lstFinalPageModel = clsData.MRData.ConvertDataTableToListGeneric<clsFinalPageModel>(dt);


//                List<clsFinalPageModel> lstFinal = lstFinalPageModel.Where(n => n.strPageGroupName == chartName).ToList();

//                clsFinalPageNumberRange objPageNum = new clsFinalPageNumberRange();

//                foreach (clsFinalPageModel obj in lstFinal)
//                {
//                    objPageNum.firstPage = obj.intPPTSlideIndexFirst;
//                    objPageNum.lastPage = obj.intPPTSlideIndexLast;

//                }

//                //  Task<clsFinalPageNumberRange> obj = new Task<clsFinalPageNumberRange>();


//                return objPageNum;
//            });


//        }



//        public async Task ProcessMultipleNoParam()
//        {
//            //APIWrapper wrapper = new APIWrapper("RACKEM", "Fit to Concept", "Fit to Concept", "OVERALL", "Fit to Concept", "2022-HCPS");
//            //APIWrapper wrapper3 = new APIWrapper("DESIGNATION_REDO", "Potential For Error - Bar", "Potential For Error - Bar", "OVERALL", "", "");
//            //APIWrapper wrapper4 = new APIWrapper("EDAT_2", "Personal Preferences", "Personal Preferences", "OVERALL", "", "");
//            //APIWrapper wrapper5 = new APIWrapper("SOLITAIRE", "Likeability", "01 Untrue", "OVERALL", "", "");
//            //APIWrapper wrapper6 = new APIWrapper("MEADOW", "Memorability", "Memorability", "OVERALL", "", "");
//            //APIWrapper wrapper7 = new APIWrapper("TRIBUTE_C2", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
//            //APIWrapper wrapper8 = new APIWrapper("HABITABLE", "Overall Impressions", "Overall Impressions", "Canada Overall", "", "");
//            //APIWrapper wrapper9 = new APIWrapper("ZIPLOCK", "Suffix", "Modifier Confusion", "U.S. Overall", "", "");
//            //APIWrapper wrapper10 = new APIWrapper("ZIPLOCK", "PromotionalReview", "Modifier Confusion", "U.S. Overall", "", "");
//            //APIWrapper wrapper11 = new APIWrapper("MEADOW", "Ease Of Pronounciation", "Att 1", "OVERALL", "", "");
//            //APIWrapper wrapper12 = new APIWrapper("MEADOW", "Ease Of Spelling", "Att 1", "OVERALL", "", "");
//            //APIWrapper wrapper13 = new APIWrapper("DOMINOS", "03 exagg", "03 exagg", "OVERALL", "", "");
//            //APIWrapper wrapper14 = new APIWrapper("DORAEMON", "Innovation", "Innovation", "OVERALL", "", "");
//            //APIWrapper wrapper15 = new APIWrapper("ATEMPORAL", "Modifier", "Modifier Meaning (Aided)", "Canada Overall", "", "");
//            //APIWrapper wrapper16 = new APIWrapper("QUEENSLAND", "Written Understanding - Bar", "Written Understanding - Bar", "Canada and Europe Overall", "", "");

//            //APIWrapper wrapper17 = new APIWrapper("RACKEM", "Attribute 1", "Attribute 1", "OVERALL", "", "");
//            //APIWrapper wrapper18 = new APIWrapper("RACKEM", "Attribute 2", "Attribute 2", "OVERALL", "", "");
//            //APIWrapper wrapper19 = new APIWrapper("RACKEM", "Attribute 3", "Attribute 3", "OVERALL", "", "");
//            //APIWrapper wrapper20 = new APIWrapper("RACKEM", "01 Untrue", "01 Untrue", "OVERALL", "", "");
//            //APIWrapper wrapper21 = new APIWrapper("RACKEM", "02 Misleading", "02 Misleading", "OVERALL", "", "");
//            //APIWrapper wrapper22 = new APIWrapper("RACKEM", "03 Exagg", "03 Exagg", "OVERALL", "", "");
//            //APIWrapper wrapper24 = new APIWrapper("RACKEM", "Memorability", "Memorability", "OVERALL", "", "");
//            //APIWrapper wrapper25 = new APIWrapper("RACKEM", "Overall Impressions", "Overall Impressions", "OVERALL", "", "");
//            //APIWrapper wrapper26 = new APIWrapper("RACKEM", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
//            //APIWrapper wrapper27 = new APIWrapper("RACKEM", "Written Understanding - Bar", "Written Understanding - Bar", "Overall", "", "");
//            ////APIWrapper wrapper28 = new APIWrapper("RACKEM", "Potential For Error - Bar", "Potential For Error - Bar", "Overall", "", "");
//            //APIWrapper wrapper29 = new APIWrapper("VGT_BND", "Preference Ranking", "Preference Ranking", "Overall", "", "");
//            //APIWrapper wrapper30 = new APIWrapper("RACKEM", "Attribute evaluation Aggregate", "Attribute evaluation Aggregate", "Overall", "", "");
//            //APIWrapper wrapper31 = new APIWrapper("VAYNSMR_HEME", "Initial Recall", "Initial Recall", "Canada and EU Medical Professionals", "", "");
//            //APIWrapper wrapper32 = new APIWrapper("ALL_FAME", "Chemical Structure Appropriateness", "", "Overall", "", "");

//            //APIWrapper wrapper33 = new APIWrapper("COPEMVIBO", "Likeability Rationale", "Likeability Rationale", "Overall", "", "");


//            Task[] taskArr = new Task[]
//            {
//                //Task.Run(() => wrapper.Process()),
//                //Task.Run(() => wrapper3.Process()),
//                //Task.Run(() => wrapper4.Process()),
//                //Task.Run(() => wrapper5.Process()),
//                //Task.Run(() => wrapper6.Process()),
//                //Task.Run(() => wrapper7.Process()),
//                //Task.Run(() => wrapper8.Process()),
//                //Task.Run(() => wrapper9.Process()),
//                //Task.Run(() => wrapper10.Process()),
//                //Task.Run(() => wrapper11.Process()),
//                //Task.Run(() => wrapper12.Process()),
//                //Task.Run(() => wrapper13.Process()),
//                //Task.Run(() => wrapper14.Process()),
//                //Task.Run(() => wrapper15.Process()),
//                //Task.Run(() => wrapper16.Process()),

//                // Task.Run(() => wrapper17.Process()),
//                // Task.Run(() => wrapper18.Process()),
//                // Task.Run(() => wrapper19.Process()),
//                // Task.Run(() => wrapper20.Process()),
//                // Task.Run(() => wrapper21.Process()),
//                // Task.Run(() => wrapper22.Process()),
//                // Task.Run(() => wrapper24.Process()),
//                // Task.Run(() => wrapper25.Process()),
//                // Task.Run(() => wrapper26.Process()),
//                // Task.Run(() => wrapper27.Process()),
//                //Task.Run(() => wrapper28.Process()),

//                //Task.Run(() => wrapper29.Process()),
//                //Task.Run(() => wrapper30.Process()),
//                //Task.Run(() => wrapper31.Process()),
//                //Task.Run(() => wrapper32.Process()),
//                //Task.Run(() => wrapper33.Process())
//            };

//            await Task.WhenAll(taskArr);
//        }

//        public async Task Process()
//        {

//            //get the pageGroup

//            this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");


//            string sourcePath = "";
//            var apiCall = await CallAPI(project, template, group, breakdown, HistoricalMeanType, HistoricalMeanDescription);

//            DLLClass dLLClass = new DLLClass();
//            XDocument xDoc = XDocument.Parse(apiCall);

//            var data = xDoc.Root.Value;

//            var jsonSettings = new JsonSerializerSettings
//            {
//                NullValueHandling = NullValueHandling.Ignore,
//                MissingMemberHandling = MissingMemberHandling.Ignore
//            };


//            if (data.Length == 0)
//            {
//                Console.WriteLine("No data present in api");
//                return;
//            }

//            if (data == null || data.ToString() == "[]")
//            {
//                sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                notAvailable(CreateTargetPath(sourcePath), template, breakdown);
//            }


//            if (data.Length != 0)
//            {
//                switch (this.templateType)
//                {
//                    case "Fit to Concept":


//                        List<DLLTemplate.FitToConceptModel> fitToConceptData = new List<DLLTemplate.FitToConceptModel>();

//                        fitToConceptData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


//                        if (fitToConceptData.Count == 0)
//                        {
//                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            //CreateTargetPath(sourcePath);

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept" + fitToConceptData.Count + ".pptx";
//                            dLLClass.FitToConceptMethod(CreateTargetPath(sourcePath), fitToConceptData, HistoricalMeanType, HistoricalMeanDescription);
//                        }
//                        break;

//                    case "Attribute 1":
//                    case "Attribute 2":
//                    case "Attribute 3":
//                    case "Att 1":
//                    case "Att 2":
//                    case "Att 3":
//                    case "Attribute Evaluation":



//                        List<DLLTemplate.FitToConceptModel> Att1Data = new List<DLLTemplate.FitToConceptModel>();

//                        Att1Data = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

//                        if (Att1Data.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                            //update the notavailable 


//                        }
//                        else
//                        {


//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation" + Att1Data.Count + ".pptx";
//                            dLLClass.AttributeMethod(CreateTargetPath(sourcePath), Att1Data, group, HistoricalMeanType, HistoricalMeanDescription, project);


//                        }
//                        break;

//                    case "Attribute evaluation Aggregate":
//                    case "Attribute Evaluation Aggregate":

//                        List<DLLTemplate.FitToConceptModel> attEvalData = new List<DLLTemplate.FitToConceptModel>();

//                        attEvalData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


//                        if (attEvalData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluationAggregate" + attEvalData.Count + ".pptx";
//                            dLLClass.AttributeMethodForAttributeEvalAggreg(CreateTargetPath(sourcePath), attEvalData);
//                        }
//                        break;



//                    case "BRANDEX SUFFIX STRATEGIC":
//                        List<BrandexSuffixStrategicModel> brandexSuffixStrategicData = new List<BrandexSuffixStrategicModel>();
//                        brandexSuffixStrategicData = JsonConvert.DeserializeObject<List<BrandexSuffixStrategicModel>>(data);

//                        if (brandexSuffixStrategicData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            List<BrandexSuffixStrategicShortModel> brandexSuffixStrategicShortModel = new List<BrandexSuffixStrategicShortModel>();

//                            double dblAverage1Max = 0.0;
//                            double dblAverage2Max = 0.0;

//                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
//                            {
//                                var dblAverage1MaxValue = brandexSuffixStrategicData[i].dblAveragePage1;
//                                if (dblAverage1MaxValue == 0)
//                                {
//                                    dblAverage1Max = 0.0;
//                                }
//                                if (dblAverage1MaxValue > dblAverage1Max)
//                                {
//                                    dblAverage1Max = dblAverage1MaxValue;
//                                }

//                                var dblAverage2MaxValue = brandexSuffixStrategicData[i].dblAveragePage2;
//                                if (dblAverage2MaxValue == 0)
//                                {
//                                    dblAverage2Max = 0;
//                                }
//                                if (dblAverage2MaxValue > dblAverage2Max)
//                                {
//                                    dblAverage2Max = dblAverage2MaxValue;
//                                }
//                            }

//                            // scaling factor 
//                            double scalingFactorForStrategicDistinctiveness = 1.0;

//                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
//                            {
//                                // for the table
//                                BrandexSuffixStrategicShortModel brandexSuffixStrategicModelData = new BrandexSuffixStrategicShortModel();

//                                var dataForBrandexSuffix = brandexSuffixStrategicData[i];

//                                double averagePage1WeightedValue = 0.0;
//                                double averagePage2WeightedValue = 0.0;

//                                if (dblAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValue =
//                                        (dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * dataForBrandexSuffix.dblPage1Weight;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValue = 0;
//                                }

//                                if (dblAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValue =
//                                       (dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * dataForBrandexSuffix.dblPage2Weight;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValue = 0;
//                                }

//                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue) * scalingFactorForStrategicDistinctiveness;

//                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;

//                                brandexSuffixStrategicModelData.dblAveragePage1 = dataForBrandexSuffix.dblAveragePage1;
//                                brandexSuffixStrategicModelData.dblPage1Weight = dataForBrandexSuffix.dblPage1Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

//                                brandexSuffixStrategicModelData.dblAveragePage2 = dataForBrandexSuffix.dblAveragePage2;
//                                brandexSuffixStrategicModelData.dblPage2Weight = dataForBrandexSuffix.dblPage2Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

//                                brandexSuffixStrategicModelData.dblAveragePage3 = dataForBrandexSuffix.dblAveragePage3;
//                                brandexSuffixStrategicModelData.dblPage3Weight = dataForBrandexSuffix.dblPage3Weight;

//                                brandexSuffixStrategicModelData.dblAveragePage4 = dataForBrandexSuffix.dblAveragePage4;
//                                brandexSuffixStrategicModelData.dblPage4Weight = dataForBrandexSuffix.dblPage4Weight;

//                                brandexSuffixStrategicModelData.dblAveragePage5 = dataForBrandexSuffix.dblAveragePage5;
//                                brandexSuffixStrategicModelData.dblPage5Weight = dataForBrandexSuffix.dblPage5Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage5Weighted = dataForBrandexSuffix.dblAveragePage5;

//                                brandexSuffixStrategicModelData.dblAveragePage6 = dataForBrandexSuffix.dblAveragePage6;
//                                brandexSuffixStrategicModelData.dblPage6Weight = dataForBrandexSuffix.dblPage6Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage6Weighted = dataForBrandexSuffix.dblAveragePage6Weighted;

//                                brandexSuffixStrategicModelData.dblAveragePage7 = dataForBrandexSuffix.dblAveragePage7;
//                                brandexSuffixStrategicModelData.dblPage7Weight = dataForBrandexSuffix.dblPage7Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage7Weighted = dataForBrandexSuffix.dblAveragePage7Weighted;

//                                brandexSuffixStrategicModelData.dblAveragePage8 = dataForBrandexSuffix.dblAveragePage8;
//                                brandexSuffixStrategicModelData.dblPage8Weight = dataForBrandexSuffix.dblPage8Weight;
//                                brandexSuffixStrategicModelData.dblAveragePage8Weighted = dataForBrandexSuffix.dblAveragePage8Weighted;

//                                brandexSuffixStrategicModelData.dblIndex = indexSum;
//                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
//                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
//                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
//                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
//                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

//                                // for chart
//                                int memorabilityStrategicWeightPage = 50;
//                                int personalPreferenceStrategicWeightPage = 50;

//                                double averagePage1WeightedValueForChart = 0.0;
//                                double averagePage2WeightedValueForChart = 0.0;

//                                if (dblAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForChart = 0;
//                                }

//                                if (dblAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForChart = 0;
//                                }

//                                double indexSumForChart = averagePage1WeightedValueForChart +
//                                                          averagePage2WeightedValueForChart;

//                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;
//                                brandexSuffixStrategicModelData.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;
//                                brandexSuffixStrategicModelData.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

//                                brandexSuffixStrategicModelData.dblIndexForChart = indexSumForChart;

//                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
//                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
//                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
//                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
//                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

//                                brandexSuffixStrategicShortModel.Add(brandexSuffixStrategicModelData);
//                            }

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSuffixStrategicTemplateModified" + brandexSuffixStrategicShortModel.Count + ".pptx";
//                            dLLClass.BrandexSuffixStrategicMethod(CreateTargetPath(sourcePath), brandexSuffixStrategicShortModel);
//                        }
//                        break;





//                    //case "Initial Recall":
//                    //    List<DLLTemplate.FitToConceptNewModel> initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();
//                    //    initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data);
//                    //    if (initialRecallData.Count == 0)
//                    //    {
//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                    //    }
//                    //    else
//                    //    {
//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + initialRecallData.Count + ".pptx";
//                    //        dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), initialRecallData, this.HistoricalMeanType, this.HistoricalMeanDescription);
//                    //    }
//                    //    break;





//                    case "Brandex Strategic Distinctiveness":

//                        List<BrandexStrategicDistinctivenessModel> brandexStrategicDistinctivenessesData = new List<BrandexStrategicDistinctivenessModel>();
//                        brandexStrategicDistinctivenessesData = JsonConvert.DeserializeObject<List<BrandexStrategicDistinctivenessModel>>(data);

//                        if (brandexStrategicDistinctivenessesData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            List<BrandexStrategicDistinctivenessShortModel> brandexStrategicDistinctivenessesShortData = new List<BrandexStrategicDistinctivenessShortModel>();

//                            double dblAverage1Max = 0.0;
//                            double dblAverage2Max = 0.0;
//                            double dblAverage3Max = 0.0;
//                            double dblAverage4Max = 0.0;

//                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
//                            {
//                                var dblAverage1MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage1;
//                                if (dblAverage1MaxValue == 0)
//                                {
//                                    dblAverage1Max = 0.0;
//                                }
//                                if (dblAverage1MaxValue > dblAverage1Max)
//                                {
//                                    dblAverage1Max = dblAverage1MaxValue;
//                                }

//                                var dblAverage2MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage2;
//                                if (dblAverage2MaxValue == 0)
//                                {
//                                    dblAverage2Max = 0;
//                                }
//                                if (dblAverage2MaxValue > dblAverage2Max)
//                                {
//                                    dblAverage2Max = dblAverage2MaxValue;
//                                }

//                                var dblAverage3MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage3;
//                                if (dblAverage3MaxValue == 0)
//                                {
//                                    dblAverage3Max = 0;
//                                }
//                                if (dblAverage3MaxValue > dblAverage3Max)
//                                {
//                                    dblAverage3Max = dblAverage3MaxValue;
//                                }

//                                var dblAverage4MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage4;
//                                if (dblAverage4MaxValue == 0)
//                                {
//                                    dblAverage4Max = 0;
//                                }
//                                if (dblAverage4MaxValue > dblAverage4Max)
//                                {
//                                    dblAverage4Max = dblAverage4MaxValue;
//                                }
//                            }

//                            // scaling factor 
//                            double scalingFactorForStrategicDistinctiveness = 1.00777;

//                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
//                            {
//                                // for the table
//                                BrandexStrategicDistinctivenessShortModel brandexStrategicDistinctivenessModelData = new BrandexStrategicDistinctivenessShortModel();

//                                var dataForDistinctivessModel = brandexStrategicDistinctivenessesData[i];

//                                double averagePage1WeightedValue = 0.0;
//                                double averagePage2WeightedValue = 0.0;
//                                double averagePage3WeightedValue = 0.0;
//                                double averagePage4WeightedValue = 0.0;

//                                if (dblAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValue =
//                                        (dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * dataForDistinctivessModel.dblPage1Weight;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValue = 0;
//                                }

//                                if (dblAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValue =
//                                       (dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * dataForDistinctivessModel.dblPage2Weight;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValue = 0;
//                                }

//                                if (dblAverage3Max > 0)
//                                {
//                                    averagePage3WeightedValue =
//                                        (dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * dataForDistinctivessModel.dblPage3Weight;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValue = 0;
//                                }

//                                if (dblAverage4Max > 0)
//                                {
//                                    averagePage4WeightedValue =
//                                        (dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * dataForDistinctivessModel.dblPage4Weight;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValue = 0;
//                                }

//                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
//                                             averagePage4WeightedValue) * scalingFactorForStrategicDistinctiveness;

//                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage1 = dataForDistinctivessModel.dblAveragePage1;
//                                brandexStrategicDistinctivenessModelData.dblPage1Weight = dataForDistinctivessModel.dblPage1Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage2 = dataForDistinctivessModel.dblAveragePage2;
//                                brandexStrategicDistinctivenessModelData.dblPage2Weight = dataForDistinctivessModel.dblPage2Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage3 = dataForDistinctivessModel.dblAveragePage3;
//                                brandexStrategicDistinctivenessModelData.dblPage3Weight = dataForDistinctivessModel.dblPage3Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage4 = dataForDistinctivessModel.dblAveragePage4;
//                                brandexStrategicDistinctivenessModelData.dblPage4Weight = dataForDistinctivessModel.dblPage4Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage5 = dataForDistinctivessModel.dblAveragePage5;
//                                brandexStrategicDistinctivenessModelData.dblPage5Weight = dataForDistinctivessModel.dblPage5Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage5Weighted = dataForDistinctivessModel.dblAveragePage5;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage6 = dataForDistinctivessModel.dblAveragePage6;
//                                brandexStrategicDistinctivenessModelData.dblPage6Weight = dataForDistinctivessModel.dblPage6Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage6Weighted = dataForDistinctivessModel.dblAveragePage6Weighted;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage7 = dataForDistinctivessModel.dblAveragePage7;
//                                brandexStrategicDistinctivenessModelData.dblPage7Weight = dataForDistinctivessModel.dblPage7Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage7Weighted = dataForDistinctivessModel.dblAveragePage7Weighted;

//                                brandexStrategicDistinctivenessModelData.dblAveragePage8 = dataForDistinctivessModel.dblAveragePage8;
//                                brandexStrategicDistinctivenessModelData.dblPage8Weight = dataForDistinctivessModel.dblPage8Weight;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage8Weighted = dataForDistinctivessModel.dblAveragePage8Weighted;

//                                brandexStrategicDistinctivenessModelData.dblIndex = indexSum;
//                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
//                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
//                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
//                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
//                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

//                                // for chart distinctiveness

//                                int fitToConceptStrategicWeightPage = 40;
//                                int memorabilityStrategicWeightPage = 15;
//                                int personalPreferenceStrategicWeightPage = 15;
//                                int attributeEvaluationStrategicWeightPage = 30;

//                                double averagePage1WeightedValueForChart = 0.0;
//                                double averagePage2WeightedValueForChart = 0.0;
//                                double averagePage3WeightedValueForChart = 0.0;
//                                double averagePage4WeightedValueForChart = 0.0;

//                                if (dblAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForChart = 0;
//                                }

//                                if (dblAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForChart = 0;
//                                }

//                                if (dblAverage3Max > 0)
//                                {
//                                    averagePage3WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValueForChart = 0;
//                                }

//                                if (dblAverage4Max > 0)
//                                {
//                                    averagePage4WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValueForChart = 0;
//                                }

//                                double indexSumForChartDistinctiveness = averagePage1WeightedValueForChart +
//                                                          averagePage2WeightedValueForChart +
//                                                          averagePage3WeightedValueForChart +
//                                                          averagePage4WeightedValueForChart;

//                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForDistinctivenessChart = averagePage1WeightedValueForChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForDistinctivenessChart = averagePage2WeightedValueForChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForDistinctivenessChart = averagePage3WeightedValueForChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForDistinctivenessChart = averagePage4WeightedValueForChart;

//                                brandexStrategicDistinctivenessModelData.dblIndexForMarketingChart = indexSumForChartDistinctiveness;

//                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
//                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
//                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
//                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
//                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

//                                // for chart marketing

//                                int fitToConceptDistinctivenessWeightPage = 10;
//                                int memorabilityDistinctivenssWeightPage = 30;
//                                int personalPreferenceDistinctivenessWeightPage = 40;
//                                int attributeEvaluationDistinctivenessWeightPage = 20;

//                                double scalingFactorForMarketingChart = 1.00327;

//                                double averagePage1WeightedValueForMarketingChart = 0.0;
//                                double averagePage2WeightedValueForMarketingChart = 0.0;
//                                double averagePage3WeightedValueForMarketingChart = 0.0;
//                                double averagePage4WeightedValueForMarketingChart = 0.0;

//                                if (dblAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptDistinctivenessWeightPage) * scalingFactorForMarketingChart;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForMarketingChart = 0;
//                                }

//                                if (dblAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityDistinctivenssWeightPage) * scalingFactorForMarketingChart;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForMarketingChart = 0;
//                                }

//                                if (dblAverage3Max > 0)
//                                {
//                                    averagePage3WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceDistinctivenessWeightPage) * scalingFactorForMarketingChart;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValueForMarketingChart = 0;
//                                }

//                                if (dblAverage4Max > 0)
//                                {
//                                    averagePage4WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationDistinctivenessWeightPage) * scalingFactorForMarketingChart;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValueForMarketingChart = 0;
//                                }

//                                double indexSumForChartMarketing = averagePage1WeightedValueForMarketingChart +
//                                                          averagePage2WeightedValueForMarketingChart +
//                                                          averagePage3WeightedValueForMarketingChart +
//                                                          averagePage4WeightedValueForMarketingChart;

//                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage1WeightedForMarketingChart = averagePage1WeightedValueForMarketingChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage2WeightedForMarketingChart = averagePage2WeightedValueForMarketingChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage3WeightedForMarketingChart = averagePage3WeightedValueForMarketingChart;
//                                brandexStrategicDistinctivenessModelData.dblAveragePage4WeightedForMarketingChart = averagePage4WeightedValueForMarketingChart;

//                                brandexStrategicDistinctivenessModelData.dblIndexForDistinctivenessChart = indexSumForChartMarketing;
//                                brandexStrategicDistinctivenessModelData.strDSIScore = dataForDistinctivessModel.strDSIScore;
//                                brandexStrategicDistinctivenessModelData.intRed = dataForDistinctivessModel.intRed;
//                                brandexStrategicDistinctivenessModelData.intGreen = dataForDistinctivessModel.intGreen;
//                                brandexStrategicDistinctivenessModelData.intBlue = dataForDistinctivessModel.intBlue;
//                                brandexStrategicDistinctivenessModelData.boolBold = dataForDistinctivessModel.boolBold;

//                                brandexStrategicDistinctivenessesShortData.Add(brandexStrategicDistinctivenessModelData);
//                            }

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexStrategicDistinctiveness" + brandexStrategicDistinctivenessesData.Count + ".pptx";
//                            dLLClass.BrandexStrategicDistinctivenessMethod(CreateTargetPath(sourcePath), brandexStrategicDistinctivenessesShortData);
//                        }




//                        break;


//                    //case "Brandex Safety":
//                    //    List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
//                    //    brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

//                    //    List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

//                    //    if (brandexSafetyData.Count == 0)
//                    //    {
//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                    //    }
//                    //    else
//                    //    {
//                    //        double average1Max = 0.0000;
//                    //        double average2Max = 0.0000;
//                    //        double average3Max = 0.0;
//                    //        double average4Max = 0.0000;
//                    //        double average5Max = 0.0000;

//                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
//                    //        {
//                    //            var dblAverage1max = brandexSafetyData[i].dblAveragePage1;
//                    //            if (dblAverage1max == 0.0)
//                    //            {
//                    //                average1Max = 0;
//                    //            }
//                    //            if (average1Max < dblAverage1max)
//                    //            {
//                    //                average1Max = dblAverage1max;
//                    //            }

//                    //            var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
//                    //            if (dblAverage2max == 0.0)
//                    //            {
//                    //                average2Max = 0;
//                    //            }
//                    //            if (average2Max < dblAverage2max)
//                    //            {
//                    //                average2Max = dblAverage2max;
//                    //            }

//                    //            var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
//                    //            if (dblAverage3max == 0.0)
//                    //            {
//                    //                average3Max = 0;
//                    //            }
//                    //            if (average3Max < dblAverage3max)
//                    //            {
//                    //                average3Max = dblAverage3max;
//                    //            }

//                    //            var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
//                    //            if (dblAverage4max == 0.0)
//                    //            {
//                    //                average4Max = 0;
//                    //            }
//                    //            if (average4Max < dblAverage4max)
//                    //            {
//                    //                average4Max = dblAverage4max;
//                    //            }

//                    //            var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
//                    //            if (dblAverage5max == 0.0)
//                    //            {
//                    //                average5Max = 0;
//                    //            }
//                    //            if (average5Max < dblAverage5max)
//                    //            {
//                    //                average5Max = dblAverage5max;
//                    //            }
//                    //        }

//                    //        double scalingFactor = 0.750610000;

//                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
//                    //        {
//                    //            //for the table
//                    //            BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

//                    //            var dataEl = brandexSafetyData[i];

//                    //            double averagePage1WeightedValue = 0.0;
//                    //            double averagePage2WeightedValue = 0.0;
//                    //            double averagePage3WeightedValue = 0.0;
//                    //            double averagePage4WeightedValue = 0.0;
//                    //            double averagePage5WeightedValue = 0.0;

//                    //            if (average1Max > 0)
//                    //            {
//                    //                averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage1WeightedValue = 0;
//                    //            }

//                    //            if (average2Max > 0)
//                    //            {
//                    //                averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage2WeightedValue = 0;
//                    //            }

//                    //            if (average3Max > 0)
//                    //            {
//                    //                averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage3WeightedValue = 0;
//                    //            }

//                    //            if (average4Max > 0)
//                    //            {
//                    //                averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage4WeightedValue = 0;
//                    //            }

//                    //            if (average5Max > 0)
//                    //            {
//                    //                averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage5WeightedValue = 0;
//                    //            }


//                    //            //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
//                    //            //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
//                    //            //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
//                    //            // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
//                    //            //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

//                    //            double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
//                    //                              averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;

//                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                    //            brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
//                    //            brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
//                    //            brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

//                    //            brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
//                    //            brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
//                    //            brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

//                    //            brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
//                    //            brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
//                    //            brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

//                    //            brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
//                    //            brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
//                    //            brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

//                    //            brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
//                    //            brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
//                    //            brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

//                    //            brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
//                    //            brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
//                    //            brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

//                    //            brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
//                    //            brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
//                    //            brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

//                    //            brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
//                    //            brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
//                    //            brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

//                    //            brandexSafetyShortModel.dblIndex = indexSum;
//                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
//                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                    //            //for the chart - 

//                    //            double averagePage1WeightedValueForChart = 0.0;
//                    //            double averagePage2WeightedValueForChart = 0.0;
//                    //            double averagePage3WeightedValueForChart = 0.0;
//                    //            double averagePage4WeightedValueForChart = 0.0;
//                    //            double averagePage5WeightedValueForChart = 0.0;

//                    //            if (average1Max > 0)
//                    //            {
//                    //                averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage1WeightedValueForChart = 0;
//                    //            }

//                    //            if (average2Max > 0)
//                    //            {
//                    //                averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage2WeightedValueForChart = 0;
//                    //            }

//                    //            if (average3Max > 0)
//                    //            {
//                    //                averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage3WeightedValueForChart = 0;
//                    //            }

//                    //            if (average4Max > 0)
//                    //            {
//                    //                averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage4WeightedValueForChart = 0;
//                    //            }

//                    //            if (average5Max > 0)
//                    //            {
//                    //                averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
//                    //            }
//                    //            else
//                    //            {
//                    //                averagePage5WeightedValueForChart = 0;
//                    //            }

//                    //            //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
//                    //            //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
//                    //            //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
//                    //            //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
//                    //            //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

//                    //            double indexSumForChart = averagePage1WeightedValueForChart +
//                    //                                      averagePage2WeightedValueForChart +
//                    //                                      averagePage3WeightedValueForChart +
//                    //                                      averagePage4WeightedValueForChart +
//                    //                                      averagePage5WeightedValueForChart;

//                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                    //            brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

//                    //            brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

//                    //            brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

//                    //            brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

//                    //            brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
//                    //            ;
//                    //            brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
//                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
//                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                    //            brandexData.Add(brandexSafetyShortModel);
//                    //        }

//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
//                    //        dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
//                    //    }
//                    //    break;


//                    case "Brandex Safety":
//                        List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
//                        brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

//                        List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

//                        if (brandexSafetyData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            double average1Max = 0.0000;
//                            double average2Max = 0.0000;
//                            double average3Max = 0.0;
//                            double average4Max = 0.0000;
//                            double average5Max = 0.0000;

//                            for (int i = 0; i < brandexSafetyData.Count; i++)
//                            {
//                                var dblAverage1max = brandexSafetyData[i]?.dblAveragePage1;
//                                if (dblAverage1max == 0.0 || dblAverage1max == null)
//                                {
//                                    average1Max = 0;
//                                }
//                                if (average1Max < dblAverage1max)
//                                {
//                                    average1Max = (double)dblAverage1max;
//                                }

//                                var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
//                                if (dblAverage2max == 0.0)
//                                {
//                                    average2Max = 0;
//                                }
//                                if (average2Max < dblAverage2max)
//                                {
//                                    average2Max = (double)dblAverage2max;
//                                }

//                                var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
//                                if (dblAverage3max == 0.0)
//                                {
//                                    average3Max = 0;
//                                }
//                                if (average3Max < dblAverage3max)
//                                {
//                                    average3Max = (double)dblAverage3max;
//                                }

//                                var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
//                                if (dblAverage4max == 0.0)
//                                {
//                                    average4Max = 0;
//                                }
//                                if (average4Max < dblAverage4max)
//                                {
//                                    average4Max = (double)dblAverage4max;
//                                }

//                                var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
//                                if (dblAverage5max == 0.0)
//                                {
//                                    average5Max = 0;
//                                }
//                                if (average5Max < dblAverage5max)
//                                {
//                                    average5Max = (double)dblAverage5max;
//                                }
//                            }

//                            double scalingFactor = 0.750610000;

//                            for (int i = 0; i < brandexSafetyData.Count; i++)
//                            {
//                                //for the table
//                                BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

//                                var dataEl = brandexSafetyData[i];

//                                double averagePage1WeightedValue = 0.0;
//                                double averagePage2WeightedValue = 0.0;
//                                double averagePage3WeightedValue = 0.0;
//                                double averagePage4WeightedValue = 0.0;
//                                double averagePage5WeightedValue = 0.0;

//                                if (average1Max > 0)
//                                {
//                                    // averagePage1WeightedValue =(dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) ;

//                                    averagePage1WeightedValue = (double)((double)(dataEl.dblAveragePage1 / average1Max) * (double)(dataEl.dblPage1Weight));
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValue = 0;
//                                }

//                                if (average2Max > 0)
//                                {
//                                    // averagePage2WeightedValue = (double)(dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;

//                                    averagePage2WeightedValue = (double)((double)(dataEl.dblAveragePage2 / average2Max) * (double)(dataEl.dblPage2Weight));
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValue = 0;
//                                }

//                                if (average3Max > 0)
//                                {
//                                    averagePage3WeightedValue = (double)((double)(dataEl.dblAveragePage3 / average3Max) * (double)(dataEl.dblPage3Weight));
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValue = 0;
//                                }

//                                if (average4Max > 0)
//                                {
//                                    //averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;

//                                    averagePage4WeightedValue = (double)((double)(dataEl.dblAveragePage4 / average4Max) * (double)(dataEl.dblPage4Weight));
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValue = 0;
//                                }

//                                if (average5Max > 0)
//                                {
//                                    // averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
//                                    averagePage5WeightedValue = (double)((double)(dataEl.dblAveragePage5 / average5Max) * (double)(dataEl.dblPage5Weight));


//                                }
//                                else
//                                {
//                                    averagePage5WeightedValue = 0;
//                                }


//                                //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
//                                //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
//                                //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
//                                // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
//                                //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

//                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
//                                                  averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;

//                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1 ?? 0.0;
//                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2 ?? 0.0;
//                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3 ?? 0.0;
//                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4 ?? 0.0;
//                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5 ?? 0.0;
//                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6 ?? 0.0;
//                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

//                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7 ?? 0.0;
//                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

//                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8 ?? 0.0;
//                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight ?? 0.0;
//                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

//                                brandexSafetyShortModel.dblIndex = indexSum;
//                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                                brandexSafetyShortModel.intRed = dataEl.intRed;
//                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                                //for the chart - 

//                                double averagePage1WeightedValueForChart = 0.0;
//                                double averagePage2WeightedValueForChart = 0.0;
//                                double averagePage3WeightedValueForChart = 0.0;
//                                double averagePage4WeightedValueForChart = 0.0;
//                                double averagePage5WeightedValueForChart = 0.0;

//                                if (average1Max > 0)
//                                {
//                                    averagePage1WeightedValueForChart = (double)((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForChart = 0;
//                                }

//                                if (average2Max > 0)
//                                {
//                                    averagePage2WeightedValueForChart = (double)((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForChart = 0;
//                                }

//                                if (average3Max > 0)
//                                {
//                                    averagePage3WeightedValueForChart = (double)((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValueForChart = 0;
//                                }

//                                if (average4Max > 0)
//                                {
//                                    averagePage4WeightedValueForChart = (double)((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValueForChart = 0;
//                                }

//                                if (average5Max > 0)
//                                {
//                                    averagePage5WeightedValueForChart = (double)((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage5WeightedValueForChart = 0;
//                                }

//                                //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
//                                //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
//                                //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
//                                //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
//                                //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

//                                double indexSumForChart = averagePage1WeightedValueForChart +
//                                                          averagePage2WeightedValueForChart +
//                                                          averagePage3WeightedValueForChart +
//                                                          averagePage4WeightedValueForChart +
//                                                          averagePage5WeightedValueForChart;

//                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                                brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
//                                ;
//                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
//                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                                brandexSafetyShortModel.intRed = dataEl.intRed;
//                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                                brandexData.Add(brandexSafetyShortModel);
//                            }

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
//                            dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
//                        }
//                        break;

//                    case "Potential For Error - Bar":
//                    case "Potential For Error":


//                        List<DLLTemplate.FitToConceptModel> potentialForError = new List<DLLTemplate.FitToConceptModel>();

//                        potentialForError = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


//                        if (potentialForError.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PotentialForError-Bar" + potentialForError.Count + ".pptx";
//                            dLLClass.PotentialForErrorBarMethod(CreateTargetPath(sourcePath), potentialForError);
//                        }
//                        break;

//                    case "Personal Preferences":


//                        List<DLLTemplate.FitToConceptModel> personalPref = new List<DLLTemplate.FitToConceptModel>();


//                        personalPref = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

//                        if (personalPref.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            //lok needs a fix RADIANT : error (Personal Preference testname null)

//                            // notAvailable(this.template);
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PersonalPreferences" + personalPref.Count + ".pptx";
//                            dLLClass.PersonalPreferencesMethod(CreateTargetPath(sourcePath), personalPref, HistoricalMeanType, HistoricalMeanDescription);
//                        }
//                        break;

//                    case "Likeability":


//                        List<DLLTemplate.OverallImpressionModel> likeability = new List<DLLTemplate.OverallImpressionModel>();
//                        likeability = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

//                        if (likeability.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Likeability" + likeability.Count + ".pptx";
//                            dLLClass.LikeabilityMethod(CreateTargetPath(sourcePath), likeability);
//                        }
//                        break;



//                    case "Likeability Rationale":
//                        List<LikeabilityRationaleModel> likeabilyRationalesData = new List<LikeabilityRationaleModel>();
//                        likeabilyRationalesData = JsonConvert.DeserializeObject<List<LikeabilityRationaleModel>>(data);

//                        if (likeabilyRationalesData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationale.pptx";
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationaleOrdered.pptx";
//                            dLLClass.LikeabilityRationaleMethod(CreateTargetPath(sourcePath), likeabilyRationalesData, breakdown);
//                        }
//                        break;

//                    case "Memorability":


//                        List<DLLTemplate.FitToConceptModel> memorability = new List<DLLTemplate.FitToConceptModel>();

//                        memorability = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

//                        if (memorability.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Memorability" + memorability.Count + ".pptx";
//                            dLLClass.MemorabilityMethod(CreateTargetPath(sourcePath), memorability, HistoricalMeanType, HistoricalMeanDescription);
//                        }
//                        break;

//                    case "Verbal Understanding - Bar":
//                    case "Verbal Understanding":


//                        List<DLLTemplate.FitToConceptModel> verbalUnde = new List<DLLTemplate.FitToConceptModel>();


//                        verbalUnde = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

//                        if (verbalUnde.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\VerbalUnderstanding-Bar" + verbalUnde.Count + ".pptx";
//                            dLLClass.VerbalUnderstandingBarMethod(CreateTargetPath(sourcePath), verbalUnde);
//                        }
//                        break;

//                    case "Overall Impressions":


//                        List<DLLTemplate.OverallImpressionNewModel> overallImpression = new List<DLLTemplate.OverallImpressionNewModel>();



//                        overallImpression = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

//                        if (overallImpression.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\OverallImpressions" + overallImpression.Count + ".pptx";
//                            dLLClass.OverallImpressionsMethod(CreateTargetPath(sourcePath), overallImpression);
//                        }
//                        break;

//                    case "Suffix":
//                    case "Suffix Meaning":
//                    case "Existing Abbreviation ID":
//                    case "Existing Suffix ID":


//                        List<DLLTemplate.OverallImpressionModel> suffixOverallImpressionData = new List<DLLTemplate.OverallImpressionModel>();

//                        suffixOverallImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

//                        if (suffixOverallImpressionData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Suffix" + suffixOverallImpressionData.Count + ".pptx";
//                            dLLClass.SuffixMethod(CreateTargetPath(sourcePath), suffixOverallImpressionData, breakdown);
//                        }
//                        break;

//                    case "PromotionalReview":

//                        List<DLLTemplate.OverallImpressionModel> promotionalImpressionData = new List<DLLTemplate.OverallImpressionModel>();

//                        promotionalImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

//                        if (promotionalImpressionData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PromotionalReview" + promotionalImpressionData.Count + ".pptx";
//                            dLLClass.PromotionalReviewMethod(CreateTargetPath(sourcePath), promotionalImpressionData);
//                        }
//                        break;

//                    case "Ease Of Pronounciation":
//                    case "Ease of Pronunciation":



//                        List<DLLTemplate.FitToConceptModel> easeOfPronounicationData = new List<DLLTemplate.FitToConceptModel>();
//                        easeOfPronounicationData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

//                        if (easeOfPronounicationData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfPronunciation" + easeOfPronounicationData.Count + ".pptx";
//                            dLLClass.EaseOfPronounicationMethod(CreateTargetPath(sourcePath), easeOfPronounicationData, HistoricalMeanType, HistoricalMeanDescription);
//                        }
//                        break;

//                    case "Ease Of Spelling":

//                        List<DLLTemplate.FitToConceptNewModel> easeOfSpellingData = new List<DLLTemplate.FitToConceptNewModel>();

//                        easeOfSpellingData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

//                        if (easeOfSpellingData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfSpelling" + easeOfSpellingData.Count + ".pptx";
//                            dLLClass.EaseOfSpellingMethod(CreateTargetPath(sourcePath), easeOfSpellingData);
//                        }

//                        break;

//                    case "03 Exagg":
//                    case "01 Untrue":
//                    case "02 Misleading":
//                    case "02 Mislead":
//                    case "03 Exaggerative":
//                    case "Exaggerative-Inappropriate":


//                        List<DLLTemplate.OverallImpressionNewModel> ExaggData = new List<DLLTemplate.OverallImpressionNewModel>();
//                        ExaggData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);
//                        if (ExaggData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            Console.WriteLine(ExaggData);

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative" + ExaggData.Count + ".pptx";
//                            dLLClass.ExaggerativeMethod(CreateTargetPath(sourcePath), ExaggData, this.template);
//                        }
//                        break;


//                    case "Initial Recall":

//                        List<DLLTemplate.FitToConceptNewModel> Vaynsmr_heme_initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();

//                        Vaynsmr_heme_initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

//                        if (Vaynsmr_heme_initialRecallData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + Vaynsmr_heme_initialRecallData.Count + ".pptx";
//                            dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), Vaynsmr_heme_initialRecallData, HistoricalMeanType, HistoricalMeanDescription);
//                        }
//                        break;

//                    case "Innovation":
//                        List<DLLTemplate.InnovationModel> doraemonInnovationData = new List<DLLTemplate.InnovationModel>();
//                        doraemonInnovationData = JsonConvert.DeserializeObject<List<DLLTemplate.InnovationModel>>(data);
//                        if (doraemonInnovationData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Innovation" + doraemonInnovationData.Count + ".pptx";
//                            dLLClass.InnovationMethod(CreateTargetPath(sourcePath), doraemonInnovationData, this.HistoricalMeanType, this.HistoricalMeanDescription);
//                        }
//                        break;


//                    case "Modifier":

//                        List<DLLTemplate.OverallImpressionNewModel> atemporalModifierData = new List<DLLTemplate.OverallImpressionNewModel>();

//                        atemporalModifierData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

//                        if (atemporalModifierData.Count == 0)
//                        {

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Modifier" + atemporalModifierData.Count + ".pptx";
//                            dLLClass.ModifierMethod(CreateTargetPath(sourcePath), atemporalModifierData, group);
//                        }
//                        break;

//                    case "Written Understanding - Bar":
//                    case "Written Understanding":


//                        List<DLLTemplate.WrittenUnderstadingModel> writtenUnde = new List<DLLTemplate.WrittenUnderstadingModel>();

//                        writtenUnde = JsonConvert.DeserializeObject<List<DLLTemplate.WrittenUnderstadingModel>>(data, jsonSettings);

//                        if (writtenUnde.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\WrittenUnderstanding" + writtenUnde.Count + ".pptx";
//                            dLLClass.WrittenUnderstandingMethod(CreateTargetPath(sourcePath), writtenUnde);
//                        }
//                        break;




//                    case "Preference Ranking":
//                        List<PreferenceRankingModel> preferenceRankingData = new List<PreferenceRankingModel>();
//                        preferenceRankingData = JsonConvert.DeserializeObject<List<PreferenceRankingModel>>(data);

//                        if (preferenceRankingData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            List<int> counterList = new List<int>();

//                            int count = 0;

//                            for (int j = 0; j < preferenceRankingData.Count; j++)
//                            {
//                                var apiDataEl = preferenceRankingData[j];
//                                if (apiDataEl.intRankCount1 != 0)
//                                {
//                                    count = 1;
//                                }
//                                if (apiDataEl.intRankCount2 != 0)
//                                {
//                                    count = 2;
//                                }
//                                if (apiDataEl.intRankCount3 != 0)
//                                {
//                                    count = 3;
//                                }
//                                if (apiDataEl.intRankCount4 != 0)
//                                {
//                                    count = 4;
//                                }
//                                if (apiDataEl.intRankCount5 != 0)
//                                {
//                                    count = 5;
//                                }
//                                if (apiDataEl.intRankCount6 != 0)
//                                {
//                                    count = 6;
//                                }
//                                if (apiDataEl.intRankCount7 != 0)
//                                {
//                                    count = 7;
//                                }
//                                if (apiDataEl.intRankCount8 != 0)
//                                {
//                                    count = 8;
//                                }
//                                if (apiDataEl.intRankCount9 != 0)
//                                {
//                                    count = 9;
//                                }
//                                counterList.Add(count);
//                            }

//                            List<PreferenceRankingSortedDataModel> apiDataWithWeightedList = new List<PreferenceRankingSortedDataModel>();

//                            List<int> lstWeightedData = new List<int>();

//                            int weightedSum = 0;

//                            for (int j = 0; j < preferenceRankingData.Count; j++)
//                            {
//                                PreferenceRankingSortedDataModel preferenceRankingSortedData = new PreferenceRankingSortedDataModel();

//                                var counterVal = counterList[j];

//                                weightedSum = 0;
//                                var apiDataEl = preferenceRankingData[j];

//                                preferenceRankingSortedData.strTestName = apiDataEl.strTestName;

//                                if (apiDataEl.intRankCount1 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount1 * counterVal;
//                                    preferenceRankingSortedData.intRankCount1 = apiDataEl.intRankCount1;

//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount2 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount2 * counterVal;
//                                    preferenceRankingSortedData.intRankCount2 = apiDataEl.intRankCount2;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount3 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount3 * counterVal;
//                                    preferenceRankingSortedData.intRankCount3 = apiDataEl.intRankCount3;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount4 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount4 * counterVal;
//                                    preferenceRankingSortedData.intRankCount4 = apiDataEl.intRankCount4;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount5 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount5 * counterVal;
//                                    preferenceRankingSortedData.intRankCount5 = apiDataEl.intRankCount5;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount6 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount6 * counterVal;
//                                    preferenceRankingSortedData.intRankCount6 = apiDataEl.intRankCount6;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount7 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount7 * counterVal;
//                                    preferenceRankingSortedData.intRankCount7 = apiDataEl.intRankCount7;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount8 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount8 * counterVal;
//                                    preferenceRankingSortedData.intRankCount8 = apiDataEl.intRankCount8;
//                                    counterVal--;
//                                }
//                                if (apiDataEl.intRankCount9 != 0)
//                                {
//                                    weightedSum += apiDataEl.intRankCount9 * counterVal;
//                                    preferenceRankingSortedData.intRankCount9 = apiDataEl.intRankCount9;
//                                    counterVal--;
//                                }

//                                preferenceRankingSortedData.weightedScore = weightedSum;
//                                lstWeightedData.Add(weightedSum);
//                                apiDataWithWeightedList.Add(preferenceRankingSortedData);
//                            }

//                            var max = apiDataWithWeightedList.Max(m => m.weightedScore);

//                            if (max > 100 && max < 200)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking200Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 200 && max < 300)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking300Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 300 && max < 400)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking400Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 400 && max < 500)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking500Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 500 && max < 600)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking600Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 600 && max < 700)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking700Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 700 && max < 800)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking800Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 800 && max < 900)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking900Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                            if (max > 900 && max < 1000)
//                            {
//                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking1000Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
//                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
//                            }
//                        }

//                        break;




//                    case "Chemical Structure Appropriateness":
//                    case "Chemical Structure":
//                        List<ChemicalModel> chemicalData = new List<ChemicalModel>();
//                        chemicalData = JsonConvert.DeserializeObject<List<ChemicalModel>>(data);

//                        if (chemicalData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ChemicalStructure" + chemicalData.Count + ".pptx";
//                            dLLClass.ChemicalStructureMethod(CreateTargetPath(sourcePath), chemicalData);
//                        }
//                        break;









//                    case "Fit to Therapeutic Class":

//                        List<FitToTherapeuticClassModel> therapeuticClassData = new List<FitToTherapeuticClassModel>();
//                        therapeuticClassData = JsonConvert.DeserializeObject<List<FitToTherapeuticClassModel>>(data);

//                        if (therapeuticClassData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticClass" + therapeuticClassData.Count + ".pptx";
//                            dLLClass.FitToTherapeuticClassMethod(CreateTargetPath(sourcePath), therapeuticClassData, this.breakdown);
//                        }
//                        break;



//                    case "Associations":
//                        List<AssociationsModel> associationsData = new List<AssociationsModel>();
//                        associationsData = JsonConvert.DeserializeObject<List<AssociationsModel>>(data, jsonSettings);

//                        if (associationsData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Associations.pptx";
//                            dLLClass.AssociationsMethod(CreateTargetPath(sourcePath), associationsData, breakdown);
//                        }
//                        break;




//                    case "QTC":
//                        List<QTCModel> qtcData = new List<QTCModel>();
//                        qtcData = JsonConvert.DeserializeObject<List<QTCModel>>(data, jsonSettings);

//                        if (qtcData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTC.pptx";
//                            dLLClass.QTCMethod(CreateTargetPath(sourcePath), qtcData, breakdown);
//                        }
//                        break;


//                    case "QTCCustom":
//                        List<QTCCustomModel> qtcCustomData = new List<QTCCustomModel>();
//                        qtcCustomData = JsonConvert.DeserializeObject<List<QTCCustomModel>>(data);

//                        int counter = 0;

//                        for (int i = 0; i < qtcCustomData.Count; i++)
//                        {
//                            var dataEl = qtcCustomData[i];
//                            if (dataEl.strPageType1 != null)
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType2 != null)
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType3 != null)
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType4 != null)
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType5 != null)
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType6 != "")
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType7 != "")
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType8 != "")
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageType9 != "")
//                            {
//                                counter++;
//                            }
//                            if (dataEl.strPageName10 != "")
//                            {
//                                counter++;
//                            }
//                            break;
//                        }

//                        if (qtcCustomData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTCCustom" + (counter * 2) + ".pptx";
//                            dLLClass.QTCCustomMethod(CreateTargetPath(sourcePath), qtcCustomData, this.breakdown);
//                        }
//                        break;



//                    case "Reflective of Mechanism of Action":
//                        List<ReflectiveOfMechanismOfActionModel> reflectiveOfMechanismsData = new List<ReflectiveOfMechanismOfActionModel>();
//                        reflectiveOfMechanismsData = JsonConvert.DeserializeObject<List<ReflectiveOfMechanismOfActionModel>>(data, jsonSettings);

//                        if (reflectiveOfMechanismsData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ReflectiveofMechanismofAction" + reflectiveOfMechanismsData.Count + ".pptx";
//                            dLLClass.ReflectiveofMechanismofActionMethod(CreateTargetPath(sourcePath), reflectiveOfMechanismsData);
//                        }
//                        break;


//                    case "PhoneticTesting":
//                        List<PhoneticTestingModel> phoneticData = new List<PhoneticTestingModel>();
//                        phoneticData = JsonConvert.DeserializeObject<List<PhoneticTestingModel>>(data, jsonSettings);

//                        if (phoneticData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PhoneticTesting" + phoneticData.Count + ".pptx";
//                            dLLClass.PhoneticTestingMethod(CreateTargetPath(sourcePath), phoneticData);
//                        }
//                        break;



//                    case "Modifier Confusion":
//                        List<OverallImpressionNewModel> modifierConfusionData = new List<OverallImpressionNewModel>();
//                        modifierConfusionData = JsonConvert.DeserializeObject<List<OverallImpressionNewModel>>(data, jsonSettings);

//                        if (modifierConfusionData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ModifierConfusion" + modifierConfusionData.Count + ".pptx";
//                            dLLClass.ModifierConfusionMethod(CreateTargetPath(sourcePath), modifierConfusionData, group);
//                        }
//                        break;


//                    case "Medical Terms":
//                    case "Medical Terms Similarity":

//                        List<MedicalTermsModel> medicalData = new List<MedicalTermsModel>();
//                        medicalData = JsonConvert.DeserializeObject<List<MedicalTermsModel>>(data);

//                        if (medicalData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);

//                        }
//                        else
//                        {
//                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholder.pptx";
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholderRow.pptx";
//                            dLLClass.MedicalTermsMethod(CreateTargetPath(sourcePath), medicalData);
//                        }
//                        break;


//                    case "Non-Medical Terms":
//                    case "Non - Medical Terms Similarity":
//                    case "Non-Medical Terms Similarity":

//                        List<NonMedicalTermsModel> nonMedicalData = new List<NonMedicalTermsModel>();
//                        nonMedicalData = JsonConvert.DeserializeObject<List<NonMedicalTermsModel>>(data);

//                        if (nonMedicalData.Count == 0)
//                        {
//                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholderRow.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholder.pptx";
//                            dLLClass.NonMedicalTermsMethod(CreateTargetPath(sourcePath), nonMedicalData);
//                        }
//                        break;





//                    case "BRANDEX LOGO":
//                        List<BrandexLogoModel> brandexLogoData = new List<BrandexLogoModel>();
//                        brandexLogoData = JsonConvert.DeserializeObject<List<BrandexLogoModel>>(data);

//                        if (brandexLogoData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            List<BrandexLogoShortModel> brandexLogoShortData = new List<BrandexLogoShortModel>();

//                            double average1MaxLogo = 0.0;
//                            double average2MaxLogo = 0.0;
//                            double average3MaxLogo = 0.0;
//                            double average4MaxLogo = 0.0;
//                            double average5MaxLogo = 0.0;

//                            for (int i = 0; i < brandexLogoData.Count; i++)
//                            {
//                                var dblAverage1max = brandexLogoData[i].dblAveragePage1;
//                                if (dblAverage1max == 0.0)
//                                {
//                                    average1MaxLogo = 0;
//                                }
//                                if (average1MaxLogo < dblAverage1max)
//                                {
//                                    average1MaxLogo = dblAverage1max;
//                                }

//                                var dblAverage2max = brandexLogoData[i].dblAveragePage2;
//                                if (dblAverage2max == 0.0)
//                                {
//                                    average2MaxLogo = 0;
//                                }
//                                if (average2MaxLogo < dblAverage2max)
//                                {
//                                    average2MaxLogo = dblAverage2max;
//                                }

//                                var dblAverage3max = brandexLogoData[i].dblAveragePage3;
//                                if (dblAverage3max == 0.0)
//                                {
//                                    average3MaxLogo = 0;
//                                }
//                                if (average3MaxLogo < dblAverage3max)
//                                {
//                                    average3MaxLogo = dblAverage3max;
//                                }

//                                var dblAverage4max = brandexLogoData[i].dblAveragePage4;
//                                if (dblAverage4max == 0.0)
//                                {
//                                    average4MaxLogo = 0;
//                                }
//                                if (average4MaxLogo < dblAverage4max)
//                                {
//                                    average4MaxLogo = dblAverage4max;
//                                }

//                                var dblAverage5max = brandexLogoData[i].dblAveragePage5;
//                                if (dblAverage5max == 0.0)
//                                {
//                                    average5MaxLogo = 0;
//                                }
//                                if (average5MaxLogo < dblAverage5max)
//                                {
//                                    average5MaxLogo = dblAverage5max;
//                                }
//                            }

//                            double brandexLogoscalingFactor = 1.105380082;

//                            for (int i = 0; i < brandexLogoData.Count; i++)
//                            {
//                                //for the table
//                                BrandexLogoShortModel brandexLogoShortModelData = new BrandexLogoShortModel();

//                                var dataEl = brandexLogoData[i];

//                                double averagePage1WeightedValue = 0.0;
//                                double averagePage2WeightedValue = 0.0;
//                                double averagePage3WeightedValue = 0.0;
//                                double averagePage4WeightedValue = 0.0;
//                                double averagePage5WeightedValue = 0.0;

//                                if (average1MaxLogo > 0)
//                                {
//                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValue = 0;
//                                }

//                                if (average2MaxLogo > 0)
//                                {
//                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValue = 0;
//                                }

//                                if (average3MaxLogo > 0)
//                                {
//                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValue = 0;
//                                }

//                                if (average4MaxLogo > 0)
//                                {
//                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValue = 0;
//                                }

//                                if (average5MaxLogo > 0)
//                                {
//                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight;
//                                }
//                                else
//                                {
//                                    averagePage5WeightedValue = 0;
//                                }

//                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
//                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexLogoscalingFactor;

//                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

//                                brandexLogoShortModelData.dblAveragePage1 = dataEl.dblAveragePage1;
//                                brandexLogoShortModelData.dblPage1Weight = dataEl.dblPage1Weight;
//                                brandexLogoShortModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

//                                brandexLogoShortModelData.dblAveragePage2 = dataEl.dblAveragePage2;
//                                brandexLogoShortModelData.dblPage2Weight = dataEl.dblPage2Weight;
//                                brandexLogoShortModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

//                                brandexLogoShortModelData.dblAveragePage3 = dataEl.dblAveragePage3;
//                                brandexLogoShortModelData.dblPage3Weight = dataEl.dblPage3Weight;
//                                brandexLogoShortModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

//                                brandexLogoShortModelData.dblAveragePage4 = dataEl.dblAveragePage4;
//                                brandexLogoShortModelData.dblPage4Weight = dataEl.dblPage4Weight;
//                                brandexLogoShortModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

//                                brandexLogoShortModelData.dblAveragePage5 = dataEl.dblAveragePage5;
//                                brandexLogoShortModelData.dblPage5Weight = dataEl.dblPage5Weight;
//                                brandexLogoShortModelData.dblAveragePage5Weighted = averagePage5WeightedValue;

//                                brandexLogoShortModelData.dblAveragePage6 = dataEl.dblAveragePage6;
//                                brandexLogoShortModelData.dblPage6Weight = dataEl.dblPage6Weight;
//                                brandexLogoShortModelData.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

//                                brandexLogoShortModelData.dblAveragePage7 = dataEl.dblAveragePage7;
//                                brandexLogoShortModelData.dblPage7Weight = dataEl.dblPage7Weight;
//                                brandexLogoShortModelData.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

//                                brandexLogoShortModelData.dblAveragePage8 = dataEl.dblAveragePage8;
//                                brandexLogoShortModelData.dblPage8Weight = dataEl.dblPage8Weight;
//                                brandexLogoShortModelData.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

//                                brandexLogoShortModelData.dblIndex = indexSum;
//                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
//                                brandexLogoShortModelData.intRed = dataEl.intRed;
//                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
//                                brandexLogoShortModelData.intBlue = dataEl.intBlue;


//                                //for the chart - 
//                                double averagePage1WeightedValueForChart = 0.0;
//                                double averagePage2WeightedValueForChart = 0.0;
//                                double averagePage3WeightedValueForChart = 0.0;
//                                double averagePage4WeightedValueForChart = 0.0;
//                                double averagePage5WeightedValueForChart = 0.0;

//                                if (average1MaxLogo > 0)
//                                {
//                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight) * brandexLogoscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForChart = 0;
//                                }

//                                if (average2MaxLogo > 0)
//                                {
//                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight) * brandexLogoscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForChart = 0;
//                                }

//                                if (average3MaxLogo > 0)
//                                {
//                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight) * brandexLogoscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValueForChart = 0;
//                                }

//                                if (average4MaxLogo > 0)
//                                {
//                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight) * brandexLogoscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValueForChart = 0;
//                                }

//                                if (average5MaxLogo > 0)
//                                {
//                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight) * brandexLogoscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage5WeightedValueForChart = 0;
//                                }

//                                double indexSumForChart = averagePage1WeightedValueForChart +
//                                                          averagePage2WeightedValueForChart +
//                                                          averagePage3WeightedValueForChart +
//                                                          averagePage4WeightedValueForChart +
//                                                          averagePage5WeightedValueForChart;

//                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

//                                brandexLogoShortModelData.dblAveragePage1WeightedForChart = Math.Round(averagePage1WeightedValueForChart, 1);

//                                brandexLogoShortModelData.dblAveragePage2WeightedForChart = Math.Round(averagePage2WeightedValueForChart, 1);

//                                brandexLogoShortModelData.dblAveragePage3WeightedForChart = Math.Round(averagePage3WeightedValueForChart, 1);

//                                brandexLogoShortModelData.dblAveragePage4WeightedForChart = Math.Round(averagePage4WeightedValueForChart, 1); ;

//                                brandexLogoShortModelData.dblAveragePage5WeightedForChart = Math.Round(averagePage5WeightedValueForChart, 1); ;

//                                brandexLogoShortModelData.dblIndexForChart = indexSumForChart;
//                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
//                                brandexLogoShortModelData.intRed = dataEl.intRed;
//                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
//                                brandexLogoShortModelData.intBlue = dataEl.intBlue;

//                                brandexLogoShortData.Add(brandexLogoShortModelData);
//                            }

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexLogoTemplate" + brandexLogoShortData.Count + ".pptx";

//                            dLLClass.BrandexLogoMethod(CreateTargetPath(sourcePath), brandexLogoShortData, this.breakdown);
//                        }
//                        break;








//                    case "SALA":
//                    case "Sound Alike-Look Alike":
//                        List<SALANewModel> salaData = new List<SALANewModel>();
//                        salaData = JsonConvert.DeserializeObject<List<SALANewModel>>(data, jsonSettings);

//                        if (salaData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            CreateTargetPath(sourcePath);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";
//                            dLLClass.SALANewMethod(CreateTargetPath(sourcePath), salaData);
//                        }
//                        break;





//                    case "Negative Communication":
//                    case "Negative or Offensive Communication":


//                        List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
//                        negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data);

//                        if (negCommData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + negCommData.Count + ".pptx";
//                            dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), negCommData);
//                        }
//                        break;


//                    case "Negative Communication Rationale":


//                        List<NegativeCommunicationRationaleModel> negCommRationaleData = new List<NegativeCommunicationRationaleModel>();
//                        negCommRationaleData = JsonConvert.DeserializeObject<List<NegativeCommunicationRationaleModel>>(data);

//                        if (negCommRationaleData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunicationRationalePlaceholder.pptx";
//                            dLLClass.NegativeCommRationaleMethod(CreateTargetPath(sourcePath), negCommRationaleData, this.breakdown);
//                        }
//                        break;


//                    case "Fit to Theraputic Category":
//                    case "Fit to Therapeutic Category":
//                        List<FitToTheraputicCategoryModel> theraputicData = new List<FitToTheraputicCategoryModel>();
//                        theraputicData = JsonConvert.DeserializeObject<List<FitToTheraputicCategoryModel>>(data);

//                        List<FitToTheraputicCategoryModel> testnameCountList = new List<FitToTheraputicCategoryModel>();

//                        for (int i = 0; i < theraputicData.Count; i++)
//                        {
//                            var groupData = theraputicData.GroupBy(x => x.strTestName).ToList();
//                            if (groupData != null)
//                            {
//                                var firstTestnameList = groupData[0].ToList();
//                                testnameCountList = firstTestnameList;
//                            }
//                            break;
//                        }

//                        int tableHeaderCounter = testnameCountList.Count;

//                        if (theraputicData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticCategory" + tableHeaderCounter + ".pptx";
//                            dLLClass.FitToTheraputicMethod(CreateTargetPath(sourcePath), theraputicData);
//                        }
//                        break;






//                    //    List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
//                    //    negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data, jsonSettings);


//                    //    if (negCommData.Count == 0)
//                    //    {
//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                    //        CreateTargetPath(sourcePath);
//                    //    }
//                    //    else
//                    //    {
//                    //        if (this.template == "Sound Alike-Look Alike")
//                    //        {
//                    //            this.template = "SALA";
//                    //        }

//                    //        //sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";


//                    //        int dataIndex = 0;
//                    //        double sum = 0;
//                    //        List<NegativeCommunicationModelShort> lstNegCommValue = new List<NegativeCommunicationModelShort>();

//                    //        var groupApiData = negCommData.GroupBy(item => item.strTestName).OrderBy(group => group.Key).ToList();


//                    //        while (dataIndex < groupApiData.Count)
//                    //        {
//                    //            sum = 0;
//                    //            var testNameGroup = groupApiData[dataIndex];
//                    //            var testNameData = testNameGroup.ToList();

//                    //            foreach (var Tdata in testNameData)
//                    //            {
//                    //                sum += Tdata.intSum;
//                    //            }

//                    //            double total = testNameData.FirstOrDefault().intTotal;

//                    //            double percentageVal = (double)(((total - sum) / total) * 100);

//                    //            double remainingPercentageVal = 100 - percentageVal;

//                    //            // lstNegCommValue.FirstOrDefault().percentage = percentageVal.ToString();

//                    //            NegativeCommunicationModelShort ncomShort = new NegativeCommunicationModelShort();

//                    //            ncomShort.percentage = percentageVal.ToString();
//                    //            ncomShort.strTestName = testNameData.FirstOrDefault().strTestName;
//                    //            ncomShort.intBlue = testNameData.Max(x => x.intBlue);
//                    //            ncomShort.intGreen = testNameData.Max(x => x.intGreen);
//                    //            ncomShort.intRed = testNameData.Max(x => x.intRed);
//                    //            ncomShort.remainingPercentage = remainingPercentageVal.ToString();

//                    //            lstNegCommValue.Add(ncomShort);

//                    //            dataIndex++;
//                    //        }


//                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + lstNegCommValue.Count + ".pptx";
//                    //        dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), lstNegCommValue);




//                    //    }
//                    //    break;


//                    case "Distinctiveness":
//                        List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
//                        distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

//                        if (distinctData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            notAvailable(sourcePath, template, breakdown);
//                        }
//                        else
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
//                            dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
//                        }
//                        break;



//                    case "BRANDEX MEDICAL DEVICE 1":
//                    case "Brandex Medical Device 1":

//                        List<BrandexMedicalDevice1Model> brandexMedicalData = new List<BrandexMedicalDevice1Model>();
//                        brandexMedicalData = JsonConvert.DeserializeObject<List<BrandexMedicalDevice1Model>>(data);

//                        if (brandexMedicalData.Count == 0)
//                        {
//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
//                        }
//                        else
//                        {
//                            List<BrandexMedicalDevice1ShortModel> brandexMedicalShortData = new List<BrandexMedicalDevice1ShortModel>();

//                            double brandexMedicalAverage1Max = 0.0;
//                            double brandexMedicalAverage2Max = 0.0;
//                            double brandexMedicalAverage3Max = 0.0;
//                            double brandexMedicalAverage4Max = 0.0;
//                            double brandexMedicalAverage5Max = 0.0;

//                            for (int i = 0; i < brandexMedicalData.Count; i++)
//                            {
//                                var dblAverage1max = brandexMedicalData[i].dblAveragePage1;
//                                if (dblAverage1max == 0.0)
//                                {
//                                    brandexMedicalAverage1Max = 0;
//                                }
//                                if (brandexMedicalAverage1Max < dblAverage1max)
//                                {
//                                    brandexMedicalAverage1Max = dblAverage1max;
//                                }

//                                var dblAverage2max = brandexMedicalData[i].dblAveragePage2;
//                                if (dblAverage2max == 0.0)
//                                {
//                                    brandexMedicalAverage2Max = 0;
//                                }
//                                if (brandexMedicalAverage2Max < dblAverage2max)
//                                {
//                                    brandexMedicalAverage2Max = dblAverage2max;
//                                }

//                                var dblAverage3max = brandexMedicalData[i].dblAveragePage3;
//                                if (dblAverage3max == 0.0)
//                                {
//                                    brandexMedicalAverage3Max = 0;
//                                }
//                                if (brandexMedicalAverage3Max < dblAverage3max)
//                                {
//                                    brandexMedicalAverage3Max = dblAverage3max;
//                                }

//                                var dblAverage4max = brandexMedicalData[i].dblAveragePage4;
//                                if (dblAverage4max == 0.0)
//                                {
//                                    brandexMedicalAverage4Max = 0;
//                                }
//                                if (brandexMedicalAverage4Max < dblAverage4max)
//                                {
//                                    brandexMedicalAverage4Max = dblAverage4max;
//                                }

//                                var dblAverage5max = brandexMedicalData[i].dblAveragePage5;
//                                if (dblAverage5max == 0.0)
//                                {
//                                    brandexMedicalAverage5Max = 0;
//                                }
//                                if (brandexMedicalAverage5Max < dblAverage5max)
//                                {
//                                    brandexMedicalAverage5Max = dblAverage5max;
//                                }
//                            }

//                            double brandexMedicalscalingFactor = 1.0023;

//                            for (int i = 0; i < brandexMedicalData.Count; i++)
//                            {
//                                //for the table
//                                BrandexMedicalDevice1ShortModel brandexSafetyShortModel = new BrandexMedicalDevice1ShortModel();

//                                var dataEl = brandexMedicalData[i];

//                                double averagePage1WeightedValue = 0.0;
//                                double averagePage2WeightedValue = 0.0;
//                                double averagePage3WeightedValue = 0.0;
//                                double averagePage4WeightedValue = 0.0;
//                                double averagePage5WeightedValue = 0.0;

//                                if (brandexMedicalAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValue = 0;
//                                }

//                                if (brandexMedicalAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValue = 0;
//                                }

//                                if (brandexMedicalAverage3Max > 0)
//                                {
//                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValue = 0;
//                                }

//                                if (brandexMedicalAverage4Max > 0)
//                                {
//                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValue = 0;
//                                }

//                                if (brandexMedicalAverage5Max > 0)
//                                {
//                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight;
//                                }
//                                else
//                                {
//                                    averagePage5WeightedValue = 0;
//                                }

//                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
//                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexMedicalscalingFactor;

//                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
//                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
//                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
//                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
//                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
//                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
//                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
//                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
//                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
//                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
//                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

//                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
//                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
//                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

//                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
//                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
//                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

//                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
//                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
//                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

//                                brandexSafetyShortModel.dblIndex = indexSum;
//                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                                brandexSafetyShortModel.intRed = dataEl.intRed;
//                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                                //for the chart - 

//                                double averagePage1WeightedValueForChart = 0.0;
//                                double averagePage2WeightedValueForChart = 0.0;
//                                double averagePage3WeightedValueForChart = 0.0;
//                                double averagePage4WeightedValueForChart = 0.0;
//                                double averagePage5WeightedValueForChart = 0.0;

//                                if (brandexMedicalAverage1Max > 0)
//                                {
//                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight) * brandexMedicalscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage1WeightedValueForChart = 0;
//                                }

//                                if (brandexMedicalAverage2Max > 0)
//                                {
//                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight) * brandexMedicalscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage2WeightedValueForChart = 0;
//                                }

//                                if (brandexMedicalAverage3Max > 0)
//                                {
//                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight) * brandexMedicalscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage3WeightedValueForChart = 0;
//                                }

//                                if (brandexMedicalAverage4Max > 0)
//                                {
//                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight) * brandexMedicalscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage4WeightedValueForChart = 0;
//                                }

//                                if (brandexMedicalAverage5Max > 0)
//                                {
//                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight) * brandexMedicalscalingFactor;
//                                }
//                                else
//                                {
//                                    averagePage5WeightedValueForChart = 0;
//                                }

//                                double indexSumForChart = averagePage1WeightedValueForChart +
//                                                          averagePage2WeightedValueForChart +
//                                                          averagePage3WeightedValueForChart +
//                                                          averagePage4WeightedValueForChart +
//                                                          averagePage5WeightedValueForChart;

//                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

//                                brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

//                                brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
//                                ;
//                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
//                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
//                                brandexSafetyShortModel.intRed = dataEl.intRed;
//                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
//                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
//                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


//                                brandexMedicalShortData.Add(brandexSafetyShortModel);
//                            }

//                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexMedicalDevice1Template" + brandexMedicalShortData.Count + ".pptx";
//                            dLLClass.BrandexMedicalDevice1Method(CreateTargetPath(sourcePath), brandexMedicalShortData);
//                        }
//                        break;











//                    default:
//                        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                        notAvailable(CreateTargetPath(sourcePath), template, breakdown);
//                        break;


//                        //case "":
//                        //    List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
//                        //    distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

//                        //    if (distinctData.Count == 0)
//                        //    {
//                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                        //        notAvailable(sourcePath, template, breakdown);
//                        //    }
//                        //    else
//                        //    {
//                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
//                        //        dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
//                        //    }
//                        //    break;


//                        //default:
//                        //    sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
//                        //    notAvailable(CreateTargetPath(sourcePath), template, breakdown);
//                        //    break;



//                }
//            }




//        }


//        public async Task<string> addChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string BreakDown)
//        {
//            //hello

//            try
//            {
//                await Task.Run(() => fnaddChartsToFinalTemplate1(project, charts, finalTemplate, BreakDown));
//                return "Process sucessful";
//            }
//            catch (Exception)
//            {
//                throw;
//            }

//        }


//        public string getAttributeTitle(string project, string chart)
//        {
//            try
//            {
//                string repText = "";
//                System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

//                foreach (DataRow row in dt1.Rows)
//                {
//                    repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
//                }

//                return repText;
//            }
//            catch (Exception)
//            {
//                throw;
//            }
//        }


//        public async Task fnaddChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string breakDown)
//        {
//            try
//            {
//                if (!Directory.Exists($"C:\\excelfiles\\{project}\\Final"))
//                {
//                    Directory.CreateDirectory($"C:\\excelfiles\\{project}\\Final");
//                }

//                //temp copy the final template file to the path 

//                // File.Copy("\\\\miafs02\\Market Research\\MR Programs\\ExcelCharts_Chartsdll\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", true);

//                File.Copy("C:\\ExcelChartFiles\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", true);

//                //get all the pagegroup names for the chart 

//                List<clschartPageGroup> lstPg = getProjectPagegroupNames(project);


//                //get the display names of the charts in excel sheet app

//                List<clschartPageDisplayName> lstChartDispNames = getProjectChartDisplayNames(project);

//                string op = "";

//                //get the page settings for the template 

//                System.Data.DataTable dt = clsData.MRData.getDataTable("ExcelChartsPrc_getPPTFinalSettings " + "'" + finalTemplate + "'," + "'" + "BI - 2024" + "'");
//                List<clsPPTFinalSettings> lstPPTFinalSettings = new List<clsPPTFinalSettings>();

//                int DelLastPage = 0;
//                int DelFirstPage = 0;
//                List<int> intChartPages = new List<int>();


//                foreach (DataRow row in dt.Rows)
//                {
//                    // Create a new object for each row and add it to the list
//                    clsPPTFinalSettings data1 = new clsPPTFinalSettings
//                    {
//                        intPPTSlideIndexFirst = Convert.ToInt32(row["intPPTSlideIndexFirst"]),
//                        intPPTSlideIndexLast = Convert.ToInt32(row["intPPTSlideIndexLast"]),
//                        strTemplateName = Convert.ToString(row["strTemplateName"]),
//                        strTemplateSourcePath = Convert.ToString(row["strTemplateSourcePath"]),
//                        strPageGroupName = Convert.ToString(row["strPageGroupName"]),
//                        strPageGroupType = Convert.ToString(row["strPageGroupType"]),
//                    };
//                    lstPPTFinalSettings.Add(data1);
//                }



//                int specialChartTypeCount = 0;

//                //copy a final in the path

//                createFolder($"C:\\excelfiles\\{project}\\Final");
//                copyFile("C:\\ExcelChartsTemplatesNew\\ExcelCharts_ChartTemplates\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx", $"C:\\excelfiles\\{project}\\Final\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx");


//                lstPPTFinalSettings = lstPPTFinalSettings.OrderByDescending(p => p.intPPTSlideIndexFirst).ToList();
//                int chartsCompCount = 0;
//                List<string> chartsCompleted = new List<string>();
//                List<string> pageGroupNameCompleted = new List<string>();
//                List<clsChartDislayNamePageGroupName> lstchartDispPageGroupName = new List<clsChartDislayNamePageGroupName>();


//                //
//                foreach (clschartPageDisplayName obj in lstChartDispNames)
//                {
//                    foreach (string chart in charts)
//                    {
//                        if (obj?.strPageName == chart)
//                        {
//                            obj.isReportSelectedByUser = true;


//                            //updating Attribute evaluation cover page :


//                            if (obj?.strPageType?.ToLower() == "attribute evaluation")
//                            {
//                                if (obj.strPageName.Contains("1"))
//                                {
//                                    //update the slide 38

//                                    string repText = getAttributeTitle(project, chart);


//                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);



//                                    //await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);


//                                }

//                                if (obj.strPageName.Contains("2"))
//                                {
//                                    //update the slide 38

//                                    string repText = getAttributeTitle(project, chart);

//                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", repText, 38);

//                                    // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

//                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle2", repText, 38);


//                                }

//                                if (obj.strPageName.Contains("3"))
//                                {
//                                    //update the slide 38

//                                    string repText = getAttributeTitle(project, chart);


//                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", repText, 38);

//                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

//                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle3", repText, 38);


//                                }

//                                if (obj.strPageName.Contains("4"))
//                                {
//                                    //update the slide 38

//                                    string repText = getAttributeTitle(project, chart);

//                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", repText, 38);

//                                    //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

//                                    // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle4", repText, 38);


//                                }


//                            }



//                        }

//                    }

//                }


//                //in case if the attribute evaluations are not there 

//                List<clschartPageDisplayName> lstAtts = lstChartDispNames?.Where(p => p?.strPageType?.ToLower() == "attribute evaluation").ToList();


//                if (!lstAtts.Any(obj => obj.strPageName.Contains("1")) || lstAtts.Any(obj => obj.strPageName.Contains("1") && obj?.isReportSelectedByUser == false))
//                {

//                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);

//                }

//                if (!lstAtts.Any(obj => obj.strPageName.Contains("2")) || lstAtts.Any(obj => obj.strPageName.Contains("2") && obj?.isReportSelectedByUser == false))
//                {

//                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);

//                }


//                if (!lstAtts.Any(obj => obj.strPageName.Contains("3")) || lstAtts.Any(obj => obj.strPageName.Contains("3") && obj?.isReportSelectedByUser == false))
//                {

//                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);

//                }

//                if (!lstAtts.Any(obj => obj.strPageName.Contains("4")) || lstAtts.Any(obj => obj.strPageName.Contains("4") && obj?.isReportSelectedByUser == false))
//                {

//                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);

//                }

//                //merge


//                //bool first chart name replaced

//                bool firstChartTextReplaced = false;

//                foreach (clsPPTFinalSettings objclsPPT in lstPPTFinalSettings)
//                {

//                    try
//                    {
//                        if (chartsCompCount == charts.Count())
//                        {
//                            break;
//                        }

//                        specialChartTypeCount = 0;
//                        //check if chart type is selected  

//                        // bool ifThePageTypeIsSelected = true;


//                        //first chart project name and details change 

//                        if (!firstChartTextReplaced)
//                        {
//                            DateTime currentDate = DateTime.Now;

//                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "ProjectName", project, 0);

//                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<", "", 0);

//                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", ">>", "", 0);

//                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Month", currentDate.ToString("MMMM"), 0);

//                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Year", DateTime.Now.Year.ToString(), 0);


//                            //replacing the attribute evaluation cover page :

//                            firstChartTextReplaced = true;

//                        }



//                        //delete the slides that are not part of the report 


//                        if (!lstChartDispNames.Any(p => p.strPageType == objclsPPT?.strPageGroupType))
//                        {
//                            // skip the Attribute Evaluation Cover if the report has Attribute eavluation

//                            if (pageGroupNameCompleted.Any(z => z.Contains("Attribute")) && objclsPPT?.strPageGroupType == "Attribute Evaluation Cover")
//                            {
//                                continue;
//                            }

//                            else
//                            {
//                                DelLastPage = objclsPPT.intPPTSlideIndexLast;
//                                DelFirstPage = objclsPPT.intPPTSlideIndexFirst;


//                                if (DelLastPage - DelFirstPage == 0)
//                                {
//                                    for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
//                                    {
//                                        if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
//                                        {
//                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

//                                        }
//                                    }
//                                }

//                                else
//                                {
//                                    for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
//                                    {
//                                        if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
//                                        {
//                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

//                                        }

//                                    }

//                                }


//                                continue;

//                            }





//                        }



//                        foreach (string chart in charts)
//                        {

//                            //if the chart is completed break

//                            if (pageGroupNameCompleted.Contains(objclsPPT?.strPageGroupName))
//                            {
//                                break;
//                            }


//                            //get the chart group type

//                            List<clschartPageDisplayName> lstCheck = lstChartDispNames.Where(p => p?.strPageName == chart)?.ToList();

//                            bool noproceed = false;

//                            //delete the slides from the template if a report is not selected .

//                            bool isReportSelected = false;


//                            if (lstChartDispNames.Any(item => item?.strPageType == objclsPPT?.strPageGroupType))
//                            {

//                                if (objclsPPT.strPageGroupType.ToLower() == "exaggerative-inappropriate")
//                                {

//                                    //if( lstChartDispNames.Any(item =>item.strPageType.ToLower() == "exaggerative-inappropriate" && getNumbersFromString(item.strPageName)==getNumbersFromString(chart)))
//                                    //{
//                                    //     isReportSelected = true;
//                                    //}



//                                    List<clschartPageDisplayName> lstFiltered = lstChartDispNames?.Where(item => item?.strPageType?.ToLower() == "exaggerative-inappropriate")?.ToList();

//                                    if(lstFiltered.Count>0)
//                                    {
//                                        foreach (clschartPageDisplayName obj in lstFiltered)
//                                        {
//                                            if (getNumbersFromString(obj?.strPageName) == getNumbersFromString(objclsPPT?.strPageGroupName))
//                                            {

//                                                if ((bool)(obj?.isReportSelectedByUser))
//                                                {
//                                                    isReportSelected = true;
//                                                    break;

//                                                }

//                                            }

//                                        }
//                                    }


//                                }

//                                else if (objclsPPT.strPageGroupType.ToLower() == "attribute evaluation")
//                                {


//                                    List<clschartPageDisplayName> lstFiltered = lstChartDispNames?.Where(item => item?.strPageType?.ToLower() == "attribute evaluation" || item.strPageType.ToLower() == "attribute evaluation aggregate")?.ToList();

//                                    if(lstFiltered.Count>0)
//                                    {
//                                        foreach (clschartPageDisplayName obj in lstFiltered)
//                                        {
//                                            if (getNumbersFromString(obj?.strPageName) == getNumbersFromString(objclsPPT?.strPageGroupName))
//                                            {
//                                                if ((bool)(obj?.isReportSelectedByUser))
//                                                {
//                                                    isReportSelected = true;
//                                                    break;

//                                                }
//                                            }

//                                        }

//                                    }



//                                }

//                                else
//                                {
//                                    List<clschartPageDisplayName> lstFiltered = lstChartDispNames?.Where(item => item?.strPageType?.ToLower() == objclsPPT?.strPageGroupName?.ToLower())?.ToList();

//                                    if(lstFiltered.Count>0)
//                                    {
//                                        foreach (clschartPageDisplayName obj in lstFiltered)
//                                        {
//                                            if ((bool)(obj?.isReportSelectedByUser))
//                                            {
//                                                isReportSelected = true;
//                                                break;

//                                            }

//                                        }
//                                    }



//                                }


//                            }


//                            //if (lstChartDispNames.Any(item => item.strPageType == objclsPPT.strPageGroupType))
//                            //{
//                            //    foreach (clschartPageDisplayName item in lstChartDispNames)
//                            //    {
//                            //        if (item.strPageType == objclsPPT.strPageGroupType && item.strPageName==chart)
//                            //        {
//                            //            if (!item.isReportSelectedByUser)
//                            //            {
//                            //                isReportSelected = false;
//                            //                break;
//                            //            }
//                            //        }
//                            //    }


//                            //}

//                            //else
//                            //{
//                            //    isReportSelected = false;


//                            //}


//                            if (!isReportSelected)
//                            {
//                                //delete if the project is not checked 
//                                DelLastPage = objclsPPT.intPPTSlideIndexLast;
//                                DelFirstPage = objclsPPT.intPPTSlideIndexFirst;

//                                if (DelLastPage - DelFirstPage == 1 && objclsPPT?.strPageGroupName?.ToLower() == "attribute evaluation cover")
//                                {
//                                    List<clschartPageDisplayName> lstFiltered = lstChartDispNames?.Where(item => item?.strPageType?.ToLower() == "attribute evaluation" && item?.isReportSelectedByUser == true)?.ToList();

//                                    if (lstFiltered.Count > 0)
//                                    {
//                                        break;
//                                    }

//                                    else
//                                    {
//                                        DelLastPage = DelFirstPage;
//                                    }
//                                }

//                                if (!pageGroupNameCompleted.Contains(objclsPPT?.strPageGroupName) && (DelLastPage != 0 && DelFirstPage != 0))
//                                {

//                                    if (DelLastPage - DelFirstPage == 0)
//                                    {
//                                        for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
//                                        {
//                                            if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
//                                            {
//                                                op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

//                                            }

//                                        }

//                                    }

//                                    else
//                                    {
//                                        for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
//                                        {
//                                            if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
//                                            {
//                                                op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

//                                            }

//                                        }

//                                    }


//                                }


//                                break;


//                            }



//                            foreach (clschartPageDisplayName oj in lstCheck)
//                            {

//                                if (objclsPPT?.strPageGroupType?.ToLower() == "exaggerative-inappropriate" || objclsPPT?.strPageGroupType?.ToLower() == "attribute evaluation")
//                                {

//                                    if (oj.strPageName.ToLower() == "attribute evaluation aggregate")
//                                    {
//                                        oj.strPageType = "Attribute Evaluation";
//                                    }

//                                    if ((oj?.strPageType == objclsPPT?.strPageGroupType) && (oj?.strPageName == chart) && (getNumbersFromString(objclsPPT?.strPageGroupName) == getNumbersFromString(chart)))
//                                    {
//                                        noproceed = false;
//                                        break;

//                                    }

//                                    else
//                                    {
//                                        noproceed = true;
//                                        break;

//                                    }

//                                }

//                                else
//                                {
//                                    if ((oj?.strPageType == objclsPPT?.strPageGroupType) && (oj?.strPageName == chart))
//                                    {
//                                        noproceed = false;
//                                        break;

//                                    }

//                                    else
//                                    {
//                                        noproceed = true;
//                                        break;

//                                    }
//                                }
//                            }

//                            if (noproceed)
//                            {
//                                continue;
//                            }




//                            if (!chartsCompleted.Contains(chart))
//                            {
//                                if (objclsPPT?.strPageGroupName != "Main Cover")
//                                {
//                                    DelLastPage = objclsPPT.intPPTSlideIndexLast;
//                                    DelFirstPage = objclsPPT.intPPTSlideIndexFirst;


//                                    //add

//                                    string firstMatchingName = lstPg
//                                   .Where(p => p.strPageGroup == chart)
//                                   .Select(p => p.strPageGroupType)
//                                   .FirstOrDefault();


//                                    if (objclsPPT?.strPageGroupType != "Attribute Evaluation" && objclsPPT?.strPageGroupType != "Exaggerative-Inappropriate")
//                                    {
//                                        //item array of pages with slide numbers 
//                                        intChartPages = new List<int>();

//                                        if (DelLastPage > DelFirstPage)
//                                        {
//                                            if (objclsPPT?.strPageGroupType == "Phonetic Testing" || objclsPPT?.strPageGroupType == "JSCAN")
//                                            {
//                                                intChartPages.Add(DelLastPage);

//                                            }

//                                            else if (objclsPPT?.strPageGroupType == "Sound Alike-Look Alike" || objclsPPT?.strPageGroupType == "Medical Terms Similarity" || objclsPPT.strPageGroupType == "Non-Medical Terms Similarity" || objclsPPT.strPageGroupType == "Brandex Strategic Distinctiveness" || objclsPPT.strPageGroupType == "Brandex Safety")
//                                            {
//                                                for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
//                                                {
//                                                    intChartPages.Add(DelFirstPage + l);
//                                                }

//                                            }
//                                            else
//                                            {
//                                                for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
//                                                {
//                                                    intChartPages.Add(DelFirstPage + l - 1);
//                                                }


//                                            }



//                                        }
//                                        else
//                                        {
//                                            intChartPages.Add(DelLastPage);
//                                        }



//                                        if (objclsPPT?.strPageGroupName?.ToLower() == chart.ToLower() || objclsPPT?.strPageGroupName?.ToLower() == firstMatchingName?.ToLower())
//                                        {
//                                            for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
//                                            {
//                                                op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
//                                                //Thread.Sleep(100);
//                                            }

//                                            if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
//                                            {
//                                                for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
//                                                {
//                                                    intChartPages.RemoveAt(i);
//                                                }

//                                            }

//                                            if (File.Exists(getIndividualChartPath(chart, project, breakDown)))
//                                            {
//                                                op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);


//                                            }


//                                            chartsCompCount = chartsCompCount + 1;

//                                            lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT?.strPageGroupName));

//                                            chartsCompleted.Add(chart);

//                                            pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);
//                                        }


//                                    }

//                                    else
//                                    {
//                                        //item array of pages with slide numbers 
//                                        intChartPages = new List<int>();

//                                        if (DelLastPage > DelFirstPage && DelLastPage - DelFirstPage > 1)
//                                        {
//                                            for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
//                                            {
//                                                intChartPages.Add(DelFirstPage + l - 1);
//                                            }

//                                        }
//                                        else
//                                        {
//                                            intChartPages.Add(DelLastPage);
//                                        }

//                                        //get the attribute evaluation category and Exaggerative 

//                                        string PageType = lstChartDispNames?.FirstOrDefault(item => item.strPageName == chart)?.strPageType;

//                                        string resPageType = "";

//                                        op = "";

//                                        if (PageType == "Attribute Evaluation" && chart.Contains("1"))
//                                        {

//                                            resPageType = "Attribute Evaluation 1";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");





//                                        }

//                                        else if (PageType == "Attribute Evaluation" && chart.Contains("2"))
//                                        {
//                                            resPageType = "Attribute Evaluation 2";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "FirstPage");
//                                            // await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "Attribute #2", 38);

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


//                                            //}



//                                        }

//                                        else if (PageType == "Attribute Evaluation" && chart.Contains("3"))
//                                        {

//                                            resPageType = "Attribute Evaluation 3";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "FirstPage");
//                                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "Attribute #3", 38);

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


//                                            //}



//                                        }

//                                        else if (PageType == "Attribute Evaluation" && chart.Contains("4"))
//                                        {
//                                            resPageType = "Attribute Evaluation 4";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "FirstPage");
//                                            await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "Attribute #4", 38);

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "", "", 38);


//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


//                                            //}

//                                            //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
//                                            //{
//                                            //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


//                                            //}


//                                        }


//                                        else if (PageType.Contains("Attribute Evaluation") && chart.Contains("Aggregate") && chart.Contains("Attribute"))
//                                        {
//                                            resPageType = "Attribute Evaluation Aggregate";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "FirstPage");


//                                        }


//                                        else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "1" || getNumbersFromString(chart) == "01"))
//                                        {
//                                            resPageType = "01 Untrue";

//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "LastPage");

//                                            // DelFirstPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "FirstPage");

//                                            DelFirstPage = DelLastPage;


//                                        }


//                                        else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "2" || getNumbersFromString(chart) == "02"))
//                                        {

//                                            resPageType = "02 Mislead";

//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "LastPage");

//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "FirstPage");


//                                        }


//                                        else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "3" || getNumbersFromString(chart) == "03"))
//                                        {

//                                            resPageType = "03 Exagg";

//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "LastPage");

//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "FirstPage");


//                                        }

//                                        //if (objclsPPT.strPageGroupName.ToLower() == resPageType.ToLower())
//                                        //{


//                                        if (resPageType == "01 Untrue")
//                                        {

//                                            op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - 1);
//                                        }

//                                        else
//                                        {
//                                            for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
//                                            {
//                                                op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
//                                                //Thread.Sleep(100);
//                                            }
//                                        }




//                                        if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
//                                        {
//                                            for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
//                                            {
//                                                intChartPages.RemoveAt(i);
//                                            }

//                                        }

//                                        op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);



//                                        if (PageType == "Attribute Evaluation" && chart.Contains("1"))
//                                        {

//                                            resPageType = "Attribute Evaluation 1";
//                                            DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
//                                            DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");

//                                            //update the slide 38

//                                            //update the slide 38

//                                            string repText = "";
//                                            System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

//                                            foreach (DataRow row in dt1.Rows)
//                                            {
//                                                repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
//                                            }






//                                            await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);

//                                            await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

//                                            await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);



//                                            chartsCompCount = chartsCompCount + 1;

//                                            lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT.strPageGroupName));

//                                            chartsCompleted.Add(chart);

//                                            pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);

//                                        }
//                                        // }




//                                    }









//                                }





//                            }



//                        }


//                    }
//                    catch (Exception)
//                    {
//                        throw;
//                    }










//                }





//                //get the final file in  the folder

//                createFolder($"C:\\excelfiles\\{project}\\PPT");
//                createFolder($"C:\\excelfiles\\{project}\\PPT\\01 Final");


//                while (true)
//                {


//                    if (File.Exists($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
//                    {
//                        copyFile($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\PPT\\01 Final\\ {project}_{finalTemplate}_{breakDown}.pptx");


//                        break;

//                    }

//                }



//            }
//            catch (Exception ex)
//            {
//                throw ex;
//            }
//        }



//        private int getPageIndex(List<clsPPTFinalSettings> lst, string pageGroupName, string pageIndex)
//        {
//            int result = 0;

//            if (pageIndex == "LastPage")
//            {
//                result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexLast;
//            }

//            else if (pageIndex == "FirstPage")
//            {

//                result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexFirst;
//            }

//            return result;
//        }


//        private bool checkAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName)
//        {

//            bool result = false;
//            if (lstPageDispName.Any(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true))
//            {

//                result = true;

//            }

//            return result;

//        }

//        private bool checkParticularAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
//        {

//            bool result = false;

//            var lst = lstPageDispName.Where(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true);

//            foreach (clschartPageDisplayName obj in lst)
//            {
//                if (obj.strPageName.Contains(attValue.ToString()))
//                {

//                    result = true;

//                }

//            }

//            return result;

//        }


//        private bool checkParticularExaggIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
//        {

//            bool result = false;

//            var lst = lstPageDispName.Where(m => m.strPageType == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true && m.strPageName.Contains(attValue.ToString()));

//            foreach (clschartPageDisplayName obj in lst)
//            {
//                if (obj.strPageName.Contains(attValue.ToString()))
//                {
//                    result = true;

//                }

//            }


//            return result;

//        }


//        public static string getNumbersFromString(string input)
//        {
//            try
//            {
//                // Use a regular expression to match all digits in the string
//                string numbers = Regex.Replace(input, @"[^0-9]", "");
//                return numbers;
//            }
//            catch (Exception)
//            {

//                throw;
//            }
//        }


//        private bool checkParticularChartIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue, string chart)
//        {

//            bool result = false;

//            var lst = lstPageDispName.Where(m => m.strPageName == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true);

//            foreach (clschartPageDisplayName obj in lst)
//            {
//                if (obj.strPageName.Contains(attValue.ToString()))
//                {
//                    result = true;

//                }

//            }


//            return result;

//        }




//        private string getReverseChartName(string chartName)
//        {
//            if (chartName == "Attribute Evaluation 1")
//            {
//                return "Attribute 1";
//            }
//            else if (chartName == "Attribute Evaluation 2")
//            {
//                return "Attribute 2";
//            }

//            else if (chartName == "Attribute Evaluation 3")
//            {
//                return "Attribute 3";
//            }

//            else if (chartName == "Attribute Evaluation 4")
//            {
//                return "Attribute 4";
//            }


//            else
//            {
//                return chartName;
//            }

//        }



//        private int getmaxSlideid(string fileName)
//        {
//            int maxSlideNumber = 0;
//            if (!File.Exists(fileName))
//            {
//                return 0;
//            }

//            else
//            {
//                maxSlideNumber = 1;
//                using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
//                {

//                    PresentationPart presentationPart = doc.PresentationPart;

//                    maxSlideNumber = presentationPart.SlideParts.Count();
//                }

//                return Convert.ToInt32(maxSlideNumber);


//            }




//        }



//        private string getIndividualChartPath(string chart, string projectName, string breakDown)
//        {



//            string outputpath = $"C:\\excelfiles\\{projectName}\\{chart}_{breakDown}.pptx";


//            //get the right chart name 

//            return outputpath;

//        }

//        public int CountSlides(string presentationFile)
//        {
//            try
//            {
//                if (File.Exists(presentationFile))
//                {
//                    using (PresentationDocument ppt = PresentationDocument.Open(presentationFile, false))
//                    {
//                        // Get the presentation part of the presentation document
//                        PresentationPart presentationPart = ppt?.PresentationPart;
//                        //if (presentationPart == null || presentationPart.Presentation == null)
//                        //{
//                        //    return 0;
//                        //}

//                        // Get the slide ID list
//                        SlideIdList slideIdList = presentationPart?.Presentation?.SlideIdList;
//                        //if (slideIdList == null)
//                        //{
//                        //    return 0;
//                        //}

//                        // Return the count of slide IDs
//                        return slideIdList.ChildElements.Count;
//                    }
//                }

//                else
//                {
//                    return 0;

//                }
//            }
//            catch (Exception)
//            {
//                throw;
//            }
//        }



//        private string getRightChartName(string pageName, int index)
//        {
//            if (pageName == "Attribute Evaluation" && index == 1)
//            {
//                return "Attribute Evaluation " + index;
//            }
//            else if (pageName == "Attribute Evaluation" && index == 2)
//            {
//                return "Attribute Evaluation " + index;
//            }
//            else if (pageName == "Attribute Evaluation" && index == 3)
//            {
//                return "Attribute Evaluation " + index;
//            }

//            else if (pageName == "Attribute Evaluation" && index == 3)
//            {
//                return "Attribute Evaluation " + index;
//            }

//            else if (pageName == "Exaggerative-Inappropriate" && index == 1)
//            {
//                return "01 Untrue";
//            }

//            else if (pageName == "Exaggerative-Inappropriate" && index == 2)
//            {
//                return "02 Misleading";
//            }

//            else if (pageName == "Exaggerative-Inappropriate" && index == 3)
//            {
//                return "03 Exagg";
//            }

//            else
//            {
//                return pageName;
//            }

//        }


//        private string getChartsPagegroupName(string projectName, string chartName)
//        {
//            string realchartName = "";

//            System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");



//            List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
//            {
//                strPageGroup = row.Field<string>("strPageGroup"),
//                strPageGroupType = row.Field<string>("strPageGroupType")
//            }).ToList();


//            List<clschartPageGroup> lstPgChart = lstPg.Where(p => p.strPageGroup == chartName).ToList();

//            foreach (clschartPageGroup obj in lstPgChart)
//            {

//                realchartName = obj.strPageGroupType;

//            }

//            return realchartName;

//        }



//        private List<clschartPageGroup> getProjectPagegroupNames(string projectName)
//        {
//            string realchartName = "";

//            System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");


//            List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
//            {
//                strPageGroup = row.Field<string>("strPageGroup"),
//                strPageGroupType = row.Field<string>("strPageGroupType")
//            }).ToList();


//            return lstPg;

//        }


//        private List<clschartPageDisplayName> getProjectChartDisplayNames(string projectName)
//        {
//            string realchartName = "";

//            System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPages]  " + 1 + ",'" + projectName + "'");

//            List<clschartPageDisplayName> lstPg = dt.AsEnumerable().Select(row => new clschartPageDisplayName
//            {
//                strPageName = row.Field<string>("strPageName"),
//                strPageType = row.Field<string>("strPageType"),
//                intPageID = row.Field<Int32>("intPageID").ToString(),
//                intPageIndex = row.Field<Int32>("intPageIndex").ToString(),
//                intSurveyID = row.Field<Int32>("intSurveyID").ToString(),

//            }).ToList();


//            return lstPg;

//        }


//        public static void createFolder(string path)
//        {
//            try
//            {
//                if (!Directory.Exists(path))
//                {
//                    // Try to create the directory.
//                    DirectoryInfo di = Directory.CreateDirectory(path);
//                }
//            }
//            catch (IOException ioex)
//            {
//                Console.WriteLine(ioex.Message);
//            }


//        }


//        public static void copyFile(string source, string destination)
//        {
//            try
//            {
//                File.Copy(source, destination, true);

//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine(ex.Message);
//            }


//        }


//        public static void notAvailable(string sourcePath, string template, string breakdown)
//        {
//            DLLClass dLLClass = new DLLClass();
//            APIWrapper obj = new APIWrapper();
//            dLLClass.NotAvailableMethod(sourcePath, template, breakdown);

//        }


//        public static void copyTemplates()
//        {

//            //lok please check currently doing only for the first time

//            string sourceDirectory = @"\\miafs02\Market Research\MR Programs\ExcelCharts_Chartsdll\Templates";
//            string destinationDirectory = "C:\\ExcelChartFiles\\Templates\\";

//            if (!Directory.Exists("C:\\ExcelChartFiles\\"))
//            {

//                if (!Directory.Exists("C:\\ExcelChartFiles\\"))
//                {
//                    // Create the directory if it doesn't exist
//                    Directory.CreateDirectory("C:\\ExcelChartFiles\\");
//                }

//                if (!Directory.Exists(destinationDirectory))
//                {
//                    // Create the directory if it doesn't exist
//                    Directory.CreateDirectory(destinationDirectory);

//                }

//                foreach (var file in Directory.GetFiles(sourceDirectory))
//                {
//                    File.Copy(file, Path.Combine(destinationDirectory, Path.GetFileName(file)), true);
//                }


//            }


//        }


//        public void notAvailable(string template)
//        {

//            //DLLClass obj = new DLLClass();
//            //obj.NotAvailableMethod(CreateTargetPath($"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx"), template);

//            //template.NotAvailableTemplate(destination, templateName, breakdown);

//        }


//        public void copyInsertSlide(string source, string destination, int pos)
//        {
//            clsMisc.CopySlide(source, destination, pos);

//        }
//    }
//}

using ClassLibrary2.BackgroundServicesDLL;
using ClassLibrary2.Models;
using ClassLibrary2.TaskLogger;
using ClassLibrary2.TaskTrackingDLL;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Newtonsoft.Json;
using OpenXmlDll;
using prjData;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using static OpenXmlDLLDotnetFramework.DLLTemplate;

namespace OpenXmlDLLDotnetFramework
{

    public class APIWrapper
    {
        string project = "";
        string template = "";
        string templateType = "";
        string finalTemplate = "";
        string group = "";
        string breakdown = "";
        string HistoricalMeanType = "";
        string HistoricalMeanDescription = "";
        //#pragma warning disable CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used
        string genPath = "";
        //#pragma warning restore CS0414 // The field 'APIWrapper.genPath' is assigned but its value is never used

        private ITaskLog _taskLogging = new TaskLog();

        public APIWrapper() { }

        //public ITaskLog TaskLogging
        //{
        //    get => _taskLogging;
        //    set => _taskLogging = value ?? throw new ArgumentNullException(nameof(TaskLogging));
        //}

        public APIWrapper(string project,
                            string template,
                            string group,
                            string breakdown,
                            string HistoricalMeanType,
                            string HistoricalMeanDescription, string finalTemplate)
        {
            this.project = project;
            this.template = template;
            this.group = group;
            this.breakdown = breakdown;
            this.HistoricalMeanDescription = HistoricalMeanDescription;
            this.HistoricalMeanType = HistoricalMeanType;
            this.finalTemplate = finalTemplate;
        }

        public async Task<string> CallAPI(string project,
                                        string template,
                                        string group,
                                        string breakdown,
                                        string HistoricalMeanType,
                                        string HistoricalMeanDescription)
        {
            using (HttpClient http = new HttpClient())
            {
                const string URL = "https://tools.brandinstitute.com/wsXlCharts/wsExcelCharts.asmx/GetChartBreakDownData";

                //this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");



                var payload = new[]
                {
                new KeyValuePair<string,string>("token","2BF27A11-E318-447A-98FD-70AFE3871AA9"),
                new KeyValuePair<string,string>("project",project),
               new KeyValuePair<string,string>("template",template),
               //new KeyValuePair<string,string>("template",this.templateType),
                new KeyValuePair<string,string>("group",group),
                new KeyValuePair<string,string>("breakdown",breakdown),
                new KeyValuePair<string,string>("historicalMeanType",HistoricalMeanType),
                new KeyValuePair<string,string>("historicalMeanDescription",HistoricalMeanDescription)
            };

                var content = new FormUrlEncodedContent(payload);

                try
                {
                    var response = await http.PostAsync(URL, content);

                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        return $"Error: {response.StatusCode}, {response.ReasonPhrase}";
                    }
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }

        public string CreateTargetPath(string myTemplate)
        {
            string path = $"C:\\excelfiles\\{this.project}";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            path = $"C:\\excelfiles\\{this.project}\\{this.template}_{this.breakdown}.pptx";
            File.Copy(myTemplate, path, true);
            return path;
        }

        public string getChartPathToCopyToFinal(string myTemplate, string project, string breakdown)
        {
            string path = $"C:\\excelfiles\\{project}";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            path = $"C:\\excelfiles\\{project}\\{myTemplate}_{breakdown}.pptx";
            return path;
        }



        public async Task OpenXMLParallelProcess(string project,
                                            List<string> templates,
                                            List<string> breakdowns,
                                            string HistoricalMeanType,
                                            string HistoricalMeanDescription, string finalTemplateName)
        {

            //copy the template folder to the local system

            //copyTemplates();

            string user = "";

            WindowsIdentity identity = WindowsIdentity.GetCurrent();

            user = identity.Name.Replace("BI\\","").ToString();

          //  WindowsIdentity wId = WindowsIdentity.GetCurrent();
          //  var currentUser = wId.User.ToString();

            var userId = _taskLogging.getUserIdByName(user);

            var statusId = _taskLogging.GetStatusIdByName("Queued");

            var taskId = _taskLogging.GetTaskIdByProjectAndUser(project, userId);

            int subtaskId = 0;

            List<Task> taskArr = new List<Task>();
            List<int> guidList = new List<int>();

            foreach (var breakdown in breakdowns)
            {
                foreach (var template in templates)
                {
                    subtaskId = 0;

                    IndividualReportModel taskLog = new IndividualReportModel
                    {
                        TemplateName = template,
                        UserID = userId,
                        StatusID = statusId,
                        CreatedOn = DateTime.Now,
                        StatusMessage = "Task Created",
                       TaskID = taskId
                    };

                    try
                    {
                        //Guid guid = Guid.NewGuid();
                       subtaskId = _taskLogging.InsertIndividualReport(taskLog, user, "Queued");

                        APIWrapper wrapper = new APIWrapper(project, template, template, breakdown, HistoricalMeanType, HistoricalMeanDescription, finalTemplateName);
                        taskArr.Add(Task.Run(() => wrapper.Process()));

                        _taskLogging.UpdateIndividualReport(subtaskId, "Processing","In Process");
                        


                       // _taskLogging.UpdateStatusForIndividualReportTask(taskLog, "Process Running", "Processing");

                        // _taskLogging.InsertTask(taskLog);
                        //_taskLogging.SetTaskStatusState(guid, "Queued", user);

                        //taskLog.CompletedOn = DateTime.UtcNow;
                        //wrapper._taskLogging.SetTaskStatusState(guid, "Done", user);
                        //wrapper._taskLogging.MarkTaskAsCompleted(guid.ToString(), (DateTime)taskLog.CompletedOn, "Done");
                    }
                    catch (Exception ex)
                    {
                        _taskLogging.UpdateIndividualReport(subtaskId, "Fail", ex.Message);

                        //        _taskLogging.UpdateStatusForIndividualReportTask(taskLog, ex.Message, "Failed");

                        //taskLog.CompletedOn = DateTime.UtcNow;
                        //wrapper._taskLogging.SetTaskStatusState(guid, "Fail", user);
                        //wrapper._taskLogging.MarkTaskAsCompleted(guid.ToString(), (DateTime)taskLog.CompletedOn, "Fail");
                        throw new Exception(ex.Message);
                    }
                }
            }

            await Task.WhenAll(taskArr);

            int i = 0;

            foreach(var t in taskArr)
            {
                if (t.Status.ToString() == "RanToCompletion")
                {
                    _taskLogging.UpdateIndividualReport(subtaskId, "Done", "Completed");
                }
                else
                {
                    _taskLogging.UpdateIndividualReport(subtaskId, "Fail", "");
                }
            }

            //foreach (var t in taskArr)
            //{
            //    if (t.Status.ToString() == "RanToCompletion")
            //    {
            //        _taskLogging.SetTaskStatusState(guidList[i], "Done", user);
            //        _taskLogging.MarkTaskAsCompleted(guidList[i].ToString(), DateTime.UtcNow, "Done");
            //    }
            //    else
            //    {
            //        _taskLogging.SetTaskStatusState(guidList[i], "Fail", user);
            //        _taskLogging.MarkTaskAsCompleted(guidList[i].ToString(), DateTime.UtcNow, "Fail");
            //    }
            //    i++;
            //}
        }


        //foreach(var tasks in taskArr)
        //{
        //    if(tasks.Status.ToString() == "RanToCompletion")
        //    {

        //    }
        //}

        //adding the template to the final

    
        



        public Task<clsFinalPageNumberRange> getPPTFinalPageSettings(string strFinalReport, string year, string chartName)
        {

            return Task.Run(() =>
            {

                System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPPTFinalSettings] " + "'" + strFinalReport + "'," + "'" + year + "'");


                List<clsFinalPageModel> lstFinalPageModel = clsData.MRData.ConvertDataTableToListGeneric<clsFinalPageModel>(dt);


                List<clsFinalPageModel> lstFinal = lstFinalPageModel.Where(n => n.strPageGroupName == chartName).ToList();

                clsFinalPageNumberRange objPageNum = new clsFinalPageNumberRange();

                foreach (clsFinalPageModel obj in lstFinal)
                {
                    objPageNum.firstPage = obj.intPPTSlideIndexFirst;
                    objPageNum.lastPage = obj.intPPTSlideIndexLast;

                }

                //  Task<clsFinalPageNumberRange> obj = new Task<clsFinalPageNumberRange>();


                return objPageNum;
            });


        }



        public async Task ProcessMultipleNoParam()
        {
            //APIWrapper wrapper = new APIWrapper("RACKEM", "Fit to Concept", "Fit to Concept", "OVERALL", "Fit to Concept", "2022-HCPS");
            //APIWrapper wrapper3 = new APIWrapper("DESIGNATION_REDO", "Potential For Error - Bar", "Potential For Error - Bar", "OVERALL", "", "");
            //APIWrapper wrapper4 = new APIWrapper("EDAT_2", "Personal Preferences", "Personal Preferences", "OVERALL", "", "");
            //APIWrapper wrapper5 = new APIWrapper("SOLITAIRE", "Likeability", "01 Untrue", "OVERALL", "", "");
            //APIWrapper wrapper6 = new APIWrapper("MEADOW", "Memorability", "Memorability", "OVERALL", "", "");
            //APIWrapper wrapper7 = new APIWrapper("TRIBUTE_C2", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
            //APIWrapper wrapper8 = new APIWrapper("HABITABLE", "Overall Impressions", "Overall Impressions", "Canada Overall", "", "");
            //APIWrapper wrapper9 = new APIWrapper("ZIPLOCK", "Suffix", "Modifier Confusion", "U.S. Overall", "", "");
            //APIWrapper wrapper10 = new APIWrapper("ZIPLOCK", "PromotionalReview", "Modifier Confusion", "U.S. Overall", "", "");
            //APIWrapper wrapper11 = new APIWrapper("MEADOW", "Ease Of Pronounciation", "Att 1", "OVERALL", "", "");
            //APIWrapper wrapper12 = new APIWrapper("MEADOW", "Ease Of Spelling", "Att 1", "OVERALL", "", "");
            //APIWrapper wrapper13 = new APIWrapper("DOMINOS", "03 exagg", "03 exagg", "OVERALL", "", "");
            //APIWrapper wrapper14 = new APIWrapper("DORAEMON", "Innovation", "Innovation", "OVERALL", "", "");
            //APIWrapper wrapper15 = new APIWrapper("ATEMPORAL", "Modifier", "Modifier Meaning (Aided)", "Canada Overall", "", "");
            //APIWrapper wrapper16 = new APIWrapper("QUEENSLAND", "Written Understanding - Bar", "Written Understanding - Bar", "Canada and Europe Overall", "", "");

            //APIWrapper wrapper17 = new APIWrapper("RACKEM", "Attribute 1", "Attribute 1", "OVERALL", "", "");
            //APIWrapper wrapper18 = new APIWrapper("RACKEM", "Attribute 2", "Attribute 2", "OVERALL", "", "");
            //APIWrapper wrapper19 = new APIWrapper("RACKEM", "Attribute 3", "Attribute 3", "OVERALL", "", "");
            //APIWrapper wrapper20 = new APIWrapper("RACKEM", "01 Untrue", "01 Untrue", "OVERALL", "", "");
            //APIWrapper wrapper21 = new APIWrapper("RACKEM", "02 Misleading", "02 Misleading", "OVERALL", "", "");
            //APIWrapper wrapper22 = new APIWrapper("RACKEM", "03 Exagg", "03 Exagg", "OVERALL", "", "");
            //APIWrapper wrapper24 = new APIWrapper("RACKEM", "Memorability", "Memorability", "OVERALL", "", "");
            //APIWrapper wrapper25 = new APIWrapper("RACKEM", "Overall Impressions", "Overall Impressions", "OVERALL", "", "");
            //APIWrapper wrapper26 = new APIWrapper("RACKEM", "Verbal Understanding - Bar", "Verbal Understanding - Bar", "OVERALL", "", "");
            //APIWrapper wrapper27 = new APIWrapper("RACKEM", "Written Understanding - Bar", "Written Understanding - Bar", "Overall", "", "");
            ////APIWrapper wrapper28 = new APIWrapper("RACKEM", "Potential For Error - Bar", "Potential For Error - Bar", "Overall", "", "");
            //APIWrapper wrapper29 = new APIWrapper("VGT_BND", "Preference Ranking", "Preference Ranking", "Overall", "", "");
            //APIWrapper wrapper30 = new APIWrapper("RACKEM", "Attribute evaluation Aggregate", "Attribute evaluation Aggregate", "Overall", "", "");
            //APIWrapper wrapper31 = new APIWrapper("VAYNSMR_HEME", "Initial Recall", "Initial Recall", "Canada and EU Medical Professionals", "", "");
            //APIWrapper wrapper32 = new APIWrapper("ALL_FAME", "Chemical Structure Appropriateness", "", "Overall", "", "");

            //APIWrapper wrapper33 = new APIWrapper("COPEMVIBO", "Likeability Rationale", "Likeability Rationale", "Overall", "", "");


            Task[] taskArr = new Task[]
            {
                //Task.Run(() => wrapper.Process()),
                //Task.Run(() => wrapper3.Process()),
                //Task.Run(() => wrapper4.Process()),
                //Task.Run(() => wrapper5.Process()),
                //Task.Run(() => wrapper6.Process()),
                //Task.Run(() => wrapper7.Process()),
                //Task.Run(() => wrapper8.Process()),
                //Task.Run(() => wrapper9.Process()),
                //Task.Run(() => wrapper10.Process()),
                //Task.Run(() => wrapper11.Process()),
                //Task.Run(() => wrapper12.Process()),
                //Task.Run(() => wrapper13.Process()),
                //Task.Run(() => wrapper14.Process()),
                //Task.Run(() => wrapper15.Process()),
                //Task.Run(() => wrapper16.Process()),

                // Task.Run(() => wrapper17.Process()),
                // Task.Run(() => wrapper18.Process()),
                // Task.Run(() => wrapper19.Process()),
                // Task.Run(() => wrapper20.Process()),
                // Task.Run(() => wrapper21.Process()),
                // Task.Run(() => wrapper22.Process()),
                // Task.Run(() => wrapper24.Process()),
                // Task.Run(() => wrapper25.Process()),
                // Task.Run(() => wrapper26.Process()),
                // Task.Run(() => wrapper27.Process()),
                //Task.Run(() => wrapper28.Process()),

                //Task.Run(() => wrapper29.Process()),
                //Task.Run(() => wrapper30.Process()),
                //Task.Run(() => wrapper31.Process()),
                //Task.Run(() => wrapper32.Process()),
                //Task.Run(() => wrapper33.Process())
            };

            await Task.WhenAll(taskArr);
        }

        public async Task Process()
        {

            //get the pageGroup

            this.templateType = await clsData.MRData.getStrValue("getPageGroupType " + "'" + project + "'," + "'" + template + "'");


            string sourcePath = "";
            var apiCall = await CallAPI(project, template, group, breakdown, HistoricalMeanType, HistoricalMeanDescription);

            DLLClass dLLClass = new DLLClass();
            XDocument xDoc = XDocument.Parse(apiCall);

            var data = xDoc.Root.Value;

            var jsonSettings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };


            if (data.Length == 0)
            {
                Console.WriteLine("No data present in api");
                return;
            }

            if (data == null || data.ToString() == "[]")
            {
                sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                notAvailable(CreateTargetPath(sourcePath), template, breakdown);
            }


            if (data.Length != 0)
            {
                switch (this.templateType)
                {
                    case "Fit to Concept":


                        List<DLLTemplate.FitToConceptModel> fitToConceptData = new List<DLLTemplate.FitToConceptModel>();

                        fitToConceptData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


                        if (fitToConceptData.Count == 0)
                        {
                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            //CreateTargetPath(sourcePath);

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FittoConcept" + fitToConceptData.Count + ".pptx";
                            dLLClass.FitToConceptMethod(CreateTargetPath(sourcePath), fitToConceptData, HistoricalMeanType, HistoricalMeanDescription);
                        }
                        break;

                    case "Attribute 1":
                    case "Attribute 2":
                    case "Attribute 3":
                    case "Att 1":
                    case "Att 2":
                    case "Att 3":
                    case "Attribute Evaluation":



                        List<DLLTemplate.FitToConceptModel> Att1Data = new List<DLLTemplate.FitToConceptModel>();

                        Att1Data = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

                        if (Att1Data.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                            //update the notavailable 


                        }
                        else
                        {


                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluation" + Att1Data.Count + ".pptx";
                            dLLClass.AttributeMethod(CreateTargetPath(sourcePath), Att1Data, group, HistoricalMeanType, HistoricalMeanDescription, project);


                        }
                        break;

                    case "Attribute evaluation Aggregate":
                    case "Attribute Evaluation Aggregate":

                        List<DLLTemplate.FitToConceptModel> attEvalData = new List<DLLTemplate.FitToConceptModel>();

                        attEvalData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


                        if (attEvalData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\AttributeEvaluationAggregate" + attEvalData.Count + ".pptx";
                            dLLClass.AttributeMethodForAttributeEvalAggreg(CreateTargetPath(sourcePath), attEvalData);
                        }
                        break;



                    case "BRANDEX SUFFIX STRATEGIC":
                        List<BrandexSuffixStrategicModel> brandexSuffixStrategicData = new List<BrandexSuffixStrategicModel>();
                        brandexSuffixStrategicData = JsonConvert.DeserializeObject<List<BrandexSuffixStrategicModel>>(data);

                        if (brandexSuffixStrategicData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            List<BrandexSuffixStrategicShortModel> brandexSuffixStrategicShortModel = new List<BrandexSuffixStrategicShortModel>();

                            double dblAverage1Max = 0.0;
                            double dblAverage2Max = 0.0;

                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
                            {
                                var dblAverage1MaxValue = brandexSuffixStrategicData[i].dblAveragePage1;
                                if (dblAverage1MaxValue == 0)
                                {
                                    dblAverage1Max = 0.0;
                                }
                                if (dblAverage1MaxValue > dblAverage1Max)
                                {
                                    dblAverage1Max = dblAverage1MaxValue;
                                }

                                var dblAverage2MaxValue = brandexSuffixStrategicData[i].dblAveragePage2;
                                if (dblAverage2MaxValue == 0)
                                {
                                    dblAverage2Max = 0;
                                }
                                if (dblAverage2MaxValue > dblAverage2Max)
                                {
                                    dblAverage2Max = dblAverage2MaxValue;
                                }
                            }

                            // scaling factor 
                            double scalingFactorForStrategicDistinctiveness = 1.0;

                            for (int i = 0; i < brandexSuffixStrategicData.Count; i++)
                            {
                                // for the table
                                BrandexSuffixStrategicShortModel brandexSuffixStrategicModelData = new BrandexSuffixStrategicShortModel();

                                var dataForBrandexSuffix = brandexSuffixStrategicData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValue =
                                        (dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * dataForBrandexSuffix.dblPage1Weight;
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValue =
                                       (dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * dataForBrandexSuffix.dblPage2Weight;
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue) * scalingFactorForStrategicDistinctiveness;

                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;

                                brandexSuffixStrategicModelData.dblAveragePage1 = dataForBrandexSuffix.dblAveragePage1;
                                brandexSuffixStrategicModelData.dblPage1Weight = dataForBrandexSuffix.dblPage1Weight;
                                brandexSuffixStrategicModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

                                brandexSuffixStrategicModelData.dblAveragePage2 = dataForBrandexSuffix.dblAveragePage2;
                                brandexSuffixStrategicModelData.dblPage2Weight = dataForBrandexSuffix.dblPage2Weight;
                                brandexSuffixStrategicModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

                                brandexSuffixStrategicModelData.dblAveragePage3 = dataForBrandexSuffix.dblAveragePage3;
                                brandexSuffixStrategicModelData.dblPage3Weight = dataForBrandexSuffix.dblPage3Weight;

                                brandexSuffixStrategicModelData.dblAveragePage4 = dataForBrandexSuffix.dblAveragePage4;
                                brandexSuffixStrategicModelData.dblPage4Weight = dataForBrandexSuffix.dblPage4Weight;

                                brandexSuffixStrategicModelData.dblAveragePage5 = dataForBrandexSuffix.dblAveragePage5;
                                brandexSuffixStrategicModelData.dblPage5Weight = dataForBrandexSuffix.dblPage5Weight;
                                brandexSuffixStrategicModelData.dblAveragePage5Weighted = dataForBrandexSuffix.dblAveragePage5;

                                brandexSuffixStrategicModelData.dblAveragePage6 = dataForBrandexSuffix.dblAveragePage6;
                                brandexSuffixStrategicModelData.dblPage6Weight = dataForBrandexSuffix.dblPage6Weight;
                                brandexSuffixStrategicModelData.dblAveragePage6Weighted = dataForBrandexSuffix.dblAveragePage6Weighted;

                                brandexSuffixStrategicModelData.dblAveragePage7 = dataForBrandexSuffix.dblAveragePage7;
                                brandexSuffixStrategicModelData.dblPage7Weight = dataForBrandexSuffix.dblPage7Weight;
                                brandexSuffixStrategicModelData.dblAveragePage7Weighted = dataForBrandexSuffix.dblAveragePage7Weighted;

                                brandexSuffixStrategicModelData.dblAveragePage8 = dataForBrandexSuffix.dblAveragePage8;
                                brandexSuffixStrategicModelData.dblPage8Weight = dataForBrandexSuffix.dblPage8Weight;
                                brandexSuffixStrategicModelData.dblAveragePage8Weighted = dataForBrandexSuffix.dblAveragePage8Weighted;

                                brandexSuffixStrategicModelData.dblIndex = indexSum;
                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

                                // for chart
                                int memorabilityStrategicWeightPage = 50;
                                int personalPreferenceStrategicWeightPage = 50;

                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage1 / dblAverage1Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = ((dataForBrandexSuffix.dblAveragePage2 / dblAverage2Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                double indexSumForChart = averagePage1WeightedValueForChart +
                                                          averagePage2WeightedValueForChart;

                                brandexSuffixStrategicModelData.strTestName = dataForBrandexSuffix.strTestName;
                                brandexSuffixStrategicModelData.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;
                                brandexSuffixStrategicModelData.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

                                brandexSuffixStrategicModelData.dblIndexForChart = indexSumForChart;

                                brandexSuffixStrategicModelData.strDSIScore = dataForBrandexSuffix.strDSIScore;
                                brandexSuffixStrategicModelData.intRed = dataForBrandexSuffix.intRed;
                                brandexSuffixStrategicModelData.intGreen = dataForBrandexSuffix.intGreen;
                                brandexSuffixStrategicModelData.intBlue = dataForBrandexSuffix.intBlue;
                                brandexSuffixStrategicModelData.boolBold = dataForBrandexSuffix.boolBold;

                                brandexSuffixStrategicShortModel.Add(brandexSuffixStrategicModelData);
                            }

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSuffixStrategicTemplateModified" + brandexSuffixStrategicShortModel.Count + ".pptx";
                            dLLClass.BrandexSuffixStrategicMethod(CreateTargetPath(sourcePath), brandexSuffixStrategicShortModel);
                        }
                        break;





                    //case "Initial Recall":
                    //    List<DLLTemplate.FitToConceptNewModel> initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();
                    //    initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data);
                    //    if (initialRecallData.Count == 0)
                    //    {
                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                    //    }
                    //    else
                    //    {
                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + initialRecallData.Count + ".pptx";
                    //        dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), initialRecallData, this.HistoricalMeanType, this.HistoricalMeanDescription);
                    //    }
                    //    break;





                    case "Brandex Strategic Distinctiveness":

                        List<BrandexStrategicDistinctivenessModel> brandexStrategicDistinctivenessesData = new List<BrandexStrategicDistinctivenessModel>();
                        brandexStrategicDistinctivenessesData = JsonConvert.DeserializeObject<List<BrandexStrategicDistinctivenessModel>>(data);

                        if (brandexStrategicDistinctivenessesData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            List<BrandexStrategicDistinctivenessShortModel> brandexStrategicDistinctivenessesShortData = new List<BrandexStrategicDistinctivenessShortModel>();

                            double dblAverage1Max = 0.0;
                            double dblAverage2Max = 0.0;
                            double dblAverage3Max = 0.0;
                            double dblAverage4Max = 0.0;

                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
                            {
                                var dblAverage1MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage1;
                                if (dblAverage1MaxValue == 0)
                                {
                                    dblAverage1Max = 0.0;
                                }
                                if (dblAverage1MaxValue > dblAverage1Max)
                                {
                                    dblAverage1Max = dblAverage1MaxValue;
                                }

                                var dblAverage2MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage2;
                                if (dblAverage2MaxValue == 0)
                                {
                                    dblAverage2Max = 0;
                                }
                                if (dblAverage2MaxValue > dblAverage2Max)
                                {
                                    dblAverage2Max = dblAverage2MaxValue;
                                }

                                var dblAverage3MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage3;
                                if (dblAverage3MaxValue == 0)
                                {
                                    dblAverage3Max = 0;
                                }
                                if (dblAverage3MaxValue > dblAverage3Max)
                                {
                                    dblAverage3Max = dblAverage3MaxValue;
                                }

                                var dblAverage4MaxValue = brandexStrategicDistinctivenessesData[i].dblAveragePage4;
                                if (dblAverage4MaxValue == 0)
                                {
                                    dblAverage4Max = 0;
                                }
                                if (dblAverage4MaxValue > dblAverage4Max)
                                {
                                    dblAverage4Max = dblAverage4MaxValue;
                                }
                            }

                            // scaling factor 
                            double scalingFactorForStrategicDistinctiveness = 1.00777;

                            for (int i = 0; i < brandexStrategicDistinctivenessesData.Count; i++)
                            {
                                // for the table
                                BrandexStrategicDistinctivenessShortModel brandexStrategicDistinctivenessModelData = new BrandexStrategicDistinctivenessShortModel();

                                var dataForDistinctivessModel = brandexStrategicDistinctivenessesData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;
                                double averagePage3WeightedValue = 0.0;
                                double averagePage4WeightedValue = 0.0;

                                if (dblAverage1Max > 0)
                                {
                                    averagePage1WeightedValue =
                                        (dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * dataForDistinctivessModel.dblPage1Weight;
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValue =
                                       (dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * dataForDistinctivessModel.dblPage2Weight;
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValue =
                                        (dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * dataForDistinctivessModel.dblPage3Weight;
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValue =
                                        (dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * dataForDistinctivessModel.dblPage4Weight;
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                                             averagePage4WeightedValue) * scalingFactorForStrategicDistinctiveness;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;

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
                                    averagePage1WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValueForChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationStrategicWeightPage) * scalingFactorForStrategicDistinctiveness;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                double indexSumForChartDistinctiveness = averagePage1WeightedValueForChart +
                                                          averagePage2WeightedValueForChart +
                                                          averagePage3WeightedValueForChart +
                                                          averagePage4WeightedValueForChart;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
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
                                    averagePage1WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage1 / dblAverage1Max) * fitToConceptDistinctivenessWeightPage) * scalingFactorForMarketingChart;
                                }
                                else
                                {
                                    averagePage1WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage2 / dblAverage2Max) * memorabilityDistinctivenssWeightPage) * scalingFactorForMarketingChart;
                                }
                                else
                                {
                                    averagePage2WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage3Max > 0)
                                {
                                    averagePage3WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage3 / dblAverage3Max) * personalPreferenceDistinctivenessWeightPage) * scalingFactorForMarketingChart;
                                }
                                else
                                {
                                    averagePage3WeightedValueForMarketingChart = 0;
                                }

                                if (dblAverage4Max > 0)
                                {
                                    averagePage4WeightedValueForMarketingChart = ((dataForDistinctivessModel.dblAveragePage4 / dblAverage4Max) * attributeEvaluationDistinctivenessWeightPage) * scalingFactorForMarketingChart;
                                }
                                else
                                {
                                    averagePage4WeightedValueForMarketingChart = 0;
                                }

                                double indexSumForChartMarketing = averagePage1WeightedValueForMarketingChart +
                                                          averagePage2WeightedValueForMarketingChart +
                                                          averagePage3WeightedValueForMarketingChart +
                                                          averagePage4WeightedValueForMarketingChart;

                                brandexStrategicDistinctivenessModelData.strTestName = dataForDistinctivessModel.strTestName;
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

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexStrategicDistinctiveness" + brandexStrategicDistinctivenessesData.Count + ".pptx";
                            dLLClass.BrandexStrategicDistinctivenessMethod(CreateTargetPath(sourcePath), brandexStrategicDistinctivenessesShortData);
                        }




                        break;


                    //case "Brandex Safety":
                    //    List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
                    //    brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

                    //    List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

                    //    if (brandexSafetyData.Count == 0)
                    //    {
                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                    //        dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                    //    }
                    //    else
                    //    {
                    //        double average1Max = 0.0000;
                    //        double average2Max = 0.0000;
                    //        double average3Max = 0.0;
                    //        double average4Max = 0.0000;
                    //        double average5Max = 0.0000;

                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
                    //        {
                    //            var dblAverage1max = brandexSafetyData[i].dblAveragePage1;
                    //            if (dblAverage1max == 0.0)
                    //            {
                    //                average1Max = 0;
                    //            }
                    //            if (average1Max < dblAverage1max)
                    //            {
                    //                average1Max = dblAverage1max;
                    //            }

                    //            var dblAverage2max = brandexSafetyData[i].dblAveragePage2;
                    //            if (dblAverage2max == 0.0)
                    //            {
                    //                average2Max = 0;
                    //            }
                    //            if (average2Max < dblAverage2max)
                    //            {
                    //                average2Max = dblAverage2max;
                    //            }

                    //            var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
                    //            if (dblAverage3max == 0.0)
                    //            {
                    //                average3Max = 0;
                    //            }
                    //            if (average3Max < dblAverage3max)
                    //            {
                    //                average3Max = dblAverage3max;
                    //            }

                    //            var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
                    //            if (dblAverage4max == 0.0)
                    //            {
                    //                average4Max = 0;
                    //            }
                    //            if (average4Max < dblAverage4max)
                    //            {
                    //                average4Max = dblAverage4max;
                    //            }

                    //            var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
                    //            if (dblAverage5max == 0.0)
                    //            {
                    //                average5Max = 0;
                    //            }
                    //            if (average5Max < dblAverage5max)
                    //            {
                    //                average5Max = dblAverage5max;
                    //            }
                    //        }

                    //        double scalingFactor = 0.750610000;

                    //        for (int i = 0; i < brandexSafetyData.Count; i++)
                    //        {
                    //            //for the table
                    //            BrandexSafetyShortModel brandexSafetyShortModel = new BrandexSafetyShortModel();

                    //            var dataEl = brandexSafetyData[i];

                    //            double averagePage1WeightedValue = 0.0;
                    //            double averagePage2WeightedValue = 0.0;
                    //            double averagePage3WeightedValue = 0.0;
                    //            double averagePage4WeightedValue = 0.0;
                    //            double averagePage5WeightedValue = 0.0;

                    //            if (average1Max > 0)
                    //            {
                    //                averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
                    //            }
                    //            else
                    //            {
                    //                averagePage1WeightedValue = 0;
                    //            }

                    //            if (average2Max > 0)
                    //            {
                    //                averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
                    //            }
                    //            else
                    //            {
                    //                averagePage2WeightedValue = 0;
                    //            }

                    //            if (average3Max > 0)
                    //            {
                    //                averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
                    //            }
                    //            else
                    //            {
                    //                averagePage3WeightedValue = 0;
                    //            }

                    //            if (average4Max > 0)
                    //            {
                    //                averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
                    //            }
                    //            else
                    //            {
                    //                averagePage4WeightedValue = 0;
                    //            }

                    //            if (average5Max > 0)
                    //            {
                    //                averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
                    //            }
                    //            else
                    //            {
                    //                averagePage5WeightedValue = 0;
                    //            }


                    //            //double averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight;
                    //            //double averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;
                    //            //double averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight;
                    //            // double averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;
                    //            //double averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;

                    //            double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                    //                              averagePage4WeightedValue + averagePage5WeightedValue) * scalingFactor;

                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

                    //            brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
                    //            brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
                    //            brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

                    //            brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
                    //            brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
                    //            brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

                    //            brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
                    //            brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
                    //            brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

                    //            brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
                    //            brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
                    //            brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

                    //            brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
                    //            brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
                    //            brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

                    //            brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
                    //            brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
                    //            brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

                    //            brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
                    //            brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
                    //            brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

                    //            brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
                    //            brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
                    //            brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

                    //            brandexSafetyShortModel.dblIndex = indexSum;
                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


                    //            //for the chart - 

                    //            double averagePage1WeightedValueForChart = 0.0;
                    //            double averagePage2WeightedValueForChart = 0.0;
                    //            double averagePage3WeightedValueForChart = 0.0;
                    //            double averagePage4WeightedValueForChart = 0.0;
                    //            double averagePage5WeightedValueForChart = 0.0;

                    //            if (average1Max > 0)
                    //            {
                    //                averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
                    //            }
                    //            else
                    //            {
                    //                averagePage1WeightedValueForChart = 0;
                    //            }

                    //            if (average2Max > 0)
                    //            {
                    //                averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
                    //            }
                    //            else
                    //            {
                    //                averagePage2WeightedValueForChart = 0;
                    //            }

                    //            if (average3Max > 0)
                    //            {
                    //                averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
                    //            }
                    //            else
                    //            {
                    //                averagePage3WeightedValueForChart = 0;
                    //            }

                    //            if (average4Max > 0)
                    //            {
                    //                averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
                    //            }
                    //            else
                    //            {
                    //                averagePage4WeightedValueForChart = 0;
                    //            }

                    //            if (average5Max > 0)
                    //            {
                    //                averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
                    //            }
                    //            else
                    //            {
                    //                averagePage5WeightedValueForChart = 0;
                    //            }

                    //            //double averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
                    //            //double averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
                    //            //double averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
                    //            //double averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
                    //            //double averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;

                    //            double indexSumForChart = averagePage1WeightedValueForChart +
                    //                                      averagePage2WeightedValueForChart +
                    //                                      averagePage3WeightedValueForChart +
                    //                                      averagePage4WeightedValueForChart +
                    //                                      averagePage5WeightedValueForChart;

                    //            brandexSafetyShortModel.strTestName = dataEl.strTestName;

                    //            brandexSafetyShortModel.dblAveragePage1WeightedForChart = averagePage1WeightedValueForChart;

                    //            brandexSafetyShortModel.dblAveragePage2WeightedForChart = averagePage2WeightedValueForChart;

                    //            brandexSafetyShortModel.dblAveragePage3WeightedForChart = averagePage3WeightedValueForChart;

                    //            brandexSafetyShortModel.dblAveragePage4WeightedForChart = averagePage4WeightedValueForChart;

                    //            brandexSafetyShortModel.dblAveragePage5WeightedForChart = averagePage5WeightedValueForChart;
                    //            ;
                    //            brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
                    //            brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                    //            brandexSafetyShortModel.intRed = dataEl.intRed;
                    //            brandexSafetyShortModel.intGreen = dataEl.intGreen;
                    //            brandexSafetyShortModel.intBlue = dataEl.intBlue;
                    //            brandexSafetyShortModel.boolBold = dataEl.boolBold;


                    //            brandexData.Add(brandexSafetyShortModel);
                    //        }

                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
                    //        dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
                    //    }
                    //    break;


                    case "Brandex Safety":
                        List<BrandexSafetyModel> brandexSafetyData = new List<BrandexSafetyModel>();
                        brandexSafetyData = JsonConvert.DeserializeObject<List<BrandexSafetyModel>>(data);

                        List<BrandexSafetyShortModel> brandexData = new List<BrandexSafetyShortModel>();

                        if (brandexSafetyData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
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
                                if (dblAverage2max == 0.0)
                                {
                                    average2Max = 0;
                                }
                                if (average2Max < dblAverage2max)
                                {
                                    average2Max = (double)dblAverage2max;
                                }

                                var dblAverage3max = brandexSafetyData[i].dblAveragePage3;
                                if (dblAverage3max == 0.0)
                                {
                                    average3Max = 0;
                                }
                                if (average3Max < dblAverage3max)
                                {
                                    average3Max = (double)dblAverage3max;
                                }

                                var dblAverage4max = brandexSafetyData[i].dblAveragePage4;
                                if (dblAverage4max == 0.0)
                                {
                                    average4Max = 0;
                                }
                                if (average4Max < dblAverage4max)
                                {
                                    average4Max = (double)dblAverage4max;
                                }

                                var dblAverage5max = brandexSafetyData[i].dblAveragePage5;
                                if (dblAverage5max == 0.0)
                                {
                                    average5Max = 0;
                                }
                                if (average5Max < dblAverage5max)
                                {
                                    average5Max = (double)dblAverage5max;
                                }
                            }

                            double scalingFactor = 0.750610000;

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

                                    averagePage1WeightedValue = (double)((double)(dataEl.dblAveragePage1 / average1Max) * (double)(dataEl.dblPage1Weight));
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (average2Max > 0)
                                {
                                    // averagePage2WeightedValue = (double)(dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight;

                                    averagePage2WeightedValue = (double)((double)(dataEl.dblAveragePage2 / average2Max) * (double)(dataEl.dblPage2Weight));
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (average3Max > 0)
                                {
                                    averagePage3WeightedValue = (double)((double)(dataEl.dblAveragePage3 / average3Max) * (double)(dataEl.dblPage3Weight));
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (average4Max > 0)
                                {
                                    //averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight;

                                    averagePage4WeightedValue = (double)((double)(dataEl.dblAveragePage4 / average4Max) * (double)(dataEl.dblPage4Weight));
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                if (average5Max > 0)
                                {
                                    // averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight;
                                    averagePage5WeightedValue = (double)((double)(dataEl.dblAveragePage5 / average5Max) * (double)(dataEl.dblPage5Weight));


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
                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7 ?? 0.0;
                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8 ?? 0.0;
                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight ?? 0.0;
                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

                                brandexSafetyShortModel.dblIndex = indexSum;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = dataEl.intRed;
                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


                                //for the chart - 

                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;
                                double averagePage3WeightedValueForChart = 0.0;
                                double averagePage4WeightedValueForChart = 0.0;
                                double averagePage5WeightedValueForChart = 0.0;

                                if (average1Max > 0)
                                {
                                    averagePage1WeightedValueForChart = (double)((dataEl.dblAveragePage1 / average1Max) * dataEl.dblPage1Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (average2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = (double)((dataEl.dblAveragePage2 / average2Max) * dataEl.dblPage2Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (average3Max > 0)
                                {
                                    averagePage3WeightedValueForChart = (double)((dataEl.dblAveragePage3 / average3Max) * dataEl.dblPage3Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (average4Max > 0)
                                {
                                    averagePage4WeightedValueForChart = (double)((dataEl.dblAveragePage4 / average4Max) * dataEl.dblPage4Weight) * scalingFactor;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                if (average5Max > 0)
                                {
                                    averagePage5WeightedValueForChart = (double)((dataEl.dblAveragePage5 / average5Max) * dataEl.dblPage5Weight) * scalingFactor;
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
                                ;
                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = dataEl.intRed;
                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


                                brandexData.Add(brandexSafetyShortModel);
                            }

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexSafetySpecific" + brandexSafetyData.Count + ".pptx";
                            dLLClass.BrandexSafetyMethod(CreateTargetPath(sourcePath), brandexData, this.breakdown);
                        }
                        break;

                    case "Potential For Error - Bar":
                    case "Potential For Error":


                        List<DLLTemplate.FitToConceptModel> potentialForError = new List<DLLTemplate.FitToConceptModel>();

                        potentialForError = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);


                        if (potentialForError.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PotentialForError-Bar" + potentialForError.Count + ".pptx";
                            dLLClass.PotentialForErrorBarMethod(CreateTargetPath(sourcePath), potentialForError);
                        }
                        break;

                    case "Personal Preferences":


                        List<DLLTemplate.FitToConceptModel> personalPref = new List<DLLTemplate.FitToConceptModel>();


                        personalPref = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

                        if (personalPref.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            //lok needs a fix RADIANT : error (Personal Preference testname null)

                            // notAvailable(this.template);
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PersonalPreferences" + personalPref.Count + ".pptx";
                            dLLClass.PersonalPreferencesMethod(CreateTargetPath(sourcePath), personalPref, HistoricalMeanType, HistoricalMeanDescription);
                        }
                        break;

                    case "Likeability":


                        List<DLLTemplate.OverallImpressionModel> likeability = new List<DLLTemplate.OverallImpressionModel>();
                        likeability = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

                        if (likeability.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Likeability" + likeability.Count + ".pptx";
                            dLLClass.LikeabilityMethod(CreateTargetPath(sourcePath), likeability);
                        }
                        break;



                    case "Likeability Rationale":
                        List<LikeabilityRationaleModel> likeabilyRationalesData = new List<LikeabilityRationaleModel>();
                        likeabilyRationalesData = JsonConvert.DeserializeObject<List<LikeabilityRationaleModel>>(data);

                        if (likeabilyRationalesData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationale.pptx";
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\LikeabilityRationaleOrdered.pptx";
                            dLLClass.LikeabilityRationaleMethod(CreateTargetPath(sourcePath), likeabilyRationalesData, breakdown);
                        }
                        break;

                    case "Memorability":


                        List<DLLTemplate.FitToConceptModel> memorability = new List<DLLTemplate.FitToConceptModel>();

                        memorability = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

                        if (memorability.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Memorability" + memorability.Count + ".pptx";
                            dLLClass.MemorabilityMethod(CreateTargetPath(sourcePath), memorability, HistoricalMeanType, HistoricalMeanDescription);
                        }
                        break;

                    case "Verbal Understanding - Bar":
                    case "Verbal Understanding":


                        List<DLLTemplate.FitToConceptModel> verbalUnde = new List<DLLTemplate.FitToConceptModel>();


                        verbalUnde = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

                        if (verbalUnde.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\VerbalUnderstanding-Bar" + verbalUnde.Count + ".pptx";
                            dLLClass.VerbalUnderstandingBarMethod(CreateTargetPath(sourcePath), verbalUnde);
                        }
                        break;

                    case "Overall Impressions":


                        List<DLLTemplate.OverallImpressionNewModel> overallImpression = new List<DLLTemplate.OverallImpressionNewModel>();



                        overallImpression = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

                        if (overallImpression.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\OverallImpressions" + overallImpression.Count + ".pptx";
                            dLLClass.OverallImpressionsMethod(CreateTargetPath(sourcePath), overallImpression);
                        }
                        break;

                    case "Suffix":
                    case "Suffix Meaning":
                    case "Existing Abbreviation ID":
                    case "Existing Suffix ID":


                        List<DLLTemplate.OverallImpressionModel> suffixOverallImpressionData = new List<DLLTemplate.OverallImpressionModel>();

                        suffixOverallImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

                        if (suffixOverallImpressionData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Suffix" + suffixOverallImpressionData.Count + ".pptx";
                            dLLClass.SuffixMethod(CreateTargetPath(sourcePath), suffixOverallImpressionData, breakdown);
                        }
                        break;

                    case "PromotionalReview":

                        List<DLLTemplate.OverallImpressionModel> promotionalImpressionData = new List<DLLTemplate.OverallImpressionModel>();

                        promotionalImpressionData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionModel>>(data, jsonSettings);

                        if (promotionalImpressionData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PromotionalReview" + promotionalImpressionData.Count + ".pptx";
                            dLLClass.PromotionalReviewMethod(CreateTargetPath(sourcePath), promotionalImpressionData);
                        }
                        break;

                    case "Ease Of Pronounciation":
                    case "Ease of Pronunciation":



                        List<DLLTemplate.FitToConceptModel> easeOfPronounicationData = new List<DLLTemplate.FitToConceptModel>();
                        easeOfPronounicationData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptModel>>(data, jsonSettings);

                        if (easeOfPronounicationData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfPronunciation" + easeOfPronounicationData.Count + ".pptx";
                            dLLClass.EaseOfPronounicationMethod(CreateTargetPath(sourcePath), easeOfPronounicationData, HistoricalMeanType, HistoricalMeanDescription);
                        }
                        break;

                    case "Ease Of Spelling":

                        List<DLLTemplate.FitToConceptNewModel> easeOfSpellingData = new List<DLLTemplate.FitToConceptNewModel>();

                        easeOfSpellingData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

                        if (easeOfSpellingData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\EaseOfSpelling" + easeOfSpellingData.Count + ".pptx";
                            dLLClass.EaseOfSpellingMethod(CreateTargetPath(sourcePath), easeOfSpellingData);
                        }

                        break;

                    case "03 Exagg":
                    case "01 Untrue":
                    case "02 Misleading":
                    case "02 Mislead":
                    case "03 Exaggerative":
                    case "Exaggerative-Inappropriate":


                        List<DLLTemplate.OverallImpressionNewModel> ExaggData = new List<DLLTemplate.OverallImpressionNewModel>();
                        ExaggData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);
                        if (ExaggData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            Console.WriteLine(ExaggData);

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Exaggerative" + ExaggData.Count + ".pptx";
                            dLLClass.ExaggerativeMethod(CreateTargetPath(sourcePath), ExaggData, this.template);
                        }
                        break;


                    case "Initial Recall":

                        List<DLLTemplate.FitToConceptNewModel> Vaynsmr_heme_initialRecallData = new List<DLLTemplate.FitToConceptNewModel>();

                        Vaynsmr_heme_initialRecallData = JsonConvert.DeserializeObject<List<DLLTemplate.FitToConceptNewModel>>(data, jsonSettings);

                        if (Vaynsmr_heme_initialRecallData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\InitialRecall" + Vaynsmr_heme_initialRecallData.Count + ".pptx";
                            dLLClass.InitialRecallMethod(CreateTargetPath(sourcePath), Vaynsmr_heme_initialRecallData, HistoricalMeanType, HistoricalMeanDescription);
                        }
                        break;

                    case "Innovation":
                        List<DLLTemplate.InnovationModel> doraemonInnovationData = new List<DLLTemplate.InnovationModel>();
                        doraemonInnovationData = JsonConvert.DeserializeObject<List<DLLTemplate.InnovationModel>>(data);
                        if (doraemonInnovationData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Innovation" + doraemonInnovationData.Count + ".pptx";
                            dLLClass.InnovationMethod(CreateTargetPath(sourcePath), doraemonInnovationData, this.HistoricalMeanType, this.HistoricalMeanDescription);
                        }
                        break;


                    case "Modifier":

                        List<DLLTemplate.OverallImpressionNewModel> atemporalModifierData = new List<DLLTemplate.OverallImpressionNewModel>();

                        atemporalModifierData = JsonConvert.DeserializeObject<List<DLLTemplate.OverallImpressionNewModel>>(data, jsonSettings);

                        if (atemporalModifierData.Count == 0)
                        {

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Modifier" + atemporalModifierData.Count + ".pptx";
                            dLLClass.ModifierMethod(CreateTargetPath(sourcePath), atemporalModifierData, group);
                        }
                        break;

                    case "Written Understanding - Bar":
                    case "Written Understanding":


                        List<DLLTemplate.WrittenUnderstadingModel> writtenUnde = new List<DLLTemplate.WrittenUnderstadingModel>();

                        writtenUnde = JsonConvert.DeserializeObject<List<DLLTemplate.WrittenUnderstadingModel>>(data, jsonSettings);

                        if (writtenUnde.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\WrittenUnderstanding" + writtenUnde.Count + ".pptx";
                            dLLClass.WrittenUnderstandingMethod(CreateTargetPath(sourcePath), writtenUnde);
                        }
                        break;




                    case "Preference Ranking":
                        List<PreferenceRankingModel> preferenceRankingData = new List<PreferenceRankingModel>();
                        preferenceRankingData = JsonConvert.DeserializeObject<List<PreferenceRankingModel>>(data);

                        if (preferenceRankingData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            List<int> counterList = new List<int>();

                            int count = 0;

                            for (int j = 0; j < preferenceRankingData.Count; j++)
                            {
                                var apiDataEl = preferenceRankingData[j];
                                if (apiDataEl.intRankCount1 != 0)
                                {
                                    count = 1;
                                }
                                if (apiDataEl.intRankCount2 != 0)
                                {
                                    count = 2;
                                }
                                if (apiDataEl.intRankCount3 != 0)
                                {
                                    count = 3;
                                }
                                if (apiDataEl.intRankCount4 != 0)
                                {
                                    count = 4;
                                }
                                if (apiDataEl.intRankCount5 != 0)
                                {
                                    count = 5;
                                }
                                if (apiDataEl.intRankCount6 != 0)
                                {
                                    count = 6;
                                }
                                if (apiDataEl.intRankCount7 != 0)
                                {
                                    count = 7;
                                }
                                if (apiDataEl.intRankCount8 != 0)
                                {
                                    count = 8;
                                }
                                if (apiDataEl.intRankCount9 != 0)
                                {
                                    count = 9;
                                }
                                counterList.Add(count);
                            }

                            List<PreferenceRankingSortedDataModel> apiDataWithWeightedList = new List<PreferenceRankingSortedDataModel>();

                            List<int> lstWeightedData = new List<int>();

                            int weightedSum = 0;

                            for (int j = 0; j < preferenceRankingData.Count; j++)
                            {
                                PreferenceRankingSortedDataModel preferenceRankingSortedData = new PreferenceRankingSortedDataModel();

                                var counterVal = counterList[j];

                                weightedSum = 0;
                                var apiDataEl = preferenceRankingData[j];

                                preferenceRankingSortedData.strTestName = apiDataEl.strTestName;

                                if (apiDataEl.intRankCount1 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount1 * counterVal;
                                    preferenceRankingSortedData.intRankCount1 = apiDataEl.intRankCount1;

                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount2 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount2 * counterVal;
                                    preferenceRankingSortedData.intRankCount2 = apiDataEl.intRankCount2;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount3 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount3 * counterVal;
                                    preferenceRankingSortedData.intRankCount3 = apiDataEl.intRankCount3;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount4 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount4 * counterVal;
                                    preferenceRankingSortedData.intRankCount4 = apiDataEl.intRankCount4;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount5 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount5 * counterVal;
                                    preferenceRankingSortedData.intRankCount5 = apiDataEl.intRankCount5;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount6 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount6 * counterVal;
                                    preferenceRankingSortedData.intRankCount6 = apiDataEl.intRankCount6;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount7 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount7 * counterVal;
                                    preferenceRankingSortedData.intRankCount7 = apiDataEl.intRankCount7;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount8 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount8 * counterVal;
                                    preferenceRankingSortedData.intRankCount8 = apiDataEl.intRankCount8;
                                    counterVal--;
                                }
                                if (apiDataEl.intRankCount9 != 0)
                                {
                                    weightedSum += apiDataEl.intRankCount9 * counterVal;
                                    preferenceRankingSortedData.intRankCount9 = apiDataEl.intRankCount9;
                                    counterVal--;
                                }

                                preferenceRankingSortedData.weightedScore = weightedSum;
                                lstWeightedData.Add(weightedSum);
                                apiDataWithWeightedList.Add(preferenceRankingSortedData);
                            }

                            var max = apiDataWithWeightedList.Max(m => m.weightedScore);

                            if (max > 100 && max < 200)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking200Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 200 && max < 300)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking300Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 300 && max < 400)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking400Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 400 && max < 500)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking500Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 500 && max < 600)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking600Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 600 && max < 700)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking700Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 700 && max < 800)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking800Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 800 && max < 900)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking900Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                            if (max > 900 && max < 1000)
                            {
                                sourcePath = $"C:\\ExcelChartFiles\\Templates\\PreferenceRanking1000Scale" + counterList.Count + "ColPlaceholder" + apiDataWithWeightedList.Count + ".pptx";
                                dLLClass.PreferenceRankingMethod(CreateTargetPath(sourcePath), apiDataWithWeightedList);
                            }
                        }

                        break;




                    case "Chemical Structure Appropriateness":
                    case "Chemical Structure":
                        List<ChemicalModel> chemicalData = new List<ChemicalModel>();
                        chemicalData = JsonConvert.DeserializeObject<List<ChemicalModel>>(data);

                        if (chemicalData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ChemicalStructure" + chemicalData.Count + ".pptx";
                            dLLClass.ChemicalStructureMethod(CreateTargetPath(sourcePath), chemicalData);
                        }
                        break;









                    case "Fit to Therapeutic Class":

                        List<FitToTherapeuticClassModel> therapeuticClassData = new List<FitToTherapeuticClassModel>();
                        therapeuticClassData = JsonConvert.DeserializeObject<List<FitToTherapeuticClassModel>>(data);

                        if (therapeuticClassData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticClass" + therapeuticClassData.Count + ".pptx";
                            dLLClass.FitToTherapeuticClassMethod(CreateTargetPath(sourcePath), therapeuticClassData, this.breakdown);
                        }
                        break;



                    case "Associations":
                        List<AssociationsModel> associationsData = new List<AssociationsModel>();
                        associationsData = JsonConvert.DeserializeObject<List<AssociationsModel>>(data, jsonSettings);

                        if (associationsData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Associations.pptx";
                            dLLClass.AssociationsMethod(CreateTargetPath(sourcePath), associationsData, breakdown);
                        }
                        break;




                    case "QTC":
                        List<QTCModel> qtcData = new List<QTCModel>();
                        qtcData = JsonConvert.DeserializeObject<List<QTCModel>>(data, jsonSettings);

                        if (qtcData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTC.pptx";
                            dLLClass.QTCMethod(CreateTargetPath(sourcePath), qtcData, breakdown);
                        }
                        break;


                    case "QTCCustom":
                        List<QTCCustomModel> qtcCustomData = new List<QTCCustomModel>();
                        qtcCustomData = JsonConvert.DeserializeObject<List<QTCCustomModel>>(data);

                        int counter = 0;

                        for (int i = 0; i < qtcCustomData.Count; i++)
                        {
                            var dataEl = qtcCustomData[i];
                            if (dataEl.strPageType1 != null)
                            {
                                counter++;
                            }
                            if (dataEl.strPageType2 != null)
                            {
                                counter++;
                            }
                            if (dataEl.strPageType3 != null)
                            {
                                counter++;
                            }
                            if (dataEl.strPageType4 != null)
                            {
                                counter++;
                            }
                            if (dataEl.strPageType5 != null)
                            {
                                counter++;
                            }
                            if (dataEl.strPageType6 != "")
                            {
                                counter++;
                            }
                            if (dataEl.strPageType7 != "")
                            {
                                counter++;
                            }
                            if (dataEl.strPageType8 != "")
                            {
                                counter++;
                            }
                            if (dataEl.strPageType9 != "")
                            {
                                counter++;
                            }
                            if (dataEl.strPageName10 != "")
                            {
                                counter++;
                            }
                            break;
                        }

                        if (qtcCustomData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\QTCCustom" + (counter * 2) + ".pptx";
                            dLLClass.QTCCustomMethod(CreateTargetPath(sourcePath), qtcCustomData, this.breakdown);
                        }
                        break;



                    case "Reflective of Mechanism of Action":
                        List<ReflectiveOfMechanismOfActionModel> reflectiveOfMechanismsData = new List<ReflectiveOfMechanismOfActionModel>();
                        reflectiveOfMechanismsData = JsonConvert.DeserializeObject<List<ReflectiveOfMechanismOfActionModel>>(data, jsonSettings);

                        if (reflectiveOfMechanismsData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ReflectiveofMechanismofAction" + reflectiveOfMechanismsData.Count + ".pptx";
                            dLLClass.ReflectiveofMechanismofActionMethod(CreateTargetPath(sourcePath), reflectiveOfMechanismsData);
                        }
                        break;


                    case "PhoneticTesting":
                        List<PhoneticTestingModel> phoneticData = new List<PhoneticTestingModel>();
                        phoneticData = JsonConvert.DeserializeObject<List<PhoneticTestingModel>>(data, jsonSettings);

                        if (phoneticData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\PhoneticTesting" + phoneticData.Count + ".pptx";
                            dLLClass.PhoneticTestingMethod(CreateTargetPath(sourcePath), phoneticData);
                        }
                        break;



                    case "Modifier Confusion":
                        List<OverallImpressionNewModel> modifierConfusionData = new List<OverallImpressionNewModel>();
                        modifierConfusionData = JsonConvert.DeserializeObject<List<OverallImpressionNewModel>>(data, jsonSettings);

                        if (modifierConfusionData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\ModifierConfusion" + modifierConfusionData.Count + ".pptx";
                            dLLClass.ModifierConfusionMethod(CreateTargetPath(sourcePath), modifierConfusionData, group);
                        }
                        break;


                    case "Medical Terms":
                    case "Medical Terms Similarity":

                        List<MedicalTermsModel> medicalData = new List<MedicalTermsModel>();
                        medicalData = JsonConvert.DeserializeObject<List<MedicalTermsModel>>(data);

                        if (medicalData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);

                        }
                        else
                        {
                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholder.pptx";
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\MedicalTermsPlaceholderRow.pptx";
                            dLLClass.MedicalTermsMethod(CreateTargetPath(sourcePath), medicalData);
                        }
                        break;


                    case "Non-Medical Terms":
                    case "Non - Medical Terms Similarity":
                    case "Non-Medical Terms Similarity":

                        List<NonMedicalTermsModel> nonMedicalData = new List<NonMedicalTermsModel>();
                        nonMedicalData = JsonConvert.DeserializeObject<List<NonMedicalTermsModel>>(data);

                        if (nonMedicalData.Count == 0)
                        {
                            //sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholderRow.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NonMedicalTermsPlaceholder.pptx";
                            dLLClass.NonMedicalTermsMethod(CreateTargetPath(sourcePath), nonMedicalData);
                        }
                        break;





                    case "BRANDEX LOGO":
                        List<BrandexLogoModel> brandexLogoData = new List<BrandexLogoModel>();
                        brandexLogoData = JsonConvert.DeserializeObject<List<BrandexLogoModel>>(data);

                        if (brandexLogoData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            List<BrandexLogoShortModel> brandexLogoShortData = new List<BrandexLogoShortModel>();

                            double average1MaxLogo = 0.0;
                            double average2MaxLogo = 0.0;
                            double average3MaxLogo = 0.0;
                            double average4MaxLogo = 0.0;
                            double average5MaxLogo = 0.0;

                            for (int i = 0; i < brandexLogoData.Count; i++)
                            {
                                var dblAverage1max = brandexLogoData[i].dblAveragePage1;
                                if (dblAverage1max == 0.0)
                                {
                                    average1MaxLogo = 0;
                                }
                                if (average1MaxLogo < dblAverage1max)
                                {
                                    average1MaxLogo = dblAverage1max;
                                }

                                var dblAverage2max = brandexLogoData[i].dblAveragePage2;
                                if (dblAverage2max == 0.0)
                                {
                                    average2MaxLogo = 0;
                                }
                                if (average2MaxLogo < dblAverage2max)
                                {
                                    average2MaxLogo = dblAverage2max;
                                }

                                var dblAverage3max = brandexLogoData[i].dblAveragePage3;
                                if (dblAverage3max == 0.0)
                                {
                                    average3MaxLogo = 0;
                                }
                                if (average3MaxLogo < dblAverage3max)
                                {
                                    average3MaxLogo = dblAverage3max;
                                }

                                var dblAverage4max = brandexLogoData[i].dblAveragePage4;
                                if (dblAverage4max == 0.0)
                                {
                                    average4MaxLogo = 0;
                                }
                                if (average4MaxLogo < dblAverage4max)
                                {
                                    average4MaxLogo = dblAverage4max;
                                }

                                var dblAverage5max = brandexLogoData[i].dblAveragePage5;
                                if (dblAverage5max == 0.0)
                                {
                                    average5MaxLogo = 0;
                                }
                                if (average5MaxLogo < dblAverage5max)
                                {
                                    average5MaxLogo = dblAverage5max;
                                }
                            }

                            double brandexLogoscalingFactor = 1.105380082;

                            for (int i = 0; i < brandexLogoData.Count; i++)
                            {
                                //for the table
                                BrandexLogoShortModel brandexLogoShortModelData = new BrandexLogoShortModel();

                                var dataEl = brandexLogoData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;
                                double averagePage3WeightedValue = 0.0;
                                double averagePage4WeightedValue = 0.0;
                                double averagePage5WeightedValue = 0.0;

                                if (average1MaxLogo > 0)
                                {
                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight;
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (average2MaxLogo > 0)
                                {
                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight;
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (average3MaxLogo > 0)
                                {
                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight;
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (average4MaxLogo > 0)
                                {
                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight;
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                if (average5MaxLogo > 0)
                                {
                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight;
                                }
                                else
                                {
                                    averagePage5WeightedValue = 0;
                                }

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexLogoscalingFactor;

                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

                                brandexLogoShortModelData.dblAveragePage1 = dataEl.dblAveragePage1;
                                brandexLogoShortModelData.dblPage1Weight = dataEl.dblPage1Weight;
                                brandexLogoShortModelData.dblAveragePage1Weighted = averagePage1WeightedValue;

                                brandexLogoShortModelData.dblAveragePage2 = dataEl.dblAveragePage2;
                                brandexLogoShortModelData.dblPage2Weight = dataEl.dblPage2Weight;
                                brandexLogoShortModelData.dblAveragePage2Weighted = averagePage2WeightedValue;

                                brandexLogoShortModelData.dblAveragePage3 = dataEl.dblAveragePage3;
                                brandexLogoShortModelData.dblPage3Weight = dataEl.dblPage3Weight;
                                brandexLogoShortModelData.dblAveragePage3Weighted = averagePage3WeightedValue;

                                brandexLogoShortModelData.dblAveragePage4 = dataEl.dblAveragePage4;
                                brandexLogoShortModelData.dblPage4Weight = dataEl.dblPage4Weight;
                                brandexLogoShortModelData.dblAveragePage4Weighted = averagePage4WeightedValue;

                                brandexLogoShortModelData.dblAveragePage5 = dataEl.dblAveragePage5;
                                brandexLogoShortModelData.dblPage5Weight = dataEl.dblPage5Weight;
                                brandexLogoShortModelData.dblAveragePage5Weighted = averagePage5WeightedValue;

                                brandexLogoShortModelData.dblAveragePage6 = dataEl.dblAveragePage6;
                                brandexLogoShortModelData.dblPage6Weight = dataEl.dblPage6Weight;
                                brandexLogoShortModelData.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

                                brandexLogoShortModelData.dblAveragePage7 = dataEl.dblAveragePage7;
                                brandexLogoShortModelData.dblPage7Weight = dataEl.dblPage7Weight;
                                brandexLogoShortModelData.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

                                brandexLogoShortModelData.dblAveragePage8 = dataEl.dblAveragePage8;
                                brandexLogoShortModelData.dblPage8Weight = dataEl.dblPage8Weight;
                                brandexLogoShortModelData.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

                                brandexLogoShortModelData.dblIndex = indexSum;
                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
                                brandexLogoShortModelData.intRed = dataEl.intRed;
                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
                                brandexLogoShortModelData.intBlue = dataEl.intBlue;


                                //for the chart - 
                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;
                                double averagePage3WeightedValueForChart = 0.0;
                                double averagePage4WeightedValueForChart = 0.0;
                                double averagePage5WeightedValueForChart = 0.0;

                                if (average1MaxLogo > 0)
                                {
                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / average1MaxLogo) * dataEl.dblPage1Weight) * brandexLogoscalingFactor;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (average2MaxLogo > 0)
                                {
                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / average2MaxLogo) * dataEl.dblPage2Weight) * brandexLogoscalingFactor;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (average3MaxLogo > 0)
                                {
                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / average3MaxLogo) * dataEl.dblPage3Weight) * brandexLogoscalingFactor;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (average4MaxLogo > 0)
                                {
                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / average4MaxLogo) * dataEl.dblPage4Weight) * brandexLogoscalingFactor;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                if (average5MaxLogo > 0)
                                {
                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / average5MaxLogo) * dataEl.dblPage5Weight) * brandexLogoscalingFactor;
                                }
                                else
                                {
                                    averagePage5WeightedValueForChart = 0;
                                }

                                double indexSumForChart = averagePage1WeightedValueForChart +
                                                          averagePage2WeightedValueForChart +
                                                          averagePage3WeightedValueForChart +
                                                          averagePage4WeightedValueForChart +
                                                          averagePage5WeightedValueForChart;

                                brandexLogoShortModelData.strTestName = dataEl.strTestName;

                                brandexLogoShortModelData.dblAveragePage1WeightedForChart = Math.Round(averagePage1WeightedValueForChart, 1);

                                brandexLogoShortModelData.dblAveragePage2WeightedForChart = Math.Round(averagePage2WeightedValueForChart, 1);

                                brandexLogoShortModelData.dblAveragePage3WeightedForChart = Math.Round(averagePage3WeightedValueForChart, 1);

                                brandexLogoShortModelData.dblAveragePage4WeightedForChart = Math.Round(averagePage4WeightedValueForChart, 1); ;

                                brandexLogoShortModelData.dblAveragePage5WeightedForChart = Math.Round(averagePage5WeightedValueForChart, 1); ;

                                brandexLogoShortModelData.dblIndexForChart = indexSumForChart;
                                brandexLogoShortModelData.strDSIScore = dataEl.strDSIScore;
                                brandexLogoShortModelData.intRed = dataEl.intRed;
                                brandexLogoShortModelData.intGreen = dataEl.intGreen;
                                brandexLogoShortModelData.intBlue = dataEl.intBlue;

                                brandexLogoShortData.Add(brandexLogoShortModelData);
                            }

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexLogoTemplate" + brandexLogoShortData.Count + ".pptx";

                            dLLClass.BrandexLogoMethod(CreateTargetPath(sourcePath), brandexLogoShortData, this.breakdown);
                        }
                        break;








                    case "SALA":
                    case "Sound Alike-Look Alike":
                        List<SALANewModel> salaData = new List<SALANewModel>();
                        salaData = JsonConvert.DeserializeObject<List<SALANewModel>>(data, jsonSettings);

                        if (salaData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            CreateTargetPath(sourcePath);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";
                            dLLClass.SALANewMethod(CreateTargetPath(sourcePath), salaData);
                        }
                        break;





                    case "Negative Communication":
                    case "Negative or Offensive Communication":


                        List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
                        negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data);

                        if (negCommData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + negCommData.Count + ".pptx";
                            dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), negCommData);
                        }
                        break;


                    case "Negative Communication Rationale":


                        List<NegativeCommunicationRationaleModel> negCommRationaleData = new List<NegativeCommunicationRationaleModel>();
                        negCommRationaleData = JsonConvert.DeserializeObject<List<NegativeCommunicationRationaleModel>>(data);

                        if (negCommRationaleData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunicationRationalePlaceholder.pptx";
                            dLLClass.NegativeCommRationaleMethod(CreateTargetPath(sourcePath), negCommRationaleData, this.breakdown);
                        }
                        break;


                    case "Fit to Theraputic Category":
                    case "Fit to Therapeutic Category":
                        List<FitToTheraputicCategoryModel> theraputicData = new List<FitToTheraputicCategoryModel>();
                        theraputicData = JsonConvert.DeserializeObject<List<FitToTheraputicCategoryModel>>(data);

                        List<FitToTheraputicCategoryModel> testnameCountList = new List<FitToTheraputicCategoryModel>();

                        for (int i = 0; i < theraputicData.Count; i++)
                        {
                            var groupData = theraputicData.GroupBy(x => x.strTestName).ToList();
                            if (groupData != null)
                            {
                                var firstTestnameList = groupData[0].ToList();
                                testnameCountList = firstTestnameList;
                            }
                            break;
                        }

                        int tableHeaderCounter = testnameCountList.Count;

                        if (theraputicData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\FitToTherapeuticCategory" + tableHeaderCounter + ".pptx";
                            dLLClass.FitToTheraputicMethod(CreateTargetPath(sourcePath), theraputicData);
                        }
                        break;






                    //    List<NegativeCommunicationModel> negCommData = new List<NegativeCommunicationModel>();
                    //    negCommData = JsonConvert.DeserializeObject<List<NegativeCommunicationModel>>(data, jsonSettings);


                    //    if (negCommData.Count == 0)
                    //    {
                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                    //        CreateTargetPath(sourcePath);
                    //    }
                    //    else
                    //    {
                    //        if (this.template == "Sound Alike-Look Alike")
                    //        {
                    //            this.template = "SALA";
                    //        }

                    //        //sourcePath = $"C:\\ExcelChartFiles\\Templates\\SALAPlaceholder.pptx";


                    //        int dataIndex = 0;
                    //        double sum = 0;
                    //        List<NegativeCommunicationModelShort> lstNegCommValue = new List<NegativeCommunicationModelShort>();

                    //        var groupApiData = negCommData.GroupBy(item => item.strTestName).OrderBy(group => group.Key).ToList();


                    //        while (dataIndex < groupApiData.Count)
                    //        {
                    //            sum = 0;
                    //            var testNameGroup = groupApiData[dataIndex];
                    //            var testNameData = testNameGroup.ToList();

                    //            foreach (var Tdata in testNameData)
                    //            {
                    //                sum += Tdata.intSum;
                    //            }

                    //            double total = testNameData.FirstOrDefault().intTotal;

                    //            double percentageVal = (double)(((total - sum) / total) * 100);

                    //            double remainingPercentageVal = 100 - percentageVal;

                    //            // lstNegCommValue.FirstOrDefault().percentage = percentageVal.ToString();

                    //            NegativeCommunicationModelShort ncomShort = new NegativeCommunicationModelShort();

                    //            ncomShort.percentage = percentageVal.ToString();
                    //            ncomShort.strTestName = testNameData.FirstOrDefault().strTestName;
                    //            ncomShort.intBlue = testNameData.Max(x => x.intBlue);
                    //            ncomShort.intGreen = testNameData.Max(x => x.intGreen);
                    //            ncomShort.intRed = testNameData.Max(x => x.intRed);
                    //            ncomShort.remainingPercentage = remainingPercentageVal.ToString();

                    //            lstNegCommValue.Add(ncomShort);

                    //            dataIndex++;
                    //        }


                    //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NegativeCommunication" + lstNegCommValue.Count + ".pptx";
                    //        dLLClass.NegativeCommMethod(CreateTargetPath(sourcePath), lstNegCommValue);




                    //    }
                    //    break;


                    case "Distinctiveness":
                        List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
                        distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

                        if (distinctData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            notAvailable(sourcePath, template, breakdown);
                        }
                        else
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
                            dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
                        }
                        break;



                    case "BRANDEX MEDICAL DEVICE 1":
                    case "Brandex Medical Device 1":

                        List<BrandexMedicalDevice1Model> brandexMedicalData = new List<BrandexMedicalDevice1Model>();
                        brandexMedicalData = JsonConvert.DeserializeObject<List<BrandexMedicalDevice1Model>>(data);

                        if (brandexMedicalData.Count == 0)
                        {
                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                            dLLClass.NotAvailableMethod(CreateTargetPath(sourcePath), this.template, this.breakdown);
                        }
                        else
                        {
                            List<BrandexMedicalDevice1ShortModel> brandexMedicalShortData = new List<BrandexMedicalDevice1ShortModel>();

                            double brandexMedicalAverage1Max = 0.0;
                            double brandexMedicalAverage2Max = 0.0;
                            double brandexMedicalAverage3Max = 0.0;
                            double brandexMedicalAverage4Max = 0.0;
                            double brandexMedicalAverage5Max = 0.0;

                            for (int i = 0; i < brandexMedicalData.Count; i++)
                            {
                                var dblAverage1max = brandexMedicalData[i].dblAveragePage1;
                                if (dblAverage1max == 0.0)
                                {
                                    brandexMedicalAverage1Max = 0;
                                }
                                if (brandexMedicalAverage1Max < dblAverage1max)
                                {
                                    brandexMedicalAverage1Max = dblAverage1max;
                                }

                                var dblAverage2max = brandexMedicalData[i].dblAveragePage2;
                                if (dblAverage2max == 0.0)
                                {
                                    brandexMedicalAverage2Max = 0;
                                }
                                if (brandexMedicalAverage2Max < dblAverage2max)
                                {
                                    brandexMedicalAverage2Max = dblAverage2max;
                                }

                                var dblAverage3max = brandexMedicalData[i].dblAveragePage3;
                                if (dblAverage3max == 0.0)
                                {
                                    brandexMedicalAverage3Max = 0;
                                }
                                if (brandexMedicalAverage3Max < dblAverage3max)
                                {
                                    brandexMedicalAverage3Max = dblAverage3max;
                                }

                                var dblAverage4max = brandexMedicalData[i].dblAveragePage4;
                                if (dblAverage4max == 0.0)
                                {
                                    brandexMedicalAverage4Max = 0;
                                }
                                if (brandexMedicalAverage4Max < dblAverage4max)
                                {
                                    brandexMedicalAverage4Max = dblAverage4max;
                                }

                                var dblAverage5max = brandexMedicalData[i].dblAveragePage5;
                                if (dblAverage5max == 0.0)
                                {
                                    brandexMedicalAverage5Max = 0;
                                }
                                if (brandexMedicalAverage5Max < dblAverage5max)
                                {
                                    brandexMedicalAverage5Max = dblAverage5max;
                                }
                            }

                            double brandexMedicalscalingFactor = 1.0023;

                            for (int i = 0; i < brandexMedicalData.Count; i++)
                            {
                                //for the table
                                BrandexMedicalDevice1ShortModel brandexSafetyShortModel = new BrandexMedicalDevice1ShortModel();

                                var dataEl = brandexMedicalData[i];

                                double averagePage1WeightedValue = 0.0;
                                double averagePage2WeightedValue = 0.0;
                                double averagePage3WeightedValue = 0.0;
                                double averagePage4WeightedValue = 0.0;
                                double averagePage5WeightedValue = 0.0;

                                if (brandexMedicalAverage1Max > 0)
                                {
                                    averagePage1WeightedValue = (dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight;
                                }
                                else
                                {
                                    averagePage1WeightedValue = 0;
                                }

                                if (brandexMedicalAverage2Max > 0)
                                {
                                    averagePage2WeightedValue = (dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight;
                                }
                                else
                                {
                                    averagePage2WeightedValue = 0;
                                }

                                if (brandexMedicalAverage3Max > 0)
                                {
                                    averagePage3WeightedValue = (dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight;
                                }
                                else
                                {
                                    averagePage3WeightedValue = 0;
                                }

                                if (brandexMedicalAverage4Max > 0)
                                {
                                    averagePage4WeightedValue = (dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight;
                                }
                                else
                                {
                                    averagePage4WeightedValue = 0;
                                }

                                if (brandexMedicalAverage5Max > 0)
                                {
                                    averagePage5WeightedValue = (dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight;
                                }
                                else
                                {
                                    averagePage5WeightedValue = 0;
                                }

                                double indexSum = (averagePage1WeightedValue + averagePage2WeightedValue + averagePage3WeightedValue +
                                                  averagePage4WeightedValue + averagePage5WeightedValue) * brandexMedicalscalingFactor;

                                brandexSafetyShortModel.strTestName = dataEl.strTestName;

                                brandexSafetyShortModel.dblAveragePage1 = dataEl.dblAveragePage1;
                                brandexSafetyShortModel.dblPage1Weight = dataEl.dblPage1Weight;
                                brandexSafetyShortModel.dblAveragePage1Weighted = averagePage1WeightedValue;

                                brandexSafetyShortModel.dblAveragePage2 = dataEl.dblAveragePage2;
                                brandexSafetyShortModel.dblPage2Weight = dataEl.dblPage2Weight;
                                brandexSafetyShortModel.dblAveragePage2Weighted = averagePage2WeightedValue;

                                brandexSafetyShortModel.dblAveragePage3 = dataEl.dblAveragePage3;
                                brandexSafetyShortModel.dblPage3Weight = dataEl.dblPage3Weight;
                                brandexSafetyShortModel.dblAveragePage3Weighted = averagePage3WeightedValue;

                                brandexSafetyShortModel.dblAveragePage4 = dataEl.dblAveragePage4;
                                brandexSafetyShortModel.dblPage4Weight = dataEl.dblPage4Weight;
                                brandexSafetyShortModel.dblAveragePage4Weighted = averagePage4WeightedValue;

                                brandexSafetyShortModel.dblAveragePage5 = dataEl.dblAveragePage5;
                                brandexSafetyShortModel.dblPage5Weight = dataEl.dblPage5Weight;
                                brandexSafetyShortModel.dblAveragePage5Weighted = averagePage5WeightedValue;

                                brandexSafetyShortModel.dblAveragePage6 = dataEl.dblAveragePage6;
                                brandexSafetyShortModel.dblPage6Weight = dataEl.dblPage6Weight;
                                brandexSafetyShortModel.dblAveragePage6Weighted = dataEl.dblAveragePage6Weighted;

                                brandexSafetyShortModel.dblAveragePage7 = dataEl.dblAveragePage7;
                                brandexSafetyShortModel.dblPage7Weight = dataEl.dblPage7Weight;
                                brandexSafetyShortModel.dblAveragePage7Weighted = dataEl.dblAveragePage7Weighted;

                                brandexSafetyShortModel.dblAveragePage8 = dataEl.dblAveragePage8;
                                brandexSafetyShortModel.dblPage8Weight = dataEl.dblPage8Weight;
                                brandexSafetyShortModel.dblAveragePage8Weighted = dataEl.dblAveragePage8Weighted;

                                brandexSafetyShortModel.dblIndex = indexSum;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = dataEl.intRed;
                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


                                //for the chart - 

                                double averagePage1WeightedValueForChart = 0.0;
                                double averagePage2WeightedValueForChart = 0.0;
                                double averagePage3WeightedValueForChart = 0.0;
                                double averagePage4WeightedValueForChart = 0.0;
                                double averagePage5WeightedValueForChart = 0.0;

                                if (brandexMedicalAverage1Max > 0)
                                {
                                    averagePage1WeightedValueForChart = ((dataEl.dblAveragePage1 / brandexMedicalAverage1Max) * dataEl.dblPage1Weight) * brandexMedicalscalingFactor;
                                }
                                else
                                {
                                    averagePage1WeightedValueForChart = 0;
                                }

                                if (brandexMedicalAverage2Max > 0)
                                {
                                    averagePage2WeightedValueForChart = ((dataEl.dblAveragePage2 / brandexMedicalAverage2Max) * dataEl.dblPage2Weight) * brandexMedicalscalingFactor;
                                }
                                else
                                {
                                    averagePage2WeightedValueForChart = 0;
                                }

                                if (brandexMedicalAverage3Max > 0)
                                {
                                    averagePage3WeightedValueForChart = ((dataEl.dblAveragePage3 / brandexMedicalAverage3Max) * dataEl.dblPage3Weight) * brandexMedicalscalingFactor;
                                }
                                else
                                {
                                    averagePage3WeightedValueForChart = 0;
                                }

                                if (brandexMedicalAverage4Max > 0)
                                {
                                    averagePage4WeightedValueForChart = ((dataEl.dblAveragePage4 / brandexMedicalAverage4Max) * dataEl.dblPage4Weight) * brandexMedicalscalingFactor;
                                }
                                else
                                {
                                    averagePage4WeightedValueForChart = 0;
                                }

                                if (brandexMedicalAverage5Max > 0)
                                {
                                    averagePage5WeightedValueForChart = ((dataEl.dblAveragePage5 / brandexMedicalAverage5Max) * dataEl.dblPage5Weight) * brandexMedicalscalingFactor;
                                }
                                else
                                {
                                    averagePage5WeightedValueForChart = 0;
                                }

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
                                ;
                                brandexSafetyShortModel.dblIndexForChart = indexSumForChart;
                                brandexSafetyShortModel.strDSIScore = dataEl.strDSIScore;
                                brandexSafetyShortModel.intRed = dataEl.intRed;
                                brandexSafetyShortModel.intGreen = dataEl.intGreen;
                                brandexSafetyShortModel.intBlue = dataEl.intBlue;
                                brandexSafetyShortModel.boolBold = dataEl.boolBold;


                                brandexMedicalShortData.Add(brandexSafetyShortModel);
                            }

                            sourcePath = $"C:\\ExcelChartFiles\\Templates\\BrandexMedicalDevice1Template" + brandexMedicalShortData.Count + ".pptx";
                            dLLClass.BrandexMedicalDevice1Method(CreateTargetPath(sourcePath), brandexMedicalShortData);
                        }
                        break;











                    default:
                        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                        notAvailable(CreateTargetPath(sourcePath), template, breakdown);
                        break;


                        //case "":
                        //    List<DistinctivenessModel> distinctData = new List<DistinctivenessModel>();
                        //    distinctData = JsonConvert.DeserializeObject<List<DistinctivenessModel>>(data, jsonSettings);

                        //    if (distinctData.Count == 0)
                        //    {
                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                        //        notAvailable(sourcePath, template, breakdown);
                        //    }
                        //    else
                        //    {
                        //        sourcePath = $"C:\\ExcelChartFiles\\Templates\\Distinctiveness" + distinctData.Count + ".pptx";
                        //        dLLClass.DistinctivenessMethod(CreateTargetPath(sourcePath), distinctData);
                        //    }
                        //    break;


                        //default:
                        //    sourcePath = $"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx";
                        //    notAvailable(CreateTargetPath(sourcePath), template, breakdown);
                        //    break;



                }
            }




        }


        public async Task<string> addChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string BreakDown)
        {
            //hello
            await Task.Run(() => fnaddChartsToFinalTemplate1(project, charts, finalTemplate, BreakDown));
            return "Process sucessful";
        }


        public string getAttributeTitle(string project, string chart)
        {
            string repText = "";
            System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

            foreach (DataRow row in dt1.Rows)
            {
                repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
            }

            return repText;


        }


        public async Task fnaddChartsToFinalTemplate1(string project, List<string> charts, string finalTemplate, string breakDown)
        {

            if (!Directory.Exists($"C:\\excelfiles\\{project}\\Final"))
            {
                Directory.CreateDirectory($"C:\\excelfiles\\{project}\\Final");
            }

            //temp copy the final template file to the path 

            File.Copy("\\\\miafs02\\Market Research\\MR Programs\\ExcelCharts_Chartsdll\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", true);


            //get all the pagegroup names for the chart 

            List<clschartPageGroup> lstPg = getProjectPagegroupNames(project);


            //get the display names of the charts in excel sheet app

            List<clschartPageDisplayName> lstChartDispNames = getProjectChartDisplayNames(project);

            string op = "";

            //get the page settings for the template 

            System.Data.DataTable dt = clsData.MRData.getDataTable("ExcelChartsPrc_getPPTFinalSettings " + "'" + finalTemplate + "'," + "'" + "BI - 2024" + "'");
            List<clsPPTFinalSettings> lstPPTFinalSettings = new List<clsPPTFinalSettings>();

            int DelLastPage = 0;
            int DelFirstPage = 0;
            List<int> intChartPages = new List<int>();


            foreach (DataRow row in dt.Rows)
            {
                // Create a new object for each row and add it to the list
                clsPPTFinalSettings data1 = new clsPPTFinalSettings
                {
                    intPPTSlideIndexFirst = Convert.ToInt32(row["intPPTSlideIndexFirst"]),
                    intPPTSlideIndexLast = Convert.ToInt32(row["intPPTSlideIndexLast"]),
                    strTemplateName = Convert.ToString(row["strTemplateName"]),
                    strTemplateSourcePath = Convert.ToString(row["strTemplateSourcePath"]),
                    strPageGroupName = Convert.ToString(row["strPageGroupName"]),
                    strPageGroupType = Convert.ToString(row["strPageGroupType"]),
                };
                lstPPTFinalSettings.Add(data1);
            }



            int specialChartTypeCount = 0;

            //copy a final in the path

            createFolder($"C:\\excelfiles\\{project}\\Final");
            copyFile("C:\\ExcelChartsTemplatesNew\\ExcelCharts_ChartTemplates\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx", $"C:\\excelfiles\\{project}\\Final\\" + finalTemplate.Replace(" ", "").Trim() + ".pptx");


            lstPPTFinalSettings = lstPPTFinalSettings.OrderByDescending(p => p.intPPTSlideIndexFirst).ToList();
            int chartsCompCount = 0;
            List<string> chartsCompleted = new List<string>();
            List<string> pageGroupNameCompleted = new List<string>();
            List<clsChartDislayNamePageGroupName> lstchartDispPageGroupName = new List<clsChartDislayNamePageGroupName>();


            //
            foreach (clschartPageDisplayName obj in lstChartDispNames)
            {
                foreach (string chart in charts)
                {
                    if (obj.strPageName == chart)
                    {
                        obj.isReportSelectedByUser = true;


                        //updating Attribute evaluation cover page :


                        if (obj.strPageType.ToLower() == "attribute evaluation")
                        {
                            if (obj.strPageName.Contains("1"))
                            {
                                //update the slide 38

                                string repText = getAttributeTitle(project, chart);


                                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);



                                //await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);


                            }

                            if (obj.strPageName.Contains("2"))
                            {
                                //update the slide 38

                                string repText = getAttributeTitle(project, chart);

                                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", repText, 38);

                                // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

                                //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle2", repText, 38);


                            }

                            if (obj.strPageName.Contains("3"))
                            {
                                //update the slide 38

                                string repText = getAttributeTitle(project, chart);


                                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", repText, 38);

                                //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

                                //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle3", repText, 38);


                            }

                            if (obj.strPageName.Contains("4"))
                            {
                                //update the slide 38

                                string repText = getAttributeTitle(project, chart);

                                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", repText, 38);

                                //  await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

                                // await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle4", repText, 38);


                            }


                        }



                    }

                }

            }


            //in case if the attribute evaluations are not there 

            List<clschartPageDisplayName> lstAtts = lstChartDispNames.Where(p => p.strPageType.ToLower() == "attribute evaluation").ToList();


            if (!lstAtts.Any(obj => obj.strPageName.Contains("1")) || lstAtts.Any(obj => obj.strPageName.Contains("1") && obj.isReportSelectedByUser == false))
            {

                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);

            }

            if (!lstAtts.Any(obj => obj.strPageName.Contains("2")) || lstAtts.Any(obj => obj.strPageName.Contains("2") && obj.isReportSelectedByUser == false))
            {

                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);

            }


            if (!lstAtts.Any(obj => obj.strPageName.Contains("3")) || lstAtts.Any(obj => obj.strPageName.Contains("3") && obj.isReportSelectedByUser == false))
            {

                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);

            }

            if (!lstAtts.Any(obj => obj.strPageName.Contains("4")) || lstAtts.Any(obj => obj.strPageName.Contains("4") && obj.isReportSelectedByUser == false))
            {

                await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);

            }






            //merge


            //bool first chart name replaced

            bool firstChartTextReplaced = false;
            foreach (clsPPTFinalSettings objclsPPT in lstPPTFinalSettings)
            {
                if (chartsCompCount == charts.Count())
                {
                    break;
                }

                specialChartTypeCount = 0;
                //check if chart type is selected  

                bool ifThePageTypeIsSelected = true;


                //first chart project name and details change 

                if (!firstChartTextReplaced)
                {
                    DateTime currentDate = DateTime.Now;

                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "ProjectName", project, 0);

                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<", "", 0);

                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", ">>", "", 0);

                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Month", currentDate.ToString("MMMM"), 0);

                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "Year", DateTime.Now.Year.ToString(), 0);


                    //replacing the attribute evaluation cover page :

                    firstChartTextReplaced = true;

                }



                //delete the slides that are not part of the report 


                if (!lstChartDispNames.Any(p => p.strPageType == objclsPPT.strPageGroupType))
                {

                    // skip the Attribute Evaluation Cover if the report has Attribute eavluation

                    if (pageGroupNameCompleted.Any(z => z.Contains("Attribute")) && objclsPPT.strPageGroupType == "Attribute Evaluation Cover")
                    {
                        continue;
                    }

                    else
                    {
                        DelLastPage = objclsPPT.intPPTSlideIndexLast;
                        DelFirstPage = objclsPPT.intPPTSlideIndexFirst;


                        if (DelLastPage - DelFirstPage == 0)
                        {
                            for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
                            {
                                if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
                                {
                                    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

                                }

                            }

                        }

                        else
                        {
                            for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
                            {
                                if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
                                {
                                    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

                                }

                            }

                        }


                        continue;

                    }





                }



                foreach (string chart in charts)
                {

                    //if the chart is completed break

                    if (pageGroupNameCompleted.Contains(objclsPPT.strPageGroupName))
                    {
                        break;
                    }





                    //get the chart group type

                    List<clschartPageDisplayName> lstCheck = lstChartDispNames.Where(p => p.strPageName == chart).ToList();

                    bool noproceed = false;

                    //delete the slides from the template if a report is not selected .

                    bool isReportSelected = false;


                    if (lstChartDispNames.Any(item => item.strPageType == objclsPPT.strPageGroupType))
                    {

                        if (objclsPPT.strPageGroupType.ToLower() == "exaggerative-inappropriate")
                        {

                            //if( lstChartDispNames.Any(item =>item.strPageType.ToLower() == "exaggerative-inappropriate" && getNumbersFromString(item.strPageName)==getNumbersFromString(chart)))
                            //{
                            //     isReportSelected = true;
                            //}



                            List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "exaggerative-inappropriate").ToList();

                            foreach (clschartPageDisplayName obj in lstFiltered)
                            {
                                if (getNumbersFromString(obj.strPageName) == getNumbersFromString(objclsPPT.strPageGroupName))
                                {

                                    if (obj.isReportSelectedByUser)
                                    {
                                        isReportSelected = true;
                                        break;

                                    }

                                }

                            }


                        }

                        else if (objclsPPT.strPageGroupType.ToLower() == "attribute evaluation")
                        {


                            List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "attribute evaluation" || item.strPageType.ToLower() == "attribute evaluation aggregate").ToList();

                            foreach (clschartPageDisplayName obj in lstFiltered)
                            {
                                if (getNumbersFromString(obj.strPageName) == getNumbersFromString(objclsPPT.strPageGroupName))
                                {


                                    if (obj.isReportSelectedByUser)
                                    {
                                        isReportSelected = true;
                                        break;

                                    }
                                }

                            }



                        }

                        else
                        {
                            List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == objclsPPT.strPageGroupName.ToLower()).ToList();

                            foreach (clschartPageDisplayName obj in lstFiltered)
                            {
                                if (obj.isReportSelectedByUser)
                                {
                                    isReportSelected = true;
                                    break;

                                }

                            }

                        }


                    }


                    //if (lstChartDispNames.Any(item => item.strPageType == objclsPPT.strPageGroupType))
                    //{
                    //    foreach (clschartPageDisplayName item in lstChartDispNames)
                    //    {
                    //        if (item.strPageType == objclsPPT.strPageGroupType && item.strPageName==chart)
                    //        {
                    //            if (!item.isReportSelectedByUser)
                    //            {
                    //                isReportSelected = false;
                    //                break;
                    //            }
                    //        }
                    //    }


                    //}

                    //else
                    //{
                    //    isReportSelected = false;


                    //}


                    if (!isReportSelected)
                    {
                        //delete if the project is not checked 
                        DelLastPage = objclsPPT.intPPTSlideIndexLast;
                        DelFirstPage = objclsPPT.intPPTSlideIndexFirst;

                        if (DelLastPage - DelFirstPage == 1 && objclsPPT.strPageGroupName.ToLower() == "attribute evaluation cover")
                        {
                            List<clschartPageDisplayName> lstFiltered = lstChartDispNames.Where(item => item.strPageType.ToLower() == "attribute evaluation" && item.isReportSelectedByUser == true).ToList();

                            if (lstFiltered.Count > 0)
                            {
                                break;
                            }

                            else
                            {
                                DelLastPage = DelFirstPage;
                            }


                        }

                        if (!pageGroupNameCompleted.Contains(objclsPPT.strPageGroupName) && (DelLastPage != 0 && DelFirstPage != 0))
                        {

                            if (DelLastPage - DelFirstPage == 0)
                            {
                                for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
                                {
                                    if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
                                    {
                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

                                    }

                                }

                            }

                            else
                            {
                                for (int k = 0; k < DelLastPage - DelFirstPage + 1; k++)
                                {
                                    if (DelLastPage - k - 1 < CountSlides($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx"))
                                    {
                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);

                                    }

                                }

                            }


                        }


                        break;


                    }



                    foreach (clschartPageDisplayName oj in lstCheck)
                    {

                        if (objclsPPT.strPageGroupType.ToLower() == "exaggerative-inappropriate" || objclsPPT.strPageGroupType.ToLower() == "attribute evaluation")
                        {

                            if (oj.strPageName.ToLower() == "attribute evaluation aggregate")
                            {
                                oj.strPageType = "Attribute Evaluation";
                            }

                            if ((oj.strPageType == objclsPPT.strPageGroupType) && (oj.strPageName == chart) && (getNumbersFromString(objclsPPT.strPageGroupName) == getNumbersFromString(chart)))
                            {
                                noproceed = false;
                                break;

                            }

                            else
                            {
                                noproceed = true;
                                break;

                            }

                        }

                        else
                        {
                            if ((oj.strPageType == objclsPPT.strPageGroupType) && (oj.strPageName == chart))
                            {
                                noproceed = false;
                                break;

                            }

                            else
                            {
                                noproceed = true;
                                break;

                            }


                        }





                    }

                    if (noproceed)
                    {
                        continue;
                    }




                    if (!chartsCompleted.Contains(chart))
                    {

                        if (objclsPPT.strPageGroupName != "Main Cover")
                        {
                            DelLastPage = objclsPPT.intPPTSlideIndexLast;
                            DelFirstPage = objclsPPT.intPPTSlideIndexFirst;


                            //add

                            string firstMatchingName = lstPg
                           .Where(p => p.strPageGroup == chart)
                           .Select(p => p.strPageGroupType)
                           .FirstOrDefault();


                            if (objclsPPT.strPageGroupType != "Attribute Evaluation" && objclsPPT.strPageGroupType != "Exaggerative-Inappropriate")
                            {
                                //item array of pages with slide numbers 
                                intChartPages = new List<int>();

                                if (DelLastPage > DelFirstPage)
                                {
                                    if (objclsPPT.strPageGroupType == "Phonetic Testing" || objclsPPT.strPageGroupType == "JSCAN")
                                    {
                                        intChartPages.Add(DelLastPage);

                                    }

                                    else if (objclsPPT.strPageGroupType == "Sound Alike-Look Alike" || objclsPPT.strPageGroupType == "Medical Terms Similarity" || objclsPPT.strPageGroupType == "Non-Medical Terms Similarity" || objclsPPT.strPageGroupType == "Brandex Strategic Distinctiveness" || objclsPPT.strPageGroupType == "Brandex Safety")
                                    {
                                        for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
                                        {
                                            intChartPages.Add(DelFirstPage + l);
                                        }

                                    }
                                    else
                                    {
                                        for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
                                        {
                                            intChartPages.Add(DelFirstPage + l - 1);
                                        }


                                    }



                                }
                                else
                                {
                                    intChartPages.Add(DelLastPage);
                                }



                                if (objclsPPT.strPageGroupName.ToLower() == chart.ToLower() || objclsPPT.strPageGroupName.ToLower() == firstMatchingName.ToLower())
                                {
                                    for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
                                    {
                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
                                        //Thread.Sleep(100);
                                    }

                                    if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
                                    {
                                        for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
                                        {
                                            intChartPages.RemoveAt(i);
                                        }

                                    }

                                    if (File.Exists(getIndividualChartPath(chart, project, breakDown)))
                                    {
                                        op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);


                                    }


                                    chartsCompCount = chartsCompCount + 1;

                                    lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT.strPageGroupName));

                                    chartsCompleted.Add(chart);

                                    pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);
                                }


                            }

                            else
                            {
                                //item array of pages with slide numbers 
                                intChartPages = new List<int>();

                                if (DelLastPage > DelFirstPage && DelLastPage - DelFirstPage > 1)
                                {
                                    for (int l = 0; l <= DelLastPage - DelFirstPage; l++)
                                    {
                                        intChartPages.Add(DelFirstPage + l - 1);
                                    }

                                }
                                else
                                {
                                    intChartPages.Add(DelLastPage);
                                }

                                //get the attribute evaluation category and Exaggerative 

                                string PageType = lstChartDispNames.FirstOrDefault(item => item.strPageName == chart)?.strPageType;

                                string resPageType = "";

                                op = "";

                                if (PageType == "Attribute Evaluation" && chart.Contains("1"))
                                {

                                    resPageType = "Attribute Evaluation 1";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");





                                }

                                else if (PageType == "Attribute Evaluation" && chart.Contains("2"))
                                {
                                    resPageType = "Attribute Evaluation 2";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 2", "FirstPage");
                                    // await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "Attribute #2", 38);

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


                                    //}



                                }

                                else if (PageType == "Attribute Evaluation" && chart.Contains("3"))
                                {

                                    resPageType = "Attribute Evaluation 3";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 3", "FirstPage");
                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "Attribute #3", 38);

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 4))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "", 38);


                                    //}



                                }

                                else if (PageType == "Attribute Evaluation" && chart.Contains("4"))
                                {
                                    resPageType = "Attribute Evaluation 4";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 4", "FirstPage");
                                    await clsMisc.repTextInSlideAsync($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle4>>", "Attribute #4", 38);

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 1))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "", "", 38);


                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 2))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle2>>", "", 38);


                                    //}

                                    //if (!checkParticularAttributeEvaluationIsSelected(lstChartDispNames, 3))
                                    //{
                                    //    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle3>>", "", 38);


                                    //}


                                }


                                else if (PageType.Contains("Attribute Evaluation") && chart.Contains("Aggregate") && chart.Contains("Attribute"))
                                {
                                    resPageType = "Attribute Evaluation Aggregate";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation Aggregate", "FirstPage");


                                }


                                else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "1" || getNumbersFromString(chart) == "01"))
                                {
                                    resPageType = "01 Untrue";

                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "LastPage");

                                    // DelFirstPage = getPageIndex(lstPPTFinalSettings, "01 Untrue", "FirstPage");

                                    DelFirstPage = DelLastPage;


                                }


                                else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "2" || getNumbersFromString(chart) == "02"))
                                {

                                    resPageType = "02 Mislead";

                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "LastPage");

                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "02 Mislead", "FirstPage");


                                }


                                else if (PageType == "Exaggerative-Inappropriate" && (getNumbersFromString(chart) == "3" || getNumbersFromString(chart) == "03"))
                                {

                                    resPageType = "03 Exagg";

                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "LastPage");

                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "03 Exagg", "FirstPage");


                                }

                                //if (objclsPPT.strPageGroupName.ToLower() == resPageType.ToLower())
                                //{


                                if (resPageType == "01 Untrue")
                                {

                                    op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - 1);
                                }

                                else
                                {
                                    for (int k = 0; k <= DelLastPage - DelFirstPage; k++)
                                    {
                                        op = await clsMisc.DeleteSlideFromPPTAsync1($"C:\\\\excelfiles\\\\{project}\\\\Final\\\\MRRxNaming.pptx", DelLastPage - k - 1);
                                        //Thread.Sleep(100);
                                    }
                                }




                                if (CountSlides(getIndividualChartPath(chart, project, breakDown)) < intChartPages.Count && CountSlides(getIndividualChartPath(chart, project, breakDown)) != 0)
                                {
                                    for (int i = intChartPages.Count - 1; i >= CountSlides(getIndividualChartPath(chart, project, breakDown)); i--)
                                    {
                                        intChartPages.RemoveAt(i);
                                    }

                                }

                                op = await clsMisc.MergeSlideWithSlideArrayAsync1(getIndividualChartPath(chart, project, breakDown), $"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", intChartPages.ToArray(), DelLastPage - 1);



                                if (PageType == "Attribute Evaluation" && chart.Contains("1"))
                                {

                                    resPageType = "Attribute Evaluation 1";
                                    DelLastPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "LastPage");
                                    DelFirstPage = getPageIndex(lstPPTFinalSettings, "Attribute Evaluation 1", "FirstPage");

                                    //update the slide 38

                                    //update the slide 38

                                    string repText = "";
                                    System.Data.DataTable dt1 = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getAttributeEvaluationTitle] " + "'" + project + "'," + "'" + chart + "'");

                                    foreach (DataRow row in dt1.Rows)
                                    {
                                        repText = Convert.ToString(row["strAttributeEvaluationTitle"]);
                                    }






                                    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "<<AttributeEvaluationTitle1>>", repText, 38);

                                    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "Attribute", repText, 38);

                                    await clsMisc.repTextInSlideAsync("C:\\excelfiles\\RACKEM\\Final\\MRRxNaming.pptx", "AttributeEvaluationTitle1", repText, 38);



                                    chartsCompCount = chartsCompCount + 1;

                                    lstchartDispPageGroupName.Add(new clsChartDislayNamePageGroupName(chart, objclsPPT.strPageGroupName));

                                    chartsCompleted.Add(chart);

                                    pageGroupNameCompleted.Add(objclsPPT.strPageGroupName);

                                }
                                // }




                            }









                        }





                    }



                }










            }





            //get the final file in  the folder

            createFolder($"C:\\excelfiles\\{project}\\PPT");
            createFolder($"C:\\excelfiles\\{project}\\PPT\\01 Final");


            while (true)
            {


                if (File.Exists($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx"))
                {
                    copyFile($"C:\\excelfiles\\{project}\\Final\\MRRxNaming.pptx", $"C:\\excelfiles\\{project}\\PPT\\01 Final\\ {project}_{finalTemplate}_{breakDown}.pptx");


                    break;

                }

            }








        }



        private int getPageIndex(List<clsPPTFinalSettings> lst, string pageGroupName, string pageIndex)
        {
            int result = 0;

            if (pageIndex == "LastPage")
            {
                result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexLast;
            }

            else if (pageIndex == "FirstPage")
            {

                result = lst.FirstOrDefault(item => item.strPageGroupName == pageGroupName).intPPTSlideIndexFirst;
            }

            return result;
        }


        private bool checkAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName)
        {

            bool result = false;
            if (lstPageDispName.Any(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true))
            {

                result = true;

            }

            return result;

        }

        private bool checkParticularAttributeEvaluationIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
        {

            bool result = false;

            var lst = lstPageDispName.Where(m => m.strPageType == "Attribute Evaluation" && m.isReportSelectedByUser == true);

            foreach (clschartPageDisplayName obj in lst)
            {
                if (obj.strPageName.Contains(attValue.ToString()))
                {

                    result = true;

                }

            }

            return result;

        }


        private bool checkParticularExaggIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue)
        {

            bool result = false;

            var lst = lstPageDispName.Where(m => m.strPageType == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true && m.strPageName.Contains(attValue.ToString()));

            foreach (clschartPageDisplayName obj in lst)
            {
                if (obj.strPageName.Contains(attValue.ToString()))
                {
                    result = true;

                }

            }


            return result;

        }


        public static string getNumbersFromString(string input)
        {
            // Use a regular expression to match all digits in the string
            string numbers = Regex.Replace(input, @"[^0-9]", "");
            return numbers;
        }


        private bool checkParticularChartIsSelected(List<clschartPageDisplayName> lstPageDispName, int attValue, string chart)
        {

            bool result = false;

            var lst = lstPageDispName.Where(m => m.strPageName == "Exaggerative-Inappropriate" && m.isReportSelectedByUser == true);

            foreach (clschartPageDisplayName obj in lst)
            {
                if (obj.strPageName.Contains(attValue.ToString()))
                {
                    result = true;

                }

            }


            return result;

        }




        private string getReverseChartName(string chartName)
        {
            if (chartName == "Attribute Evaluation 1")
            {
                return "Attribute 1";
            }
            else if (chartName == "Attribute Evaluation 2")
            {
                return "Attribute 2";
            }

            else if (chartName == "Attribute Evaluation 3")
            {
                return "Attribute 3";
            }

            else if (chartName == "Attribute Evaluation 4")
            {
                return "Attribute 4";
            }


            else
            {
                return chartName;
            }

        }



        private int getmaxSlideid(string fileName)
        {
            int maxSlideNumber = 0;
            if (!File.Exists(fileName))
            {
                return 0;
            }

            else
            {
                maxSlideNumber = 1;
                using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
                {

                    PresentationPart presentationPart = doc.PresentationPart;

                    maxSlideNumber = presentationPart.SlideParts.Count();
                }

                return Convert.ToInt32(maxSlideNumber);


            }




        }



        private string getIndividualChartPath(string chart, string projectName, string breakDown)
        {



            string outputpath = $"C:\\excelfiles\\{projectName}\\{chart}_{breakDown}.pptx";


            //get the right chart name 

            return outputpath;

        }

        public int CountSlides(string presentationFile)
        {

            if (File.Exists(presentationFile))
            {

                using (PresentationDocument ppt = PresentationDocument.Open(presentationFile, false))
                {
                    // Get the presentation part of the presentation document
                    PresentationPart presentationPart = ppt.PresentationPart;
                    if (presentationPart == null || presentationPart.Presentation == null)
                    {
                        return 0;
                    }

                    // Get the slide ID list
                    SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
                    if (slideIdList == null)
                    {
                        return 0;
                    }

                    // Return the count of slide IDs
                    return slideIdList.ChildElements.Count;
                }




            }

            else
            {
                return 0;

            }




        }



        private string getRightChartName(string pageName, int index)
        {
            if (pageName == "Attribute Evaluation" && index == 1)
            {
                return "Attribute Evaluation " + index;
            }
            else if (pageName == "Attribute Evaluation" && index == 2)
            {
                return "Attribute Evaluation " + index;
            }
            else if (pageName == "Attribute Evaluation" && index == 3)
            {
                return "Attribute Evaluation " + index;
            }

            else if (pageName == "Attribute Evaluation" && index == 3)
            {
                return "Attribute Evaluation " + index;
            }

            else if (pageName == "Exaggerative-Inappropriate" && index == 1)
            {
                return "01 Untrue";
            }

            else if (pageName == "Exaggerative-Inappropriate" && index == 2)
            {
                return "02 Misleading";
            }

            else if (pageName == "Exaggerative-Inappropriate" && index == 3)
            {
                return "03 Exagg";
            }

            else
            {
                return pageName;
            }

        }


        private string getChartsPagegroupName(string projectName, string chartName)
        {
            string realchartName = "";

            System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");



            List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
            {
                strPageGroup = row.Field<string>("strPageGroup"),
                strPageGroupType = row.Field<string>("strPageGroupType")
            }).ToList();


            List<clschartPageGroup> lstPgChart = lstPg.Where(p => p.strPageGroup == chartName).ToList();

            foreach (clschartPageGroup obj in lstPgChart)
            {

                realchartName = obj.strPageGroupType;

            }

            return realchartName;

        }



        private List<clschartPageGroup> getProjectPagegroupNames(string projectName)
        {
            string realchartName = "";

            System.Data.DataTable dt = clsData.MRData.getDataTable("[xlv1].[getPageGroupsForProject]  " + "'" + projectName + "'");


            List<clschartPageGroup> lstPg = dt.AsEnumerable().Select(row => new clschartPageGroup
            {
                strPageGroup = row.Field<string>("strPageGroup"),
                strPageGroupType = row.Field<string>("strPageGroupType")
            }).ToList();


            return lstPg;

        }


        private List<clschartPageDisplayName> getProjectChartDisplayNames(string projectName)
        {
            string realchartName = "";

            System.Data.DataTable dt = clsData.MRData.getDataTable("[dbo].[ExcelChartsPrc_getPages]  " + 1 + ",'" + projectName + "'");

            List<clschartPageDisplayName> lstPg = dt.AsEnumerable().Select(row => new clschartPageDisplayName
            {
                strPageName = row.Field<string>("strPageName"),
                strPageType = row.Field<string>("strPageType"),
                intPageID = row.Field<Int32>("intPageID").ToString(),
                intPageIndex = row.Field<Int32>("intPageIndex").ToString(),
                intSurveyID = row.Field<Int32>("intSurveyID").ToString(),

            }).ToList();


            return lstPg;

        }


        public static void createFolder(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(path);
                }
            }
            catch (IOException ioex)
            {
                Console.WriteLine(ioex.Message);
            }


        }


        public static void copyFile(string source, string destination)
        {
            try
            {
                File.Copy(source, destination, true);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }


        public static void notAvailable(string sourcePath, string template, string breakdown)
        {
            DLLClass dLLClass = new DLLClass();
            APIWrapper obj = new APIWrapper();
            dLLClass.NotAvailableMethod(sourcePath, template, breakdown);

        }


        public static void copyTemplates()
        {

            //lok please check currently doing only for the first time

            string sourceDirectory = @"\\miafs02\Market Research\MR Programs\ExcelCharts_Chartsdll\Templates";
            string destinationDirectory = "C:\\ExcelChartFiles\\Templates\\";

            if (!Directory.Exists("C:\\ExcelChartFiles\\"))
            {

                if (!Directory.Exists("C:\\ExcelChartFiles\\"))
                {
                    // Create the directory if it doesn't exist
                    Directory.CreateDirectory("C:\\ExcelChartFiles\\");
                }

                if (!Directory.Exists(destinationDirectory))
                {
                    // Create the directory if it doesn't exist
                    Directory.CreateDirectory(destinationDirectory);

                }

                foreach (var file in Directory.GetFiles(sourceDirectory))
                {
                    File.Copy(file, Path.Combine(destinationDirectory, Path.GetFileName(file)), true);
                }


            }


        }


        public void notAvailable(string template)
        {

            //DLLClass obj = new DLLClass();
            //obj.NotAvailableMethod(CreateTargetPath($"C:\\ExcelChartFiles\\Templates\\NotAvailable.pptx"), template);

            //template.NotAvailableTemplate(destination, templateName, breakdown);

        }


        public void copyInsertSlide(string source, string destination, int pos)
        {
            clsMisc.CopySlide(source, destination, pos);

        }








    }
}
