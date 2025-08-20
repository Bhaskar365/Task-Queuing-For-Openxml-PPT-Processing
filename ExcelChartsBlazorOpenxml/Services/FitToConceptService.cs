using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.VariantTypes;
using SharedModels;
using SharedModels.DTO;
using WebApplicationAPI.Models;

namespace ExcelChartsBlazorOpenxml.Services
{
    public class FitToConceptService : IFitToConceptService
    {
        private readonly HttpClient _httpClient;

        public FitToConceptService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }
        public async Task<List<FitToConceptModel>> GetFitToConceptData()
        {
            return await _httpClient.GetFromJsonAsync<List<FitToConceptModel>>($"api/data/fitToConcept");
        }

        public async Task<List<OverallImpressionsModel>> GetOverallImpressionsData()
        {
            return await _httpClient.GetFromJsonAsync<List<OverallImpressionsModel>>($"api/data/OverallImpressions");
        }

        public async Task<List<Aev1>> GetAtt1Data()
        {
            return await _httpClient.GetFromJsonAsync<List<Aev1>>($"api/data/Attribute_1");
        }

        public async Task<List<Aev2>> GetAtt2Data()
        {
            return await _httpClient.GetFromJsonAsync<List<Aev2>>($"api/data/Attribute_2");
        }

        public async Task<List<Aev3>> GetAttrAggData()
        {
            return await _httpClient.GetFromJsonAsync<List<Aev3>>($"api/data/Attribute_Aggregrate");
        }

        public async Task<List<Memorability>> MemorabilityData()
        {
            return await _httpClient.GetFromJsonAsync<List<Memorability>>($"api/data/Memorability");
        }

        public async Task<List<PersonalPreference>> PersonalPrefData()
        {
            return await _httpClient.GetFromJsonAsync<List<PersonalPreference>>($"api/data/Personal_Preferences");
        }

        public async Task<List<Suffix>> SuffixData()
        {
            return await _httpClient.GetFromJsonAsync<List<Suffix>>($"api/data/Suffix");
        }

        public async Task<List<VerbalUnderstanding>> VerbalUnderData()
        {
            return await _httpClient.GetFromJsonAsync<List<VerbalUnderstanding>>($"api/data/Verbal_Understandings");
        }

        public async Task<List<WrittenUnderstanding>> WrittenUnderData()
        {
            return await _httpClient.GetFromJsonAsync<List<WrittenUnderstanding>>($"api/data/Written_Understandings");
        }

        public async Task<List<Likeability>> ExaggData()
        {
            return await _httpClient.GetFromJsonAsync<List<Likeability>>($"api/data/Exagg");
        }

        public async Task<List<Sala>> SalaData()
        {
            return await _httpClient.GetFromJsonAsync<List<Sala>>($"api/data/Sala");
        }

        public async Task<List<QTCModel>> QtcData()
        {
            return await _httpClient.GetFromJsonAsync<List<QTCModel>>($"api/data/QTC");
        }

        public async Task<Guid> GenerateReportAsync(ReportGenerationRequest request)
        {
            var response = await _httpClient.PostAsJsonAsync("api/report/generate", request);
            var result = await response.Content.ReadFromJsonAsync<ReportGenerationResponse>();
            return result!.TaskId;
        }

        public async Task<string> GetReportStatusAsync(Guid taskId)
        {
            var response = await _httpClient.GetFromJsonAsync<ReportStatusDto>($"api/report/status/{taskId}");
            return response!.Status;
        }

        public async Task<List<Sala>> Sala154Data()
        {
            return await _httpClient.GetFromJsonAsync<List<Sala>>($"api/data/Sala154");
        }
        public async Task<List<BrandexSafetyModel>> BrandexSafetyData()
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<List<BrandexSafetyModel>>($"api/data/Brandex Safety");
                return response!;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<List<BrandexStrategicDistinctivenessModel>> BrandexStrategicDistinctivenessesData()
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<List<BrandexStrategicDistinctivenessModel>>($"api/data/Brandex Strategic Distinctiveness");
                return response!;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<List<MedicalTermsModel>> MedicalTermsData()
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<List<MedicalTermsModel>>($"api/data/Medical Terms");
                return response!;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<List<NonMedicalTermsModel>> NonMedicalTermsData()
        {
            try
            {
                var response = await _httpClient.GetFromJsonAsync<List<NonMedicalTermsModel>>($"api/data/Non-medical Terms");
                return response!;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<List<UntrueModel>> UntrueData()
        {
            var response = await _httpClient.GetFromJsonAsync<List<UntrueModel>>($"api/data/Untrue01");
            return response;
        }

        public async Task<List<MisleadingModel>> MisleadingData()
        {
            var response = await _httpClient.GetFromJsonAsync<List<MisleadingModel>>($"api/data/Misleading02");
            return response;
        }

        public async Task<List<Exagg03Model>> Exagg03Data()
        {
            var response = await _httpClient.GetFromJsonAsync<List<Exagg03Model>>($"api/data/Exagg03");
            return response;
        }


        public async Task<List<TaskLog>> GetLogs()
        {
            return await _httpClient.GetFromJsonAsync<List<TaskLog>>($"api/report/user/taskLogs");
        }

        public async Task<List<TaskLog>> GetUnfinishedLogs(string user)
        {
            var response = await _httpClient.GetFromJsonAsync<List<TaskLog>>($"api/report/{user}/logs");
            return response!;
        }

        public async Task<List<TaskLog>> GetUserLogs(string user)
        {
            var response = await _httpClient.GetFromJsonAsync<List<TaskLog>>($"api/report/user/taskLogs");

            var userData = response!.Where(x => x.CreatedBy == user).ToList();

            return userData;
        }

        public async Task<List<TaskLog>> MergeSlides(List<TaskLog> taskLog)
        {
            var response = await _httpClient.PostAsJsonAsync<List<TaskLog>>($"api/report/ppt/merge", taskLog);
            var result = await response.Content.ReadFromJsonAsync<List<TaskLog>>();
            return result!;
        }

        public async Task<Guid> GenerateReportUsingDLLAsync(ReportGenerationRequestDLL request)
        {
            var response = await _httpClient.PostAsJsonAsync("api/report/dllgenerate", request);
            var result = await response.Content.ReadFromJsonAsync<ReportGenerationResponse>();
            return result!.TaskId;
        }

        public async Task<string> SendDLLMergeRequest(List<APIRequestModel> projectWrapperAPIList)
        {
            try
            {
                var response = await _httpClient.PostAsJsonAsync("api/report/ppt/dllMerge", projectWrapperAPIList);

                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    return "Successful";
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                {
                    return "Bad Request";
                }

                return "Fail";

            }
            catch (Exception)
            {
                throw;
            }

            //var r = response.Content.ReadFromJsonAsync<string>();

            //Console.WriteLine(r);

            //var result = await response.Content.ReadFromJsonAsync<List<APIRequestModel>>();
            //return result.ToString();

            //var result = await response.Content.ReadFromJsonAsync<List<TaskLog>>();
            //return result!;
        }
        

        public async Task<string> SendDLLMergeRequestWithPanel(ReportGenerationRequestDLL projectWrapperAPIList)
        {
            try
            {
                var response = await _httpClient.PostAsJsonAsync("api/report/ppt/dllMerge", projectWrapperAPIList);

                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    return "Successful";
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                {
                    return "Bad Request";
                }

                return "Fail";

            }
            catch (Exception)
            {
                throw;
            }

            //var r = response.Content.ReadFromJsonAsync<string>();

            //Console.WriteLine(r);

            //var result = await response.Content.ReadFromJsonAsync<List<APIRequestModel>>();
            //return result.ToString();

            //var result = await response.Content.ReadFromJsonAsync<List<TaskLog>>();
            //return result!;
        }


        //public async Task SendDLLMergeRequest(ReportGenerationRequestDLL request)
        //{
        //    try
        //    {
        //        var response = await _httpClient.PostAsJsonAsync("api/report/ppt/dllMerge", request);
        //        Console.WriteLine(response);
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}

    }
}
