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

        public async Task<List<TaskLog>> GetLogs()
        {
            return await _httpClient.GetFromJsonAsync<List<TaskLog>>($"api/report/user/taskLogs");
        }
    }
}
