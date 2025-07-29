using SharedModels;
using SharedModels.DTO;
using WebApplicationAPI.Models;

namespace ExcelChartsBlazorOpenxml.Services
{
    public interface IFitToConceptService
    {
        Task<List<FitToConceptModel>> GetFitToConceptData();

        Task<List<OverallImpressionsModel>> GetOverallImpressionsData();

        Task<List<Aev1>> GetAtt1Data();

        Task<List<Aev2>> GetAtt2Data();

        Task<List<Aev3>> GetAttrAggData();

        Task<List<Memorability>> MemorabilityData();

        Task<List<PersonalPreference>> PersonalPrefData();

        Task<List<Suffix>> SuffixData();

        Task<List<VerbalUnderstanding>> VerbalUnderData();

        Task<List<WrittenUnderstanding>> WrittenUnderData();

        Task<List<Likeability>> ExaggData();

        Task<List<Sala>> SalaData();

        Task<List<Sala>> Sala154Data();

        Task<List<QTCModel>> QtcData();

        Task<List<UntrueModel>> UntrueData();

        Task<List<MisleadingModel>> MisleadingData();

        Task<List<Exagg03Model>> Exagg03Data();


        Task<Guid> GenerateReportAsync(ReportGenerationRequest request);
        Task<string> GetReportStatusAsync(Guid taskId);

        Task<List<TaskLog>> GetLogs();

        Task<List<TaskLog>> GetUnfinishedLogs(string user);

    }
}
