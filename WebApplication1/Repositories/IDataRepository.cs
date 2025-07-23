using SharedModels;
using WebApplicationAPI.Models;

namespace WebApplication1.Repositories
{
    public interface IDataRepository
    {
        Task<IEnumerable<Aev1>> GetAtt1Data();

        Task<IEnumerable<Aev2>> GetAtt2Data();

        Task<IEnumerable<Aev3>> GetAttrEvalAggregData();
     
        Task<IEnumerable<FitToConceptModel>> GetAllFitToConceptData();
        
        Task<IEnumerable<Memorability>> GetMemorabilityData();
        
        Task<IEnumerable<OverallImpressionsModel>> GetAllOverallImpressionsData();
        
        Task<IEnumerable<PersonalPreference>> GetPersonalPreferenceData();

        Task<IEnumerable<Suffix>> GetSuffixData();

        Task<IEnumerable<VerbalUnderstanding>> GetVerbalUnderstandingData();

        Task<IEnumerable<WrittenUnderstanding>> GetWrittenUnderstandingData();

        Task<IEnumerable<Likeability>> GetExagg();

        Task<IEnumerable<Sala>> GetSala();

        Task<IEnumerable<Sala154>> GetSala154();

        

    }
}
