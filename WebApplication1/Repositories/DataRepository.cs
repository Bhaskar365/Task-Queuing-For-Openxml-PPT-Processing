
using Microsoft.EntityFrameworkCore;
using SharedModels;
using SharedModels.DTO;
using WebApplication1.Context;
using WebApplicationAPI.Models;

namespace WebApplication1.Repositories
{
    public class DataRepository : IDataRepository
    {
        private readonly AppDBContext _context;

        public DataRepository(AppDBContext context)
        {
            _context = context;
        }
        public async Task<IEnumerable<FitToConceptModel>> GetAllFitToConceptData()
        {
            var result = await _context.FitToConceptTestTable.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<OverallImpressionsModel>> GetAllOverallImpressionsData()
        {
            var result = await _context.OverallImpressionsTestTable.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<Aev1>> GetAtt1Data()
        {
            var result = await _context.Aev1s.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<Aev2>> GetAtt2Data()
        {
            var result = await _context.Aev2s.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<Aev3>> GetAttrEvalAggregData()
        {
            var result = await _context.Aev3s.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<BrandexSafetyModel>> GetBrandexSafety()
        {
            try
            {
                var result = await _context.BrandexSafety.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<BrandexStrategicDistinctivenessModel>> GetBrandexStrategicDistinctivenessData()
        {
            try
            {
                var result = await _context.BrandexStrategicDistinctiveness.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<Likeability>> GetExagg()
        {
            var result = await _context.Likeabilities.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<ExaggModel>> GetExagg03()
        {
            try
            {
                var result = await _context.Exagg03.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<MedicalTermsModel>> GetMedicalTerms()
        {
            try
            {
                var result = await _context.MedicalTerms.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<Memorability>> GetMemorabilityData()
        {
            var result = await _context.Memorabilities.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<ExaggModel>> GetMisleading02()
        {
            try
            {
                var result = await _context.Misleading02.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<NonMedicalTermsModel>> GetNonMedicalTerms()
        {
            try
            {
                var result = await _context.NonMedicalTerms.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<PersonalPreference>> GetPersonalPreferenceData()
        {
            var result = await _context.PersonalPreferences.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<QTCModel>> GetQTC()
        {
            var data = await _context.Qtc.ToListAsync();
            return data;
        }

        public async Task<IEnumerable<Sala>> GetSala()
        {
            var result = await _context.Salas.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<Sala154>> GetSala154()
        {
            var result = await _context.Sala154.ToListAsync();
            return result;
        }


        public async Task<IEnumerable<Suffix>> GetSuffixData()
        {
            try
            {
                var result = await _context.Suffixes.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }          
        }

        public async Task<IEnumerable<TaskLog>> GetTaskLogs()
        {
            try
            {
                var result = await _context.TaskLogs.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<ExaggModel>> GetUntrue01()
        {
            try
            {
                var result = await _context.Untrue01.ToListAsync();
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<IEnumerable<VerbalUnderstanding>> GetVerbalUnderstandingData()
        {
            var result = await _context.VerbalUnderstandings.ToListAsync();
            return result;
        }

        public async Task<IEnumerable<WrittenUnderstanding>> GetWrittenUnderstandingData()
        {
            var result = await _context.WrittenUnderstandings.ToListAsync();
            return result;
        }
    }
}
