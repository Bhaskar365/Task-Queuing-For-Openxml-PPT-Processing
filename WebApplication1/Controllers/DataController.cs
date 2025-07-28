using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.InteropServices;
using WebApplication1.Repositories;
using WebApplicationAPI.Queueing;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DataController : ControllerBase
    {
        private readonly IDataRepository _repository;
        private readonly IBackgroundTaskQueue _queue;

        public DataController(IDataRepository repository, IBackgroundTaskQueue queue)
        {
            _repository = repository;
            _queue = queue;
        }

        [HttpGet("fitToConcept")]
        public async Task<ActionResult> GetFitToConceptData()
        {
            try
            {
                return Ok(await _repository.GetAllFitToConceptData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("OverallImpressions")]
        public async Task<ActionResult> GetOverallImpressionsData()
        {
            try
            {
                return Ok(await _repository.GetAllOverallImpressionsData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Attribute_1")]
        public async Task<ActionResult> GetAttribute1Data()
        {
            try
            {
                return Ok(await _repository.GetAtt1Data());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Attribute_2")]
        public async Task<ActionResult> GetAttribute2Data()
        {
            try
            {
                return Ok(await _repository.GetAtt2Data());
            }
            catch (Exception ex)
            { 
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Attribute_Aggregrate")]
        public async Task<ActionResult> GetAttributeAggregData()
        {
            try
            {
                return Ok(await _repository.GetAttrEvalAggregData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Memorability")]
        public async Task<ActionResult> GetMemorability()
        {
            try
            {
                return Ok(await _repository.GetMemorabilityData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Personal_Preferences")]
        public async Task<ActionResult> GetPersonalPrefData()
        {
            try
            {
                return Ok(await _repository.GetPersonalPreferenceData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Suffix")]
        public async Task<ActionResult> GetSuffixesData()
        {
            try
            {
                return Ok(await _repository.GetSuffixData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Verbal_Understandings")]
        public async Task<ActionResult> GetVerbalUndData()
        {
            try
            {
                return Ok(await _repository.GetVerbalUnderstandingData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Written_Understandings")]
        public async Task<ActionResult> GetWrittenUndData()
        {
            try
            {
                return Ok(await _repository.GetWrittenUnderstandingData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Exagg")]
        public async Task<ActionResult> GetLikeabilityData()
        {
            try
            {
                return Ok(await _repository.GetExagg());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Sala")]
        public async Task<ActionResult> GetSalaData()
        {
            try
            {
                return Ok(await _repository.GetSala());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Sala154")]
        public async Task<ActionResult> GetSala154Data()
        {
            try
            {
                return Ok(await _repository.GetSala154());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("QTC")]
        public async Task<ActionResult> GetQTCData()
        {
            try
            {
                return Ok(await _repository.GetQTC());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Brandex Safety")]
        public async Task<ActionResult> GetBrandexSafetyData()
        {
            try
            {
                return Ok(await _repository.GetBrandexSafety());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Brandex Strategic Distinctiveness")]
        public async Task<ActionResult> GetBrandexStrategicDistinctivenessData() 
        {
            try
            {
                return Ok(await _repository.GetBrandexStrategicDistinctivenessData());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        [HttpGet("Medical Terms")]
        public async Task<ActionResult> GetMedicalTermsData()
        {
            try
            {
                return Ok(await _repository.GetMedicalTerms());
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }
}
