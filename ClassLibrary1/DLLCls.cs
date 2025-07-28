using SharedModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApplicationAPI.Models;

namespace ClassLibrary1
{
    public class DLLCls
    {
        DllTemplates template = new DllTemplates();

        public void FitToConceptMethod(string destination, List<FitToConceptModel> apiData)
        {
            template.FitToConceptTemplateForAll(destination, apiData);
        }

        public void OverallImpressionsMethod(string destination, List<OverallImpressionsModel> apiData)
        {
            template.OverallImpressionsTemplateForAll(destination, apiData);
        }

        public void Attribute1Method(string destination,
                                       List<Aev1> apiData)
        {
            template.Attribute1TemplateForAll(destination, apiData);
        }

        public void Attribute2Method(string destination,
                                       List<Aev2> apiData)
        {
            template.Attribute2TemplateForAll(destination, apiData);
        }

        public void AttributeMethodForAttributeEvalAggreg(string destination, List<Aev3> apiData)
        {
            template.AttributeEvalAggregTemplateForAll(destination, apiData);
        }

        public void MemorabilityMethod(string destination, List<Memorability> apiData)
        {
            template.MemorabilityTemplateForAll(destination, apiData);
        }

        public void PersonalPreferencesMethod(string destination, List<PersonalPreference> apiData)
        {
            template.PersonalPreferencesTemplateForAll(destination, apiData);
        }


        public void SuffixMethod(string destination, List<Suffix> apiData)
        {
            template.SuffixTemplateForAll(destination, apiData);
        }

        public void VerbalUnderstandingBarMethod(string destination, List<VerbalUnderstanding> apiData)
        {
            template.VerbalUnderstandingBarTemplateForAll(destination, apiData);
        }

        public void WrittenUnderstandingMethod(string destination, List<WrittenUnderstanding> apiData)
        {
            template.WrittenUnderstandingBarTemplateForAll(destination, apiData);
        }

        public void ExaggerativeMethod(string destination, List<Likeability> apiData)
        {
            template.ExaggerativeTemplateForAll(destination, apiData);
        }

        public void SALANewMethod(string destination, List<Sala> apiData)
        {
            template.SALANewModelTemplateSpecial2(destination, apiData);
        }

        public void QTCMethod(string destination, List<QTCModel> apiData)
        {
            template.QTCTemplate(destination, apiData);
        }

        public void BrandexSafetyMethod(string destination, List<BrandexSafetyShortModel> apiData)
        {
            template.BrandexSafetyTemplate(destination, apiData);
        }

        public void BrandexStrategicDistinctivenessMethod(string destination, List<BrandexStrategicDistinctivenessShortModel> apiData)
        {
            template.BrandexStrategicDistinctivenessTemplate(destination, apiData);
        }


        public void MedicalTermsMethod(string destination, List<MedicalTermsModel> apiData)
        {
            template.MedicalTermsTemplate6(destination, apiData);
        }

        public void NonMedicalTermsMethod(string destination, List<NonMedicalTermsModel> apiData)
        {
            template.NonMedicalTermsTemplateSpecial3(destination, apiData);
        }


    }
}
