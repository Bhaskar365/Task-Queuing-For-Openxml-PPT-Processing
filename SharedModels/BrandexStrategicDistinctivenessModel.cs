using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels
{
    public class BrandexStrategicDistinctivenessModel
    {
        public string strTestName { get; set; } = string.Empty;
        public string strTestNameTranslation { get; set; } = string.Empty;

        public double? dblAveragePage1 { get; set; }
        public double? dblPage1Weight { get; set; }
        public double? dblAveragePage1Weighted { get; set; }

        public double? dblAveragePage2 { get; set; }
        public double? dblPage2Weight { get; set; }
        public double? dblAveragePage2Weighted { get; set; }

        public double? dblAveragePage3 { get; set; }
        public double? dblPage3Weight { get; set; }
        public double? dblAveragePage3Weighted { get; set; }

        public double? dblAveragePage4 { get; set; }
        public double? dblPage4Weight { get; set; }
        public double? dblAveragePage4Weighted { get; set; }

        public double? dblAveragePage5 { get; set; }
        public double? dblPage5Weight { get; set; }
        public double? dblAveragePage5Weighted { get; set; }

        public double? dblAveragePage6 { get; set; }
        public double? dblPage6Weight { get; set; }
        public double? dblAveragePage6Weighted { get; set; }

        public double? dblAveragePage7 { get; set; }
        public double? dblPage7Weight { get; set; }
        public double? dblAveragePage7Weighted { get; set; }

        public double? dblAveragePage8 { get; set; }
        public double? dblPage8Weight { get; set; }
        public double? dblAveragePage8Weighted { get; set; }

        public double? dblIndex { get; set; }
        public string strDSIScore { get; set; } = string.Empty;
        public double? intRed { get; set; }
        public double? intGreen { get; set; }
        public double? intBlue { get; set; }
        public bool? boolBold { get; set; }
        public string ProjectTemplateType { get; set; } = string.Empty;
    }

    public class BrandexStrategicDistinctivenessShortModel
    {
        public string strTestName { get; set; } = string.Empty;

        public double? dblAveragePage1 { get; set; }
        public double? dblPage1Weight { get; set; }
        public double? dblAveragePage1Weighted { get; set; }
        public double? dblAveragePage1WeightedForDistinctivenessChart { get; set; }
        public double? dblAveragePage1WeightedForMarketingChart { get; set; }

        public double? dblAveragePage2 { get; set; }
        public double? dblPage2Weight { get; set; }
        public double? dblAveragePage2Weighted { get; set; }
        public double? dblAveragePage2WeightedForDistinctivenessChart { get; set; }
        public double? dblAveragePage2WeightedForMarketingChart { get; set; }

        public double? dblAveragePage3 { get; set; }
        public double? dblPage3Weight { get; set; }
        public double? dblAveragePage3Weighted { get; set; }
        public double? dblAveragePage3WeightedForDistinctivenessChart { get; set; }
        public double? dblAveragePage3WeightedForMarketingChart { get; set; }

        public double? dblAveragePage4 { get; set; }
        public double? dblPage4Weight { get; set; }
        public double? dblAveragePage4Weighted { get; set; }
        public double? dblAveragePage4WeightedForDistinctivenessChart { get; set; }
        public double? dblAveragePage4WeightedForMarketingChart { get; set; }

        public double? dblAveragePage5 { get; set; }
        public double? dblPage5Weight { get; set; }
        public double? dblAveragePage5Weighted { get; set; }

        public double? dblAveragePage6 { get; set; }
        public double? dblPage6Weight { get; set; }
        public double? dblAveragePage6Weighted { get; set; }

        public double? dblAveragePage7 { get; set; }
        public double? dblPage7Weight { get; set; }
        public double? dblAveragePage7Weighted { get; set; }

        public double? dblAveragePage8 { get; set; }
        public double? dblPage8Weight { get; set; }
        public double? dblAveragePage8Weighted { get; set; }

        public double? dblIndex { get; set; }
        public double? dblIndexForDistinctivenessChart { get; set; }
        public double? dblIndexForMarketingChart { get; set; }
        public string strDSIScore { get; set; } = string.Empty;
        public double? intRed { get; set; }
        public double? intGreen { get; set; }
        public double? intBlue { get; set; }
        public bool? boolBold { get; set; }
    }
}
