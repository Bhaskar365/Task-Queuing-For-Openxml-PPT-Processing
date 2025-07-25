using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels
{
    public class BrandexSafetyModel
    {
        public string? strTestName { get; set; } = string.Empty;
        public string? strTestNameTranslation { get; set; } = string.Empty;

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

        public double dblIndex { get; set; }
        public string? strDSIScore { get; set; } = string.Empty;
        public int? intRed { get; set; }
        public int? intGreen { get; set; }
        public int? intBlue { get; set; }
        public bool? boolBold { get; set; }
        public string? ProjectTemplateType { get; set; } = string.Empty;
    }


    public class BrandexSafetyShortModel
    {
        public string strTestName { get; set; }

        public double dblAveragePage1 { get; set; }
        public double dblPage1Weight { get; set; }
        public double dblAveragePage1Weighted { get; set; }
        public double dblAveragePage1WeightedForChart { get; set; }

        public double dblAveragePage2 { get; set; }
        public double dblPage2Weight { get; set; }
        public double dblAveragePage2Weighted { get; set; }
        public double dblAveragePage2WeightedForChart { get; set; }

        public double dblAveragePage3 { get; set; }
        public double dblPage3Weight { get; set; }
        public double dblAveragePage3Weighted { get; set; }
        public double dblAveragePage3WeightedForChart { get; set; }

        public double dblAveragePage4 { get; set; }
        public double dblPage4Weight { get; set; }
        public double dblAveragePage4Weighted { get; set; }
        public double dblAveragePage4WeightedForChart { get; set; }

        public double dblAveragePage5 { get; set; }
        public double dblPage5Weight { get; set; }
        public double dblAveragePage5Weighted { get; set; }
        public double dblAveragePage5WeightedForChart { get; set; }

        public double dblAveragePage6 { get; set; }
        public double dblPage6Weight { get; set; }
        public double dblAveragePage6Weighted { get; set; }

        public double dblAveragePage7 { get; set; }
        public double dblPage7Weight { get; set; }
        public double dblAveragePage7Weighted { get; set; }

        public double dblAveragePage8 { get; set; }
        public double dblPage8Weight { get; set; }
        public double dblAveragePage8Weighted { get; set; }

        public double dblIndex { get; set; }
        public double dblIndexForChart { get; set; }
        public string strDSIScore { get; set; }
        public int intRed { get; set; }
        public int intGreen { get; set; }
        public int intBlue { get; set; }
        public bool boolBold { get; set; }
    }


}
