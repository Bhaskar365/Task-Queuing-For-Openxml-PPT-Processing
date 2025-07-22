using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace SharedModels
{
    public class OverallImpressionsModel
    {
        [Key]
        public int Id { get; set; }
        public string? ProjectName { get; set; }
        public string? TestName { get; set; }
        public double Positive { get; set; }
        public double Neutral { get; set; }
        public double Negative { get; set; }
        public bool? TestNameBold { get; set; }
        public string? TestNameColor { get; set; }
        public string? ProjectTemplateType { get; set; }
    }
}
