using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace SharedModels
{
    public class FitToConceptModel
    {
        [Key]
        public int Id { get; set; }
        public string? ProjectName { get; set; }
        public string? TestName { get; set; }
        public string? TestNameColor { get; set; }
        public double Average { get; set; }
        public bool TestNameBold { get; set; }
        public string? AverageColor { get; set; }
        public string? ProjectTemplateType { get; set; }

    }

    public enum ProcessState
    {
        Idle,
        Generating,
        Done
    }



}
