using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.DTO
{
    public class ReportGenerationRequest
    {
        public string ProjectTemplateType { get; set; } = string.Empty;
        public Guid TaskId { get; set; } = Guid.NewGuid();
    }

    public class ReportGenerationRequestDLL
    {
        public string project { get; set; } = string.Empty;

        public List<string>? templates { get; set; } 

        public List<string>? breakdowns { get; set; }

        public string HistoricalMeanType { get; set; } = string.Empty;

        public string HistoricalMeanDescription { get; set; } = string.Empty;

        public Guid TaskId { get; set; } = Guid.NewGuid();
    }




    public class ReportGenerationResponse
    {
        public Guid TaskId { get; set; }
    }

    public class ReportStatusDto
    {
        public Guid TaskId { get; set; }
        public string Status { get; set; } = string.Empty;

        public string ProjectType { get; set; } = string.Empty;
    }




    //public class APIRequestModel
    //{
    //    public string project { get; set; } = string.Empty;
    //   public string template { get; set; } = string.Empty;
    //    public string group { get; set; } = string.Empty;
    //    public string breakdown { get; set; } = string.Empty;
    //    public string HistoricalMeanType { get; set; } = string.Empty;
    //    public string HistoricalMeanDescription { get; set; } = string.Empty;
    //}

    public class APIRequestModel
    {
        public string project = "";
        public string template = "";
        public string group = "";
        public string breakdown = "";
        public string HistoricalMeanType = "";
        public string HistoricalMeanDescription = "";

        public APIRequestModel() { }

        public APIRequestModel(string project,
                          string template,
                          string group,
                          string breakdown,
                          string HistoricalMeanType,
                          string HistoricalMeanDescription)
        {
            this.project = project;
            this.template = template;
            this.group = group;
            this.breakdown = breakdown;
            this.HistoricalMeanDescription = HistoricalMeanDescription;
            this.HistoricalMeanType = HistoricalMeanType;
        }
    }

}
