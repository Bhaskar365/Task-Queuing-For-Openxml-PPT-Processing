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

}
