using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary2.Models
{
    public class PageStatusDto
    {
        public Guid TaskId { get; set; }
        public string Status { get; set; } = string.Empty;

        public string ProjectType { get; set; } = string.Empty;
    }
}
