using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.DTO
{
    public class TaskLog
    {
        [Key]
        public int Id { get; set; }
        public string ProjectType { get; set; } = string.Empty;
        public Guid TaskId { get; set; }
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        public DateTime? CompletedOn { get; set; } = null;
        public string CreatedBy { get; set; } = string.Empty;
        public string CurrentStatus { get; set; } = string.Empty;

    }
}
