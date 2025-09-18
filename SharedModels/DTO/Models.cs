using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedModels.DTO
{
    public class Models
    {

    }

    public class IndividualReportModel
    {
        [Key]
        public int SubtaskId { get; set; }

        [ForeignKey("TaskID")]
        public Guid TaskID { get; set; }

        [ForeignKey("UserID")]
        public int UserID { get; set; }

        [ForeignKey("StatusID")]
        public int StatusID { get; set; }

        public string TemplateName { get; set; }

        public DateTime CreatedOn { get; set; }

        public DateTime CompletedOn { get; set; }

        public string StatusMessage { get; set; }

    }

    public class UsersModel
    {
        [Key]
        public int UserID { get; set; }

        public string UserName { get; set; }

        public int RoleID { get; set; }
    }

    public class UserRoleModel
    {
        public int RoleID { get; set; }

        public string RoleName { get; set; }
    }

    public class StatusModel
    {
        public int StatusID { get; set; }
        public string StatusName { get; set; }
    }

    public class FinalReportModel
    {
        [Key]
        public Guid TaskID { get; set; }

        public string ProjectName { get; set; }

        public DateTime CreatedOn { get; set; }

        public DateTime CompletedOn { get; set; }

        [ForeignKey("UserID")]
        public int UserID { get; set; }

        [ForeignKey("StatusID")]
        public int StatusID { get; set; }
    }

    public class FinalReportModelDTO
    {
        [Key]
        public Guid TaskID { get; set; }

        public string ProjectName { get; set; }

        public DateTime CreatedOn { get; set; }

        public DateTime CompletedOn { get; set; }

        [ForeignKey("UserID")]
        public int UserID { get; set; }

        public string UserName { get; set; }

        [ForeignKey("StatusID")]
        public int StatusID { get; set; }

        public string StatusName { get; set; }
    }
}
