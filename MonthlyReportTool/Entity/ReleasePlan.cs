using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.Entity
{
    public class ReleasePlan
    {
        public int Id { get; set; }
        public string TeamProject { get; set; }
        public string WorkItemType { get; set; }
        public string State { get; set; }
        public string AssignedTo { get; set; }
        public string Title { get; set; }
        public string ReleaseDate { get; set; }
        public int RollbackCount { get; set; }
        public string PublishType { get; set; }
        public string RollbackReason { get; set; }
    }

}
