using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class WorkloadEntity
    {
        public int Id;
        public string Title;
        public string AssignedTo;
        public double SumHours;
        public double OverTimes;
        public string SupperType;
        public string Type;
        public string CreatedDate;
        public string WorkDate;
        public string InPlaned;
        public string TeamProject;

    }
}
