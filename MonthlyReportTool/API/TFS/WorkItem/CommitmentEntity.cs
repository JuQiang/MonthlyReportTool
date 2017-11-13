using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class CommitmentEntity
    {
        public int Id;
        public string Title;
        public string AssignedTo;
        public string MonthState;        
        public string State;
        public string DevelopmentFinishedDate;
        public string ReleaseFinishedDate;
        public string TeamProject;
        public string InitTargetDate;
        public string TargetDate;        
        public bool IsDevelopment;

    }
}
