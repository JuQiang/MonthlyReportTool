using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class WorkReviewEntity
    {
        public int Id;
        public string workItemType;
        public string KeyApplicationName;
        public string ModulesName;
        public string FuncName;
        public string Title;
        public string State;
        public string AssignedTo;
        public string ParentId;
        public string TeamProject;

        public string ReviewBillType;
        public string ReviewResponsibleMan;
        public string PlanSubmitDate;
        public string ActionDate;
        public string CreatedDate;
        public string ClosedDate;
        public string IterationPath;
        public int FindedBugCount;

        //把Bug的一些信息也设置一下
        public string Type;
        public string Severity;
        public string DetectionMode;
        public string DiscoveryUser;

    }
}
