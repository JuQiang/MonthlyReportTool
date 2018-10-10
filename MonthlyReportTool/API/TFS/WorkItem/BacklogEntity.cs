using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class BacklogEntity
    {
        public int Id;
        public string KeyApplicationName;
        public string ModulesName;
        public string FuncName;
        public string Title;
        public string Category;
        public string AssignedTo;
        public string AcceptanceMeasure;
        public string State;
        //public string HopeSubmitTime;
        public string IsPlaned;
        public string CreatedDate;
        public string Tags;
        //public string AcceptTime;
        //public string IsNeedInterfaceTest;
        public string IsNeedPerformanceTest;
        //public string SubmitTime;
        public string TeamProject;
        public string FinishDate;
        public string ParentId;
        
    }
}
