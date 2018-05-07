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
        public string KeyApplicationName;
        public string ModulesName;
        public string FuncName;
        public string SubmitType;
        public string Title;
        public string State;
        public string SubmitUser;
        public string AssignedTo;
        public int BackNum;
        public bool IsNeedPerformanceTest;
        public string TestFinishedTime;
        public string SubmitDate;
        public string PlanTestFinishedTime;
        public string AcceptTime;
        public string CreatedDate;
        public string BackType;
        public string IterationPath;
        public int SubmitNumberOfTime;
        public string TeamProject;
        public int FindedBugCount;
    }
    //    select[System.Id], [Teld.Scrum.KeyApplication], [Teld.Scrum.ModulesName], [Teld.Scrum.Worklog.SubmitLog.SubmitType], 
    //[System.Title], [System.State], [Teld.Scrum.Worklog.SubmitLog.SubmitUser], [System.AssignedTo], 
    //[Teld.Scrum.Worklog.SubmitLog.BackNum], [Teld.Scrum.Backlog.IsNeedPerformanceTest], [Teld.Scrum.TestFinishedTime], 
    //[Teld.Scrum.Worklog.SubmitLog.SubmitDate], [Teld.Scrum.Backlog.PlanTestFinishedTime], [Teld.Scrum.Backlog.AcceptTime], 
    //[System.CreatedDate], [Teld.Scrum.Worklog.SubmitLog.BackType], [System.IterationPath], [Teld.Scrum.SubmitNumberOfTime], 
    //[Teld.Bug.FunctionMenu], [System.TeamProject],[Teld.Scrum.FindedBugCount]
    //    from WorkItems where[System.WorkItemType] = '提交单' 
    //        and[System.TeamProject] = 'TTP' and[System.IterationPath] = 'TTP\FYQ4\Sprint35' 
    //        and[Teld.Scrum.Worklog.SubmitLog.SubmitType] <> '运维SQL' order by[Teld.Scrum.TestFinishedTime]
}
