using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class BugEntity
    {
        public int Id;
        public string KeyApplication;
        public string ModulesName;
        public string Title;
        public string AssignedTo;
        public string State;
        public string Type;
        public string Severity;
        public string ResolvedReason;
        public string Envir;
        public string CreatedDate;
        public string ChangedDate;
        public string DetectionMode;
        public string DetectionPhase;
        public string HopeFixSubmitTime;
        public string TeamProject;
        public string CreatedBy;
        public string IterationPath;
        public string TestResponsibleMan;
        public string DiscoveryUser;
        public string FunctionMenu;
        public string DevResponsibleMan;
        public string Source;
    }
//    select[System.Id], [Teld.Scrum.KeyApplication], [Teld.Scrum.ModulesName], [System.Title], 
//[System.AssignedTo], [System.State], [Teld.Bug.Type], [Microsoft.VSTS.Common.Severity], [Teld.Bug.Envir], 
//[System.CreatedDate], [System.ChangedDate], [Teld.Bug.DetectionMode], [Teld.Bug.DetectionPhase], [Teld.Bug.HopeFixSubmitTime], 
//[System.TeamProject], [System.CreatedBy], [System.IterationPath], [Teld.Scrum.TestResponsibleMan], [Teld.Bug.DiscoveryUser], 
//[Teld.Bug.FunctionMenu], [Teld.Scrum.DevResponsibleMan], [Teld.Bug.Source]
//    from WorkItems where[System.WorkItemType] = 'Bug' 
//        and[Microsoft.VSTS.Common.StateChangeDate] < '2017-11-18T00:00:00.0000000' 
//        and[Microsoft.VSTS.Common.StateChangeDate] > '2017-10-30T00:00:00.0000000' 
//        and([System.TeamProject] = 'TTP' or ([System.TeamProject] = 'Bugs' and[Teld.Scrum.BelongTeamProject] = 'TTP')) 
//    and([System.State] = '已关闭' or[System.State] = '已修复') and[Teld.Bug.Source] <> '预警引入' 
//        and([Teld.Bug.ResolvedReason] = '不是错误' or[Teld.Bug.ResolvedReason] = '不予处理') 
//        order by[System.AssignedTo]
}
