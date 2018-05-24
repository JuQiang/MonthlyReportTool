using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class FeatureEntity
    {
        public int Id;
        public string KeyApplicationName;
        public string ModulesName;
        public string FuncName;
        public string Title;
        public string AssignedTo;
        public string MonthState;        
        public string State;
        public string RequireFinishedDate;
        public string DesignFinishedDate;
        public string DevelopmentFinishedDate;
        public string TestFinishedDate;
        public string AcceptFinishedDate;
        public string ReleaseFinishedDate;
        public string TeamProject;
        public string InitTargetDate;
        public string TargetDate;
        public string IterationTargetDate;
        public string NeedRequireDevelop;
        public string ParentId;
        //public string IsDevelopment;
    }

    //select [System.Id], [System.Title], [System.AssignedTo], [Teld.Scrum.MonthState], [System.State],
    //[Teld.Scrum.DevelopmentFinishedDate], [Teld.Scrum.ReleaseFinishedDate], [System.TeamProject], 
    //[Teld.Scrum.Scheduling.InitTargetDate], [Microsoft.VSTS.Scheduling.TargetDate], [Teld.Scrum.IsDevelopment] from WorkItems 
    //    where[System.WorkItemType] = '产品特性' 
    //    and[Microsoft.VSTS.Scheduling.TargetDate] >= '2017-10-01T00:00:00.0000000' 
    //    and[Microsoft.VSTS.Scheduling.TargetDate] < '2017-11-01T00:00:00.0000000' 
    //    and[System.State] <> '已移除' and[System.State] <> '已废除' 
    //    and[System.TeamProject] <> 'Car-TSL' 
    //    and[System.TeamProject] <> 'CA' 
    //    and[System.TeamProject] <> 'PPQA' 
    //    and[System.TeamProject] <> 'HCI' 
    //    and[System.TeamProject] <> 'SOM' 
    //    and[System.TeamProject] <> 'EM' 
    //    and[System.TeamProject] <> 'MMS' 
    //    and[System.TeamProject] <> 'PSS' 
    //    order by[System.TeamProject]]}
}
