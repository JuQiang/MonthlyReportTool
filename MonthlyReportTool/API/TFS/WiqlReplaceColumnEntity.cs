using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS
{
    /// <summary>
    /// 查询条件中，替换column和value的定义
    /// select [System.Id], [Teld.Scrum.KeyApplicationName], [Teld.Scrum.ModulesName], [Teld.Scrum.FuncName], [System.Title], [System.AssignedTo], [Teld.Scrum.WorkLog.SumHours], [Teld.Scrum.Worklog.OverTimes], [Teld.Scrum.WorkLog.SupperType], [Teld.Scrum.WorkLog.Type], [System.CreatedDate], [Teld.Scrum.Worklog.WorkDate], [Teld.Scrum.InPlaned], [System.TeamProject] 
    /// from WorkItems 
    /// where [System.WorkItemType] = '工作日志' and [System.TeamProject] = '{0}' and [System.State] <> '已废除' and [System.State] <> '已移除' and [Teld.Scrum.Worklog.WorkDate] < '2017-11-20T00:00:00.0000000' and [Teld.Scrum.Worklog.WorkDate] >= '2017-10-30T00:00:00.0000000' order by [Teld.Scrum.WorkLog.Type] desc
    /// </summary>
    public class WiqlReplaceColumnEntity
    {
        public string column;//带比较符号，比如[System.TeamProject] =这样子的
        public string replacevalue;//直接是要替换的具体的值，比如TTP
        public string notinclude;//不包括某些column+value的替换，多个的话，用两个@@间隔？目前还没多个的情况，比如，不替换System.TeamProject = 'OrgPortal'的替换，空不关注
        public string exectOder;//0：顺序，1：循环，执行的顺序，是顺序一个个的替换，还是循环替换这个column，比如，如果有多个地方出现Syste.TeamProject，是只替换碰到的第一个，还是循环所有的都替换
    }
}
