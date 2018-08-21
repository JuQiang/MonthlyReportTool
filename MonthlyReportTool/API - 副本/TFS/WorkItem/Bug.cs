using MonthlyReportTool.API.TFS.Agile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Bug
    {
        private static List<BugEntity> GetBugListByIteration(string project, IterationEntity ite, string query, List<WiqlReplaceColumnEntity> listquery)
        {
            string wiql = API.TFS.Utility.GetQueryClause(query);
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql, listquery);

            return GetBugsToEntityBySql(sql);
        }
        //获取的数据,转换成实体类列表
        private static List<BugEntity> GetBugsToEntityBySql(String sql)
        {
            List<BugEntity> list = new List<BugEntity>();
            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var bugs = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var bug in bugs)
            {
                list.Add(
                    new BugEntity()
                    {
                        Id = Convert.ToInt32(bug["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(bug["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(bug["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(bug["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(bug["fields"]["System.Title"]),
                        AssignedTo = Convert.ToString(bug["fields"]["System.AssignedTo"]),
                        State = Convert.ToString(bug["fields"]["System.State"]),
                        Type = Convert.ToString(bug["fields"]["Teld.Bug.Type"]),
                        Severity = Convert.ToString(bug["fields"]["Microsoft.VSTS.Common.Severity"]),
                        ResolvedReason = Convert.ToString(bug["fields"]["Teld.Bug.ResolvedReason"]),
                        Envir = Convert.ToString(bug["fields"]["Teld.Bug.Envir"]),
                        CreatedDate = Convert.ToString(bug["fields"]["System.CreatedDate"]),
                        ChangedDate = Convert.ToString(bug["fields"]["System.ChangedDate"]),
                        DetectionMode = Convert.ToString(bug["fields"]["Teld.Bug.DetectionMode"]),
                        DetectionPhase = Convert.ToString(bug["fields"]["Teld.Bug.DetectionPhase"]),
                        HopeFixSubmitTime = Convert.ToString(bug["fields"]["Teld.Bug.HopeFixSubmitTime"]),
                        TeamProject = Convert.ToString(bug["fields"]["System.TeamProject"]),
                        CreatedBy = Convert.ToString(bug["fields"]["System.CreatedBy"]),
                        IterationPath = Convert.ToString(bug["fields"]["System.IterationPath"]),
                        TestResponsibleMan = Convert.ToString(bug["fields"]["Teld.Scrum.TestResponsibleMan"]),
                        DiscoveryUser = Convert.ToString(bug["fields"]["Teld.Bug.DiscoveryUser"]),
                        DevResponsibleMan = Convert.ToString(bug["fields"]["Teld.Scrum.DevResponsibleMan"]),
                        Source = Convert.ToString(bug["fields"]["Teld.Scrum.Source"]),
                    }
                );
            }
            return list;
        }
        public static List<BugEntity> GetAddedBugsByIteration(string project, IterationEntity ite)
        {
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();

            String start0 = DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff");//第一天是大于等于
            String endAdd1 = DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff");//最后一天要加一

            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            var added = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F00本迭代_新增Bug总数", listquery);

            return added;
        }
        public static List<List<BugEntity>> GetAllByIteration(string project, IterationEntity ite)
        {
            List<List<BugEntity>> list = new List<List<BugEntity>>();

            String start0 = DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff");
            String endAdd1 = DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff");

            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();
            var added = GetAddedBugsByIteration(project, ite);
            
            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Common.StateChangeDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Common.StateChangeDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            var _fixed = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F05本迭代_已修复Bug总数", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Common.StateChangeDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            var notfixed = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F10本迭代_遗留Bug总数", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            var critical = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F25本迭代_新增1或2级Bug总数（程序错误类）",listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Common.StateChangeDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Common.StateChangeDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); 
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            var ignore = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F30本迭代_不予处理或不是错误Bug总数",listquery);

            //var review = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F35本迭代_新增评审问题总数",
            //Tuple.Create<string, string, string, string>("[System.TeamProject] =",
            //    "[System.CreatedDate] >=",
            //    "[System.CreatedDate] <",
            //    "_FQQ_"
            //    )
            //);

            list.Add(added);
            list.Add(_fixed);
            list.Add(notfixed);
            list.Add(critical);
            list.Add(ignore);
            //list.Add(review);

            return list;
        }

        public static List<BugEntity> GetAllByDate(string project, string startDate, string endDate)
        {
            return GetByDate("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2F01BUG数量及分布情况统计分析%2FBUG数量及分布情况统计分析", project, startDate, endDate);
        }

        public static List<BugEntity> GetFixByDate(string project, string startDate, string endDate)
        {
            return GetByDate("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2F02Bug修复情况%2F未关闭BUG统计", project, startDate, endDate);
        }

        public static List<BugEntity> GetCriticalByDate(string project, string startDate, string endDate)
        {
            return GetByDate("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2F03Bug原因分析%2F维护库一二级BUG统计", project, startDate, endDate);
        }

        public static List<BugEntity> GetDevErrorByDate(string project, string startDate, string endDate)
        {
            return GetByDate("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2F03Bug原因分析%2F维护库程序错误类BUG统计", project, startDate, endDate);
        }

        public static List<BugEntity> GetAlertByDate(string project, string startDate, string endDate)
        {
            return GetByDate("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2F04预警工单分析%2F预警工单统计", project, startDate, endDate);
        }
        private static List<BugEntity> GetByDate(string query, string project, string startDate, string endDate)
        {
            List<BugEntity> list = new List<BugEntity>();

            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = wiql.Replace("[System.TeamProject] = @project and", "");//这里面要过滤掉？但是TFS UI上没有定义，我这里为啥又能搞出来呢？
            var tuple = Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[System.CreatedDate] >=",
                "[System.CreatedDate] <=",
                "[Teld.Scrum.BelongTeamProject] ="
                );

            wiql = API.TFS.Utility.ReplacePrjAndDateAndPrjFromWIQL(wiql, tuple);

            string sql = String.Format(wiql,
                project,
                startDate,
                endDate,
                project
            );

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var bugs = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var bug in bugs)
            {
                list.Add(
                    new BugEntity()
                    {
                        Id = Convert.ToInt32(bug["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(bug["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(bug["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(bug["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(bug["fields"]["System.Title"]),
                        AssignedTo = Convert.ToString(bug["fields"]["System.AssignedTo"]),
                        State = Convert.ToString(bug["fields"]["System.State"]),
                        Type = Convert.ToString(bug["fields"]["Teld.Bug.Type"]),
                        Severity = Convert.ToString(bug["fields"]["Microsoft.VSTS.Common.Severity"]),
                        CreatedDate = Convert.ToString(bug["fields"]["System.CreatedDate"]),
                        TeamProject = Convert.ToString(bug["fields"]["System.TeamProject"]),
                        BelongTeamProject = Convert.ToString(bug["fields"]["Teld.Scrum.BelongTeamProject"]),
                        AreaPath = Convert.ToString(bug["fields"]["System.AreaPath"]),
                        Principal = Convert.ToString(bug["fields"]["Teld.Scrum.Principal"]),
                        WarningGrade = Convert.ToString(bug["fields"]["Teld.Scrum.WarningGrade"]),
                    }
                );
            }
            return list;
        }
    }
}
