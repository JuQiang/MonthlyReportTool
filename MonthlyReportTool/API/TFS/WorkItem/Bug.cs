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
        private static List<BugEntity> GetBugListByIteration(string project, IterationEntity ite, string query, Tuple<string, string, string, string> tuple)
        {
            List<BugEntity> list = new List<BugEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplacePrjAndDateAndPrjFromWIQL(wiql, tuple);

            string sql = String.Format(wiql,
                project,
                DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
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
                        KeyApplication = Convert.ToString(bug["fields"]["Teld.Scrum.KeyApplication"]),
                        ModulesName = Convert.ToString(bug["fields"]["Teld.Scrum.ModulesName"]),
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
                        FunctionMenu = Convert.ToString(bug["fields"]["Teld.Bug.FunctionMenu"]),
                        DevResponsibleMan = Convert.ToString(bug["fields"]["Teld.Scrum.DevResponsibleMan"]),
                        Source = Convert.ToString(bug["fields"]["Teld.Scrum.Source"]),
                    }
                );
            }
            return list;
        }

        public static List<BugEntity> GetAddedBugsByIteration(string project, IterationEntity ite)
        {
            var added = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F00本迭代_新增Bug总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[System.CreatedDate] >=",
                "[System.CreatedDate] <",
                "[Teld.Scrum.BelongTeamProject] ="
                )
            );

            return added;
        }
        public static List<List<BugEntity>> GetAllByIteration(string project, IterationEntity ite)
        {
            List<List<BugEntity>> list = new List<List<BugEntity>>();

            var added = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F00本迭代_新增Bug总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[System.CreatedDate] >=",
                "[System.CreatedDate] <",
                "[Teld.Scrum.BelongTeamProject] ="
                )
            );

            var _fixed = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F05本迭代_已修复Bug总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Common.StateChangeDate] >=",
                "[Microsoft.VSTS.Common.StateChangeDate] <",
                "[Teld.Scrum.BelongTeamProject] ="
                )
            );

            var notfixed = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F10本迭代_遗留Bug总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "_FQQ_",
                "[Microsoft.VSTS.Common.StateChangeDate] <",
                "[Teld.Scrum.BelongTeamProject] ="
                )
            );
            var critical = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F25本迭代_新增1或2级Bug总数（程序错误类）",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
            "[System.CreatedDate] >=",
            "[System.CreatedDate] <",
            "[Teld.Scrum.BelongTeamProject] ="
                )
            );

            var ignore = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F30本迭代_不予处理或不是错误Bug总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Common.StateChangeDate] >=",
                "[Microsoft.VSTS.Common.StateChangeDate] <",
                "[Teld.Scrum.BelongTeamProject] ="
                )
            );

            var review = GetBugListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F25%20Bug统计分析%2F35本迭代_新增评审问题总数",
            Tuple.Create<string, string, string, string>("[System.TeamProject] =",
                "[System.CreatedDate] >=",
                "[System.CreatedDate] <",
                "_FQQ_"
                )
            );

            list.Add(added);
            list.Add(_fixed);
            list.Add(notfixed);
            list.Add(critical);
            list.Add(ignore);
            list.Add(review);


            return list;
        }

        public static List<BugEntity> GetAllByDate(string project, string startDate, string endDate)
        {
            List<BugEntity> list = new List<BugEntity>();
                                                          
            string wiql = API.TFS.Utility.GetQueryClause("共享查询%2F研发月度运营会议数据统计%2F产品质量分析报告%2FBUG总体数量及类别分布分析");
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
                        KeyApplication = Convert.ToString(bug["fields"]["Teld.Scrum.KeyApplication"]),
                        ModulesName = Convert.ToString(bug["fields"]["Teld.Scrum.ModulesName"]),
                        Title = Convert.ToString(bug["fields"]["System.Title"]),
                        AssignedTo = Convert.ToString(bug["fields"]["System.AssignedTo"]),
                        State = Convert.ToString(bug["fields"]["System.State"]),
                        Type = Convert.ToString(bug["fields"]["Teld.Bug.Type"]),
                        Severity = Convert.ToString(bug["fields"]["Microsoft.VSTS.Common.Severity"]),                       
                        CreatedDate = Convert.ToString(bug["fields"]["System.CreatedDate"]),                        
                        TeamProject = Convert.ToString(bug["fields"]["System.TeamProject"]),
                        BelongTeamProject = Convert.ToString(bug["fields"]["Teld.Scrum.BelongTeamProject"]),
                        AreaPath = Convert.ToString(bug["fields"]["System.AreaPath"]),
                    }
                );
            }
            return list;
        }
    }
}
