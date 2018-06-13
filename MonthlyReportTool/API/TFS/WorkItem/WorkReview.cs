using MonthlyReportTool.API.TFS.Agile;
using System;
using System.Collections;
using System.Collections.Generic;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class WorkReview
    {
        //共享查询%2F迭代总结数据查询%2F20%20代码审查分析%2F00本迭代_审查记录单总数

        private static List<WorkReviewEntity> GetWorkReviewListByIteration(string project, IterationEntity ite, string query, List<string> columns, List<string> values)
        {
            List<WorkReviewEntity> list = new List<WorkReviewEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplaceInformationFromWIQLByProject(wiql, columns);

            string sql = String.Format(wiql, values.ToArray());

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            Hashtable hs = new Hashtable();
            var workreviews = Utility.ConvertWorkitemQueryResult2Array(responseBody, ref hs);// ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var workreview in workreviews)
            {
                list.Add(
                    new WorkReviewEntity()
                    {
                        Id = Convert.ToInt32(workreview["fields"]["System.Id"]),
                        workItemType= Convert.ToString(workreview["fields"]["System.WorkItemType"]),
                        KeyApplicationName = Convert.ToString(workreview["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(workreview["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(workreview["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(workreview["fields"]["System.Title"]),
                        State = Convert.ToString(workreview["fields"]["System.State"]),
                        AssignedTo = Convert.ToString(workreview["fields"]["System.AssignedTo"]),
                        ParentId = "",
                        TeamProject = Convert.ToString(workreview["fields"]["System.TeamProject"]),

                        ReviewBillType = Convert.ToString(workreview["fields"]["Teld.Scrum.ReviewBillType"]),
                        ReviewResponsibleMan = Convert.ToString(workreview["fields"]["Teld.Scrum.ReviewResponsibleMan"]),
                        PlanSubmitDate = Convert.ToString(workreview["fields"]["Teld.Scrum.PlanSubmitDate"]),
                        CreatedDate = Convert.ToString(workreview["fields"]["System.CreatedDate"]),
                        ClosedDate = Convert.ToString(workreview["fields"]["Microsoft.VSTS.Common.ClosedDate"]),
                        IterationPath = Convert.ToString(workreview["fields"]["System.IterationPath"]),
                        FindedBugCount = Convert.ToInt32(workreview["fields"]["Teld.Scrum.FindedBugCount"]),

                        //bug的一些信息
                        Type = Convert.ToString(workreview["fields"]["Teld.Bug.Type"]),
                        Severity = Convert.ToString(workreview["fields"]["Microsoft.VSTS.Common.Severity"]),
                        DetectionMode = Convert.ToString(workreview["fields"]["Teld.Bug.DetectionMode"]),
                        DiscoveryUser = Convert.ToString(workreview["fields"]["Teld.Bug.DiscoveryUser"]),
                    }
                );
            }
            return list;
        }

        public static List<List<WorkReviewEntity>> GetAll(string project, IterationEntity ite)
        {
            List<List<WorkReviewEntity>> list = new List<List<WorkReviewEntity>>();

            var all = GetWorkReviewListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F20%20代码审查分析%2F00本迭代_审查记录单总数",
                    new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.BelongTeamProject] =", "[Microsoft.VSTS.Common.ClosedDate] >=", "[Microsoft.VSTS.Common.ClosedDate] <" },
                new List<string>(){
                    project,
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );

            var bugall = GetWorkReviewListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F20%20代码审查分析%2F01本迭代_审查记录单审查出Bug总数",
                new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.BelongTeamProject] =", "[Microsoft.VSTS.Common.ClosedDate] >=", "[Microsoft.VSTS.Common.ClosedDate] <" },
                new List<string>(){
                    project,
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );

            list.Add(all);
            list.Add(bugall);

            return list;
        }
    }
}
