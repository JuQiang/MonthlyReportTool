﻿using MonthlyReportTool.API.TFS.Agile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Backlog
    {
        public static List<List<BacklogEntity>> GetAll(string project, IterationEntity ite)
        {
            List<List<BacklogEntity>> list = new List<List<BacklogEntity>>();

            list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F05本迭代_已完成积压工作项总数"));
            list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F10本迭代_未启动积压工作项总数"));
            list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F15本迭代_已中止或已移除积压工作项总数"));
            list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F00本迭代_积压工作项总数"));
            list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F12本迭代_拖期积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F30本迭代_已提交测试积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F35本迭代_测试通过积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F40本迭代_应提交测试积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F50本迭代_需编写用例积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F60本迭代_实际编写用例积压工作项总数"));
            //list.Add(GetBacklogListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F05%20Backlog统计分析%2F65本迭代_编写用例总数"));

            return list;
        }

        private static List<BacklogEntity> GetBacklogListByIteration(string project, IterationEntity ite, string query)
        {
            string wiql = API.TFS.Utility.GetQueryClause(query);
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.IterationPath] =", replacevalue = ite.Path, notinclude = "", exectOder = "1" });
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql, listquery);

            return GetBacklogsToEntityBySql(sql);
        }
        private static List<BacklogEntity> GetBacklogsToEntityBySql(string sql)
        {
            List<BacklogEntity> list = new List<BacklogEntity>();
            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var backlogs = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var backlog in backlogs)
            {
                list.Add(
                    new BacklogEntity()
                    {
                        Id = Convert.ToInt32(backlog["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(backlog["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(backlog["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(backlog["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(backlog["fields"]["System.Title"]),
                        Category = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.Category"]),
                        AssignedTo = Convert.ToString(backlog["fields"]["System.AssignedTo"]),
                        AcceptanceMeasure = Convert.ToString(backlog["fields"]["Teld.Scrum.AcceptanceMeasure"]),
                        State = Convert.ToString(backlog["fields"]["System.State"]),
                        HopeSubmitTime = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.HopeSubmitTime"]),
                        IsPlaned = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.IsPlaned"]),
                        CreatedDate = Convert.ToString(backlog["fields"]["System.CreatedDate"]),
                        Tags = Convert.ToString(backlog["fields"]["System.Tags"]),
                        AcceptTime = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.AcceptTime"]),
                        IsNeedInterfaceTest = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.IsNeedInterfaceTest"]),
                        IsNeedPerformanceTest = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.IsNeedPerformanceTest"]),
                        SubmitTime = Convert.ToString(backlog["fields"]["Teld.Scrum.Backlog.SubmitTime"]),
                        TeamProject = Convert.ToString(backlog["fields"]["System.TeamProject"]),
                        FinishDate = Convert.ToString(backlog["fields"]["Microsoft.VSTS.Scheduling.FinishDate"]),
                    }
                );
            }

            return list;
        }
    }
}
