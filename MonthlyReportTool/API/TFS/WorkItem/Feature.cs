using MonthlyReportTool.API.TFS.TeamProject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Feature
    {
        public static List<List<FeatureEntity>> GetAll(string project)
        {
            List<List<FeatureEntity>> list = new List<List<FeatureEntity>>();
            
            string iterationPath = Utility.GetBestIteration(project).Path;
            var all = GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F00本迭代_产品特性总数",
                                                                
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));

            var completed = GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F05本迭代_已完成产品特性总数",
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));
            var removed = GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F20本迭代_已中止或已移除产品特性总数",
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Teld.Scrum.StateChangeDate] >",
                "[Teld.Scrum.StateChangeDate] <"));
            var delayed = GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F10本迭代_拖期产品特性总数",
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));
            var perfect = GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F25本迭代_按计划完成产品特性总数",
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));

            list.Add(all);
            list.Add(completed);
            list.Add(removed);
            list.Add(delayed);
            list.Add(perfect);

            return list;
        }
        
        private static List<FeatureEntity> GetFeatureListByIteration(string iterationPath, string query, Tuple<string,string,string> tuple)
        {
            List<FeatureEntity> list = new List<FeatureEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplaceProjectAndDateFromWIQL(wiql, tuple);

            string[] pathinfo = iterationPath.Split(new char[] { '\\' });
            string prj = pathinfo[0];
            var allIterations = API.TFS.Agile.Iteration.GetProjectIterations(prj);
            var matchedFirstIteration = allIterations.Where(ite => (ite.Path.ToLower() == iterationPath.ToLower())).FirstOrDefault();

            string sql = String.Format(wiql,
                prj,
                matchedFirstIteration.StartDate,
                DateTime.Parse(matchedFirstIteration.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff")//最后一天要加一
            );

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var features = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var feature in features)
            {
                list.Add(
                    new FeatureEntity()
                    {
                        Id = Convert.ToInt32(feature["fields"]["System.Id"]),
                        KeyApplication = Convert.ToString(feature["fields"]["Teld.Scrum.KeyApplication"]),
                        ModulesName = Convert.ToString(feature["fields"]["Teld.Scrum.ModulesName"]),
                        Title = Convert.ToString(feature["fields"]["System.Title"]),
                        AssignedTo = Convert.ToString(feature["fields"]["System.AssignedTo"]),
                        MonthState = Convert.ToString(feature["fields"]["Teld.Scrum.MonthState"]),
                        State = Convert.ToString(feature["fields"]["System.State"]),
                        RequireFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.RequireFinishedDate"]),
                        DesignFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.DesignFinishedDate"]),
                        TestFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.TestFinishedDate"]),
                        AcceptFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.AcceptFinishedDate"]),
                        FunctionMenu = Convert.ToString(feature["fields"]["Teld.Bug.FunctionMenu"]),
                        DevelopmentFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.DevelopmentFinishedDate"]),
                        ReleaseFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.ReleaseFinishedDate"]),
                        TeamProject = Convert.ToString(feature["fields"]["System.TeamProject"]),
                        InitTargetDate = Convert.ToString(feature["fields"]["Teld.Scrum.Scheduling.InitTargetDate"]),
                        TargetDate = Convert.ToString(feature["fields"]["Microsoft.VSTS.Scheduling.TargetDate"]),
                        IsDevelopment = Convert.ToString(feature["fields"]["Teld.Scrum.IsDevelopment"]) == "是",
                    }
                );
            }

            return list;
        }

    }
}
