using MonthlyReportTool.API.TFS.Agile;
using MonthlyReportTool.API.TFS.TeamProject;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Feature
    {
        public static List<List<FeatureEntity>> GetAll(string project, IterationEntity ite)
        {
            List<List<FeatureEntity>> list = new List<List<FeatureEntity>>();

            var all = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F00本迭代_产品特性总数（New）",

                Tuple.Create<string, string, string>("[Teld.Scrum.BelongTeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >=",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));

            var completed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F05本迭代_已完成产品特性总数（New）",
                Tuple.Create<string, string, string>("[Teld.Scrum.BelongTeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >=",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));
            var removed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F20本迭代_已中止或已移除产品特性总数（New）",
                Tuple.Create<string, string, string>("[Teld.Scrum.BelongTeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >=",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));
            var delayed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F10本迭代_未完成产品特性总数（New）",
                Tuple.Create<string, string, string>("[Teld.Scrum.BelongTeamProject] =",
               "[Microsoft.VSTS.Scheduling.TargetDate] >=",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));
            var perfect = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F25本迭代_按计划完成产品特性总数（New）",
                Tuple.Create<string, string, string>("[Teld.Scrum.BelongTeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >=",
                "[Microsoft.VSTS.Scheduling.TargetDate] <"));

            list.Add(all);
            list.Add(completed);
            list.Add(removed);
            list.Add(delayed);
            list.Add(perfect);

            return list;
        }

        private static List<FeatureEntity> GetFeatureListByIteration(string project, IterationEntity ite, string query, Tuple<string, string, string> tuple)
        {
            List<FeatureEntity> list = new List<FeatureEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplacePrjAndDateFromWIQL(wiql, tuple);

            //string[] pathinfo = iterationPath.Split(new char[] { '\\' });
            //string prj = pathinfo[0];
            //var allIterations = API.TFS.Agile.Iteration.GetProjectIterations(prj);
            //var matchedFirstIteration = allIterations.Where(ite => (ite.Path.ToLower() == iterationPath.ToLower())).FirstOrDefault();

            string sql = String.Format(wiql,
                project,
                DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff")//最后一天要加一
            );

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            Hashtable hs = new Hashtable();
            var features = Utility.ConvertWorkitemQueryResult2Array(responseBody, ref hs);
            foreach (var feature in features)
            {
                var featureEntity =
                    new FeatureEntity()
                    {
                        Id = Convert.ToInt32(feature["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(feature["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(feature["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(feature["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(feature["fields"]["System.Title"]),
                        AssignedTo = Convert.ToString(feature["fields"]["System.AssignedTo"]),
                        MonthState = Convert.ToString(feature["fields"]["Teld.Scrum.MonthState"]),
                        State = Convert.ToString(feature["fields"]["System.State"]),
                        RequireFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.RequireFinishedDate"]),
                        DesignFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.DesignFinishedDate"]),
                        TestFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.TestFinishedDate"]),
                        AcceptFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.AcceptFinishedDate"]),
                        DevelopmentFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.DevelopmentFinishedDate"]),
                        ReleaseFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.ReleaseFinishedDate"]),
                        TeamProject = Convert.ToString(feature["fields"]["System.TeamProject"]),
                        InitTargetDate = Convert.ToString(feature["fields"]["Teld.Scrum.Scheduling.InitTargetDate"]),
                        IterationTargetDate = Convert.ToString(feature["fields"]["Teld.Scrum.IterationTargetDate"]),
                        TargetDate = Convert.ToString(feature["fields"]["Microsoft.VSTS.Scheduling.TargetDate"]),
                        NeedRequireDevelop = Convert.ToString(feature["fields"]["Teld.Scrum.NeedRequireDevelop"])
                        //IsDevelopment = Convert.ToString(feature["fields"]["Teld.Scrum.IsDevelopment"]),
                    };
                featureEntity.ParentId = (hs.ContainsKey(Convert.ToString(featureEntity.Id)) ? Convert.ToString(hs[Convert.ToString(featureEntity.Id)]) : "");
                list.Add(featureEntity);
            }

            return list;
        }

    }
}
