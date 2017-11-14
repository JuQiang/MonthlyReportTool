using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Feature
    {
        private static string GetBestIteration(string project)
        {
            string ret = String.Empty;

            DateTime now = DateTime.Now;
            var itelist = API.TFS.Agile.Iteration.GetProjectIterations(project);

            for (int i = 1; i < itelist.Count; i++)
            {
                if (string.IsNullOrEmpty(itelist[i - 1].EndDate) || string.IsNullOrEmpty(itelist[i - 0].EndDate)) continue;
                DateTime last1 = DateTime.ParseExact(itelist[i - 1].EndDate, "yyyy/M/d h:mm:ss", null);
                DateTime last2 = DateTime.ParseExact(itelist[i - 0].EndDate, "yyyy/M/d h:mm:ss", null);
                if (now >= last1 && now <= last2)
                {
                    ret = itelist[i - 1].Path;
                    return ret;
                    //Console.WriteLine(itelist[i - 1]);
                    //var all = API.TFS.WorkItem.Feature.GetAllFeatureListByIteration(itelist[i - 1].Path).Count;
                    //var abandon = API.TFS.WorkItem.Feature.GetAbandonFeatureListByIteration(itelist[i - 1].Path).Count;
                    //var completed = API.TFS.WorkItem.Feature.GetCompletedFeatureListByIteration(itelist[i - 1].Path).Count;
                    //var delayed = API.TFS.WorkItem.Feature.GetDelayedFeatureListByIteration(itelist[i - 1].Path).Count;
                    //var perfect = API.TFS.WorkItem.Feature.GetPerfectFeatureListByIteration(itelist[i - 1].Path).Count;
                    //var removed = API.TFS.WorkItem.Feature.GetRemovedFeatureListByIteration(itelist[i - 1].Path).Count;

                    //Console.WriteLine("All={0},已中止={1},已完成={2},延迟={3},按计划完成={4},已移除={5}",
                    //    all, abandon, completed, delayed, perfect, removed
                    //    );
                }
            }

            return ret;
        }
        private static List<FeatureEntity> GetFeatureListByIteration(string iterationPath, string query)
        {
            List<FeatureEntity> list = new List<FeatureEntity>();
            string wiql = API.TFS.Utils.GetQueryClause(query);
            wiql = API.TFS.Utils.ReplaceProjectAndDateFromWIQL(wiql,
                Tuple.Create<string, string, string>("[System.TeamProject] =",
                "[Microsoft.VSTS.Scheduling.TargetDate] >",
                "[Microsoft.VSTS.Scheduling.TargetDate] <")
                );
            

            string[] pathinfo = iterationPath.Split(new char[] { '\\' });
            string prj = pathinfo[0];
            var allIterations = API.TFS.Agile.Iteration.GetProjectIterations(prj);
            var matchedFirstIteration = allIterations.Where(ite => (ite.Path.ToLower() == iterationPath.ToLower())).FirstOrDefault();

            string sql = String.Format(wiql,
                prj,
                matchedFirstIteration.StartDate,
                DateTime.Parse(matchedFirstIteration.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff")//最后一天要加一
            );

            string responseBody = Utils.ExecuteQueryBySQL(sql);
            var features =Utils.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var feature in features)
            {
                list.Add(
                    new FeatureEntity()
                    {
                        Id = Convert.ToInt32(feature["System.Id"]),
                        KeyApplication = Convert.ToString(feature["fields"]["Teld.Scrum.KeyApplication"]),
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

        public static List<FeatureEntity> GetAllFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_产品特性总数");
        }

        public static List<FeatureEntity> GetAbandonFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_已中止产品特性总数");
        }

        public static List<FeatureEntity> GetCompletedFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_已完成产品特性总数");
        }

        public static List<FeatureEntity> GetRemovedFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_已移除产品特性总数");
        }
        public static List<FeatureEntity> GetDelayedFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_拖期产品特性总数");
        }
        public static List<FeatureEntity> GetPerfectFeatureListByIteration(string iterationPath)
        {
            return GetFeatureListByIteration(iterationPath, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F本迭代_按计划完成产品特性总数");
        }
    }
}
