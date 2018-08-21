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

            String start0 = DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff");
            String endAdd1 = DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff");
           
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var all = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F00本迭代_产品特性总数（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.ReleaseFinishedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var completed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F05本迭代_已完成产品特性总数（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var removed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F20本迭代_已中止或已移除产品特性总数（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.ReleaseFinishedDate] >=", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var delayed = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F10本迭代_未完成产品特性总数（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var perfect = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F25本迭代_按计划完成产品特性总数（New）", listquery);
            //明细
            var alldetail = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F00本迭代_产品特性总数＿明细（New）", listquery);
            
            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.ReleaseFinishedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var completeddetail = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F05本迭代_已完成产品特性总数＿明细（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var removeddetail = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F20本迭代_已中止或已移除产品特性总数＿明细（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.ReleaseFinishedDate] >=", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var delayeddetail = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F10本迭代_未完成产品特性总数＿明细（New）", listquery);

            listquery.Clear();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.BelongTeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Microsoft.VSTS.Scheduling.TargetDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            var perfectdetail = GetFeatureListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F01%20产品特性统计分析%2F25本迭代_按计划完成产品特性总数＿明细（New）", listquery);
            
            list.Add(all);
            list.Add(completed);
            list.Add(removed);
            list.Add(delayed);
            list.Add(perfect);

            list.Add(alldetail);
            list.Add(completeddetail);
            list.Add(removeddetail);
            list.Add(delayeddetail);
            list.Add(perfectdetail);

            return list;
        }

        private static List<FeatureEntity> GetFeatureListByIteration(string project, IterationEntity ite, string query, List<WiqlReplaceColumnEntity> listquery)
        {
            string wiql = API.TFS.Utility.GetQueryClause(query);
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql, listquery);
            
            return GetFeaturesToEntityBySql(sql);            
        }
        private static List<FeatureEntity> GetFeaturesToEntityBySql(string sql) {
            List<FeatureEntity> list = new List<FeatureEntity>();
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
                        NeedRequireDevelop = Convert.ToString(feature["fields"]["Teld.Scrum.NeedRequireDevelop"]),
                        State = Convert.ToString(feature["fields"]["System.State"]),
                        PlanRequireFinishDate = Convert.ToString(feature["fields"]["Teld.Scrum.PlanRequireFinishedDate"]),
                        RequireFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.RequireFinishedDate"]),
                        ReleaseFinishedDate = Convert.ToString(feature["fields"]["Teld.Scrum.ReleaseFinishedDate"]),
                        ClosedDate = Convert.ToString(feature["fields"]["Microsoft.VSTS.Common.ClosedDate"]),
                        TeamProject = Convert.ToString(feature["fields"]["System.TeamProject"]),
                        InitTargetDate = Convert.ToString(feature["fields"]["Teld.Scrum.Scheduling.InitTargetDate"]),
                        TargetDate = Convert.ToString(feature["fields"]["Microsoft.VSTS.Scheduling.TargetDate"]),
                    };
                featureEntity.ParentId = (hs.ContainsKey(Convert.ToString(featureEntity.Id)) ? Convert.ToString(hs[Convert.ToString(featureEntity.Id)]) : "");
                list.Add(featureEntity);
            }

            return list;
        }

    }
}
