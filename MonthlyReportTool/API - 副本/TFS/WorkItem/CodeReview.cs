using MonthlyReportTool.API.TFS.Agile;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class CodeReview
    {
        public static List<CodeReviewEntity> GetAll(string prj, IterationEntity ite)
        {
            string wiql = API.TFS.Utility.GetQueryClause("共享查询%2F迭代总结数据查询%2F20%20代码审查分析%2F00本迭代_代码审查审查的Bug总数");
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();
            String start0 = DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff");
            String endAdd1 = DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff");

            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = prj, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" }); //第一天是大于等于
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.CreatedDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" }); //最后一天要加一
            
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql, listquery);

            return GetBugsToEntityBySql(sql);
        }

        private static List<CodeReviewEntity> GetBugsToEntityBySql(string sql) {
            List<CodeReviewEntity> list = new List<CodeReviewEntity>();

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var entitys = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var entity in entitys)
            {
                list.Add(
                    new CodeReviewEntity()
                    {
                        Id = Convert.ToInt32(entity["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(entity["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(entity["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(entity["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(entity["fields"]["System.Title"]),
                        DetectionMode = Convert.ToString(entity["fields"]["Teld.Bug.DetectionMode"]),
                        AssignedTo = Convert.ToString(entity["fields"]["System.AssignedTo"]),
                        CreatedDate = Convert.ToString(entity["fields"]["System.CreatedDate"]),
                        CreatedDate2 = DateTime.Parse(Convert.ToString(entity["fields"]["System.CreatedDate"])).ToString("yyyy-MM-dd"),
                        CreatedBy = Convert.ToString(entity["fields"]["System.CreatedBy"]),
                        TeamProject = Convert.ToString(entity["fields"]["System.TeamProject"]),
                        Tags = Convert.ToString(entity["fields"]["System.Tags"]),
                        IterationPath = Convert.ToString(entity["fields"]["System.IterationPath"]),
                        State = Convert.ToString(entity["fields"]["System.State"]),
                        DetectionPhase = Convert.ToString(entity["fields"]["Teld.Bug.DetectionPhase"]),
                        Source = Convert.ToString(entity["fields"]["Teld.Scrum.Source"]),
                        AreaPath = Convert.ToString(entity["fields"]["System.AreaPath"]),
                    }
                );
            }
            return list;
        }
    }
}
