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
            List<CodeReviewEntity> list = new List<CodeReviewEntity>();

            string wiql = API.TFS.Utility.GetQueryClause("共享查询%2F迭代总结数据查询%2F20%20代码审查分析%2F00本迭代_代码审查审查的Bug总数");
            var tuple = Tuple.Create<string, string, string,string>("[System.TeamProject] =",
                "[System.CreatedDate] >=",
                "[System.CreatedDate] <","");
            wiql = API.TFS.Utility.ReplacePrjAndDateFromWIQL(wiql, tuple);

            string sql = String.Format(wiql,
                prj,
                DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                prj
            );

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var bugs = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var bug in bugs)
            {
                list.Add(
                    new CodeReviewEntity()
                    {
                        Id = Convert.ToInt32(bug["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(bug["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(bug["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(bug["fields"]["Teld.Scrum.FuncName"]),
                        Title = Convert.ToString(bug["fields"]["System.Title"]),
                        DetectionMode = Convert.ToString(bug["fields"]["Teld.Bug.DetectionMode"]),
                        AssignedTo = Convert.ToString(bug["fields"]["System.AssignedTo"]),
                        CreatedDate = Convert.ToString(bug["fields"]["System.CreatedDate"]),
                        CreatedDate2 = DateTime.Parse(Convert.ToString(bug["fields"]["System.CreatedDate"])).ToString("yyyy-MM-dd"),
                        CreatedBy = Convert.ToString(bug["fields"]["System.CreatedBy"]),
                        TeamProject = Convert.ToString(bug["fields"]["System.TeamProject"]),
                        Tags = Convert.ToString(bug["fields"]["System.Tags"]),
                        IterationPath = Convert.ToString(bug["fields"]["System.IterationPath"]),
                        State = Convert.ToString(bug["fields"]["System.State"]),
                        DetectionPhase = Convert.ToString(bug["fields"]["Teld.Bug.DetectionPhase"]),                        
                        Source = Convert.ToString(bug["fields"]["Teld.Scrum.Source"]),
                        AreaPath= Convert.ToString(bug["fields"]["System.AreaPath"]),
                    }
                );
            }

            return list;
        }
    }
}
