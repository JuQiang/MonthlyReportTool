using MonthlyReportTool.API.TFS.Agile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Commitment
    {
        //共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F00本迭代_提交单总数

        private static List<CommitmentEntity> GetCommitmentListByIteration(string project, IterationEntity ite, string query, List<string> columns, List<string> values)
        {
            List<CommitmentEntity> list = new List<CommitmentEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplaceInformationFromWIQLByProject(wiql, columns);

            string sql = String.Format(wiql, values.ToArray());

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var commitments = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var commitment in commitments)
            {
                list.Add(
                    new CommitmentEntity()
                    {
                        Id = Convert.ToInt32(commitment["fields"]["System.Id"]),
                        KeyApplicationName = Convert.ToString(commitment["fields"]["Teld.Scrum.KeyApplicationName"]),
                        ModulesName = Convert.ToString(commitment["fields"]["Teld.Scrum.ModulesName"]),
                        FuncName = Convert.ToString(commitment["fields"]["Teld.Scrum.FuncName"]),
                        SubmitType = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.SubmitType"]),
                        Title = Convert.ToString(commitment["fields"]["System.Title"]),
                        State = Convert.ToString(commitment["fields"]["System.State"]),
                        AssignedTo = Convert.ToString(commitment["fields"]["System.AssignedTo"]),
                        TestResponsibleMan = Convert.ToString(commitment["fields"]["Teld.Scrum.TestResponsibleMan"]),
                        SubmitUser = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.SubmitUser"]),
                        BackNum = Convert.ToInt32(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.BackNum"]),
                        IsNeedPerformanceTest = (Convert.ToString(commitment["fields"]["System.AssignedTo"])) == "是",
                        TestFinishedTime = Convert.ToString(commitment["fields"]["Teld.Scrum.TestFinishedTime"]),
                        SubmitDate = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.SubmitDate"]),
                        PlanTestFinishedTime = Convert.ToString(commitment["fields"]["Teld.Scrum.Backlog.PlanTestFinishedTime"]),
                        AcceptTime = Convert.ToString(commitment["fields"]["Teld.Scrum.Backlog.AcceptTime"]),
                        CreatedDate = Convert.ToString(commitment["fields"]["System.CreatedDate"]),
                        BackType = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.BackType"]),
                        IterationPath = Convert.ToString(commitment["fields"]["System.IterationPath"]),
                        SubmitNumberOfTime = Convert.ToInt32(commitment["fields"]["Teld.Scrum.SubmitNumberOfTime"]),
                        TeamProject = Convert.ToString(commitment["fields"]["System.TeamProject"]),
                        FindedBugCount = Convert.ToInt32(commitment["fields"]["Teld.Scrum.FindedBugCount"]),
                    }
                );
            }
            return list;
        }

        public static List<List<CommitmentEntity>> GetAll(string project, IterationEntity ite)
        {
            List<List<CommitmentEntity>> list = new List<List<CommitmentEntity>>();

            var all = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F00本迭代_提交单总数",
                new List<string>() { "[System.TeamProject] =","[Teld.Scrum.RemovedDate] >=", "[Teld.Scrum.RemovedDate] <", "[Teld.Scrum.TestFinishedTime] >=", "[Teld.Scrum.TestFinishedTime] <", "[System.CreatedDate] <" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );
            var testpassed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F05本迭代_提交单测试通过总数",
                    new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.TestFinishedTime] >=", "[Teld.Scrum.TestFinishedTime] <" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );

            var removed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F10本迭代_已移除提交单总数",
                new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.RemovedDate] >=", "[Teld.Scrum.RemovedDate] <"},
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );
            var needperf = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F15本迭代_需性能测试提交单总数",
                    new List<string>() { "[System.TeamProject] =", "[System.CreatedDate] <", "[Teld.Scrum.TestFinishedTime] >=" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                }
            );
            var perftestpassed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F20本迭代_性能测试通过提交单总数",
                new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.TestFinishedTime] >=", "[Teld.Scrum.TestFinishedTime] <" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );

            var failed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F30本迭代_打回过提交单总数",
                new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.TestFinishedTime] >=", "[Teld.Scrum.TestFinishedTime] <" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );
            var longtime = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F30本迭代_提交单持续时间过长总数",
                new List<string>() { "[System.TeamProject] =", "[Teld.Scrum.TestFinishedTime] >=", "[Teld.Scrum.TestFinishedTime] <" },
                new List<string>(){
                    project,
                    DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff"),//第一天是大于等于
                    DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff"),//最后一天要加一
                }
            );

            list.Add(all);
            list.Add(testpassed);
            list.Add(removed);
            list.Add(needperf);
            list.Add(perftestpassed);
            list.Add(failed);
            list.Add(longtime);

            return list;
        }
    }
}
