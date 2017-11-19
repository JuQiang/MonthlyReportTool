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

        private static List<CommitmentEntity> GetCommitmentListByIteration(string project, IterationEntity ite, string query)
        {
            List<CommitmentEntity> list = new List<CommitmentEntity>();
            string wiql = API.TFS.Utility.GetQueryClause(query);
            wiql = API.TFS.Utility.ReplaceProjectAndIterationFromWIQL(wiql);

            string sql = String.Format(wiql,
                project,
                ite.Path
            );

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var commitments = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var commitment in commitments)
            {
                list.Add(
                    new CommitmentEntity()
                    {
                        Id = Convert.ToInt32(commitment["fields"]["System.Id"]),
                        KeyApplication = Convert.ToString(commitment["fields"]["Teld.Scrum.KeyApplication"]),
                        ModulesName = Convert.ToString(commitment["fields"]["Teld.Scrum.ModulesName"]),
                        SubmitType = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.SubmitType"]),
                        Title = Convert.ToString(commitment["fields"]["System.Title"]),
                        State = Convert.ToString(commitment["fields"]["System.State"]),
                        SubmitUser = Convert.ToString(commitment["fields"]["Teld.Scrum.Worklog.SubmitLog.SubmitUser"]),
                        AssignedTo = Convert.ToString(commitment["fields"]["System.AssignedTo"]),
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
                        FunctionMenu = Convert.ToString(commitment["fields"]["Teld.Bug.FunctionMenu"]),
                        TeamProject = Convert.ToString(commitment["fields"]["System.TeamProject"]),
                    }
                );
            }
            return list;
        }

        public static List<List<CommitmentEntity>> GetAll(string project, IterationEntity ite)
        {
            List<List<CommitmentEntity>> list = new List<List<CommitmentEntity>>();

            var all = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F00本迭代_提交单总数");
            var testpassed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F05本迭代_提交单测试通过总数");
            var removed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F10本迭代_已移除提交单总数");
            var needperf = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F15本迭代_需性能测试提交单总数");
            var perftestpassed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F20本迭代_性能测试通过提交单总数");
            var failed = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F30本迭代_打回过提交单总数");
            var longtime = GetCommitmentListByIteration(project, ite, "共享查询%2F迭代总结数据查询%2F15%20提交单统计分析%2F30本迭代_提交单持续时间过长总数");

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
