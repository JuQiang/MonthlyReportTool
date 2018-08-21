using MonthlyReportTool.API.TFS.Agile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class Workload
    {
        
        //共享查询%2F迭代总结数据查询%2F10%20工作量统计%2F05本迭代_实际所有的工作日志工作量
        public static List<WorkloadEntity> GetAll(string project, IterationEntity ite)
        {            
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();

            String start0 = DateTime.Parse(ite.StartDate).AddDays(0).ToString("yyyy-MM-dd HH:mm:ss.fff");//第一天是大于等于
            String endAdd1 = DateTime.Parse(ite.EndDate).AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff");//最后一天要加一

            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "[System.TeamProject] = 'OrgPortal'", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.Worklog.WorkDate] >=", replacevalue = start0, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[Teld.Scrum.Worklog.WorkDate] <", replacevalue = endAdd1, notinclude = "", exectOder = "1" });
            
            string wiql = API.TFS.Utility.GetQueryClause("共享查询%2F迭代总结数据查询%2F10%20工作量统计%2F05本迭代_实际所有的工作日志工作量");
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql, listquery);
            
            return GetWorkloadsToEntityBySql(sql);
        }
        private static List<WorkloadEntity> GetWorkloadsToEntityBySql(string sql) {
            List<WorkloadEntity> list = new List<WorkloadEntity>();

            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var workloads = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            foreach (var workload in workloads)
            {
                list.Add(
                            new WorkloadEntity()
                            {
                                Id = Convert.ToInt32(workload["fields"]["System.Id"]),
                                Title = Convert.ToString(workload["fields"]["System.Title"]),
                                AssignedTo = Convert.ToString(workload["fields"]["System.AssignedTo"]),
                                SumHours = Convert.ToDouble(workload["fields"]["Teld.Scrum.WorkLog.SumHours"]),
                                OverTimes = Convert.ToDouble(workload["fields"]["Teld.Scrum.Worklog.OverTimes"]),
                                SupperType = Convert.ToString(workload["fields"]["Teld.Scrum.WorkLog.SupperType"]),
                                Type = Convert.ToString(workload["fields"]["Teld.Scrum.WorkLog.Type"]),
                                CreatedDate = Convert.ToString(workload["fields"]["System.CreatedDate"]),
                                WorkDate = Convert.ToString(workload["fields"]["Teld.Scrum.Worklog.WorkDate"]),
                                InPlaned = Convert.ToString(workload["fields"]["Teld.Scrum.InPlaned"]),
                                TeamProject = Convert.ToString(workload["fields"]["System.TeamProject"]),

                            }
                        );
            }

            return list;
        }
        public static Tuple<double,double,double> GetEstimated(string project, IterationEntity ite)
        {
            List<WorkloadEntity> list = new List<WorkloadEntity>();

            string wiql = API.TFS.Utility.GetQueryClause("共享查询%2F迭代总结数据查询%2F10%20工作量统计%2F00本迭代_任务评估工作量以及实际工作量");
            List<WiqlReplaceColumnEntity> listquery = new List<WiqlReplaceColumnEntity>();
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.TeamProject] =", replacevalue = project, notinclude = "", exectOder = "1" });
            listquery.Add(new WiqlReplaceColumnEntity() { column = "[System.IterationPath] =", replacevalue = ite.Path, notinclude = "", exectOder = "1" });
            string sql = API.TFS.Utility.ReplaceInformationFromWIQLByReplaceList(wiql,listquery);
            
            string responseBody = Utility.ExecuteQueryBySQL(sql);
            var workloads = Utility.ConvertWorkitemFlatQueryResult2Array(responseBody);
            
            double esthour = 0.0d;
            double acthour = 0.0d;
            double lefthour = 0.0d;
            foreach (var workload in workloads)
            {
                esthour += Convert.ToDouble(workload["fields"]["Teld.Scrum.WorkItem.Task.EstimateHours"]);
                acthour += Convert.ToDouble(workload["fields"]["Teld.Scrum.WorkItem.Task.ActualHours"]);
                lefthour += Convert.ToDouble(workload["fields"]["Microsoft.VSTS.Scheduling.RemainingWork"]);
            }

            Tuple<double, double,double> ret = Tuple.Create<double, double, double>(esthour, acthour, lefthour);
            return ret;
        }
    }
}
