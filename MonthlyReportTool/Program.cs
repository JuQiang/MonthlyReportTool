using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool
{
    class Program
    {
        static void Main(string[] args)
        {
            //var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList();
            //foreach (var prj in prjlist)
            //{
            //    Console.WriteLine(prj.Name);
            //    Console.WriteLine(API.TFS.Agile.Iteration.GetProjectIterations(prj.Name));
            //    Console.WriteLine("======================");
            //}
            API.EXCEL.BuildIterationReports();

            //var features = API.TFS.Utils.GetAllFeaturesByIterations("TTP\\FYQ4\\Sprint35");

        }
    }
}
