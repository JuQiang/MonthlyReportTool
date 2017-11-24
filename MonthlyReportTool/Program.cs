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
            if (args.Length < 1)
            {
                ShowHelp();
                return;
            }

            string prjname = args[0];

            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList().Where(prj => prj.Name.ToLower() == prjname.ToLower());
            if (prjlist.Count() < 1)
            {                
                ShowHelp();
                Console.WriteLine("=======================");
                Console.WriteLine("!!! No such project !!!");
                Console.WriteLine("=======================");
                return;
            }

            var project = prjlist.ToList()[0];
            var ite = API.TFS.Utility.GetBestIteration(project.Name);
            if (null == ite)
            {                
                ShowHelp();
                Console.WriteLine("=============================================");
                Console.WriteLine("!!! No iterations defined in this project !!!");
                Console.WriteLine("=============================================");
                return;
            }

            Console.WriteLine("TELD (R) TFS Report Tool Version 1.0");
            Console.WriteLine("Written by JuQiang.");
            Console.WriteLine("");
            Console.WriteLine("正在生成 《" + project.Description + "》 迭代总结报告...");
            API.Office.Excel.Utility.BuildIterationReports(project);
        }

        private static void ShowHelp()
        {
            Console.WriteLine("TELD (R) TFS Report Tool Version 1.0");
            Console.WriteLine("Written by JuQiang.");
            Console.WriteLine("");
            Console.WriteLine("Usage  : TRT <ProjectName>");
            Console.WriteLine("Here're the projects you can use.");

            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList().OrderBy(prj => prj.Name);
            foreach (var prj in prjlist)
            {
                Console.WriteLine("\t" + prj.Name);
            }
        }
    }
}
