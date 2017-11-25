using MonthlyReportTool.API.TFS.WorkItem;
using System;
using System.Collections.Generic;
using System.IO;
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
                Console.ReadLine();
                return;
            }

            string prjname = args[0];

            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList().Where(prj => prj.Name.ToLower() == prjname.ToLower());
            if (prjlist.Count() < 1)
            {
                ShowHelp();
                Console.WriteLine("=======================");
                Console.WriteLine("!!! 没这个项目 !!!");
                Console.WriteLine("=======================");

                Console.ReadLine();
                return;
            }

            var project = prjlist.ToList()[0];
            var ite = API.TFS.Utility.GetBestIteration(project.Name);
            if (null == ite)
            {
                ShowHelp();
                Console.WriteLine("=============================================");
                Console.WriteLine("!!! 项目里没有定义迭代信息 !!!");
                Console.WriteLine("=============================================");

                Console.ReadLine();
                return;
            }

            string fname = string.Format("{0}\\{1}总结({2}_{3}).xlsx", new object[]
                        {
                            Environment.GetEnvironmentVariable("temp"),
                            ite.Path.Replace("\\", " "),
                            DateTime.Parse(ite.StartDate).ToString("yyyyMMdd"),
                            DateTime.Parse(ite.EndDate).ToString("yyyyMMdd")
                        });
            try
            {
                File.Delete(fname);
            }
            catch (Exception ex)
            {
                if(ex.TargetSite.Name == "WinIOError")
                {
                    Console.WriteLine("=============================================");
                    Console.WriteLine("!!! 文件 《" + fname + "》 正在被使用 ，请关闭Excel再重新运行本程序！！！");
                    Console.WriteLine("=============================================");
                    return;
                }
            }
            Console.WriteLine("TELD (R) TFS Report Tool Version 1.0");
            Console.WriteLine("Written by JuQiang.");
            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("！！！注意！！！程序运行时，不要使用剪贴板！！！");
            API.Office.Excel.Utility.WriteLog("！！！一分钟左右就能出来结果，别着急！！！");
            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("正在生成 《" + project.Description + "》 迭代总结报告...");

            try
            {
                API.Office.Excel.Utility.BuildIterationReports(project);
            }
            catch (Exception exception)
            {
                StringBuilder sb = new StringBuilder("！！！出错了！！！");
                for (Exception ex = exception; ex != null; ex = ex.InnerException)
                {
                    sb.AppendLine(ex.Message);
                    sb.AppendLine(ex.StackTrace);
                }
                API.Office.Excel.Utility.WriteLog(sb.ToString());
            }
        }

        private static void ShowHelp()
        {
            Console.WriteLine("TELD (R) TFS Report Tool Version 1.0");
            Console.WriteLine("Written by JuQiang.");
            Console.WriteLine("");
            Console.WriteLine("用法  : TRT <项目名称>");
            Console.WriteLine("目前有这些项目可以用：");

            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList().OrderBy(prj => prj.Name);
            foreach (var prj in prjlist)
            {
                Console.WriteLine("\t" + prj.Name);
            }

            Console.WriteLine("");
            Console.WriteLine("用法：TRT <项目名称>");
            Console.WriteLine("");
            Console.WriteLine("说明：打开命令行窗口，切换到该文件夹目录下，执行 TRT <项目名称>。项目名称见上，比如ttp或者bdp等，不区分大小写。");
            Console.WriteLine("按回车退出。");
        }
    }
}
