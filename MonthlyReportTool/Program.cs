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
            string user, pass, type, proj, path;

            ShowVersion();

            try
            {
                if (!IsValidArgument(args, out user, out pass, out type, out proj, out path)) return;

                API.TFS.Utility.User = user;
                API.TFS.Utility.Pass = pass;


                if (type == "iteration")
                {
                    GenerateIterationReport(proj, path);
                }
                else if (type == "quality")
                {
                    GenerateMonthQualityReport(proj, path);
                }
                else if (type == "month")
                {
                    GenerateMonthReport(path);
                }
                else
                {
                    ShowHelp();
                }
            }
            catch (Exception exception)
            {
                if (exception.TargetSite.Name == "EnsureSuccessStatusCode")
                {
                    API.Office.Excel.Utility.WriteLog("请确保网络通畅，或者TFS的用户名、密码、项目名称都是正确的。");
                }
                StringBuilder sb = new StringBuilder("！！！出错了！！！");
                for (Exception ex = exception; ex != null; ex = ex.InnerException)
                {
                    sb.AppendLine(ex.Message);
                    sb.AppendLine(ex.StackTrace);
                }
                API.Office.Excel.Utility.WriteLog(sb.ToString());
            }


        }

        private static void GenerateMonthReport(string path)
        {
            var today = DateTime.Now;
            var lastMonth = today.AddMonths(-1);

            string fname = string.Format("{0}\\云平台月度经营报告（{1}）.pptx", new object[]
                    {
                            path,
                            lastMonth.ToString("yyyy年MM月")
                    });
            try
            {
                File.Delete(fname);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.Name == "WinIOError")
                {
                    Console.WriteLine("=============================================");
                    Console.WriteLine("!!! 文件 《" + fname + "》 正在被使用 ，请关闭PowerPoint再重新运行本程序！！！");
                    Console.WriteLine("=============================================");
                    return;
                }
            }

            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("！！！注意！！！程序运行时，不要使用剪贴板！！！");
            API.Office.Excel.Utility.WriteLog("！！！一分钟左右就能出来结果，别着急！！！");
            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog(string.Format("正在生成 云平台月度经营报告（{0}）.pptx", lastMonth.ToString("yyyy年MM月")));

            API.Office.PowerPoint.Utility.BuildMonthReport(lastMonth.Year, lastMonth.Month);
        }

        private static void GenerateMonthQualityReport(string proj, string path)
        {
            var project = API.TFS.TeamProject.Project.RetrieveProjectList().Where(prj => prj.Name.ToLower() == proj.ToLower()).ToList()[0];

            string fname = string.Format("{0}\\{1}总结.pptx", new object[]
                    {
                            path,
                            project.Description,
                    });
            try
            {
                File.Delete(fname);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.Name == "WinIOError")
                {
                    Console.WriteLine("=============================================");
                    Console.WriteLine("!!! 文件 《" + fname + "》 正在被使用 ，请关闭PowerPoint再重新运行本程序！！！");
                    Console.WriteLine("=============================================");
                    return;
                }
            }

            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("！！！注意！！！程序运行时，不要使用剪贴板！！！");
            API.Office.Excel.Utility.WriteLog("！！！一分钟左右就能出来结果，别着急！！！");
            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("正在生成 《" + project.Description + "》 质量分析报告...");

            API.Office.PowerPoint.Utility.BuildQualityReport(project);
        }

        private static void GenerateIterationReport(string proj, string path)
        {
            var ite = API.TFS.Utility.GetBestIteration(proj);
            if (null == ite)
            {
                Console.WriteLine("该项目没有定义迭代信息。");
                return;
            }

            string fname = string.Format("{0}\\{1}总结({2}_{3}).xlsx", new object[]
                    {
                            path,
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
                if (ex.TargetSite.Name == "WinIOError")
                {
                    Console.WriteLine("=============================================");
                    Console.WriteLine("!!! 文件 《" + fname + "》 正在被使用 ，请关闭Excel再重新运行本程序！！！");
                    Console.WriteLine("=============================================");
                    return;
                }
            }

            var project = API.TFS.TeamProject.Project.RetrieveProjectList().Where(prj => prj.Name.ToLower() == proj.ToLower()).ToList()[0];

            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("！！！注意！！！程序运行时，不要使用剪贴板！！！");
            API.Office.Excel.Utility.WriteLog("！！！一分钟左右就能出来结果，别着急！！！");
            API.Office.Excel.Utility.WriteLog("");
            API.Office.Excel.Utility.WriteLog("正在生成 《" + project.Description + "》 迭代总结报告...");

            API.Office.Excel.Utility.BuildIterationReport(project,path);

        }

        private static bool IsValidArgument(string[] args, out string user, out string pass, out string type, out string proj, out string path)
        {
            user = pass = type = proj = path = "";

            if (args.Length != 6 && args.Length != 8 && args.Length != 10)
            {
                goto INVALID;
            }

            for (int i = 0; i < args.Length; i += 2)
            {
                switch (args[i].ToLower().Trim())
                {
                    case "-user":
                        user = args[i + 1].ToLower().Trim();
                        continue;
                    case "-pass":
                        pass = args[i + 1].Trim();
                        continue;
                    case "-type":
                        type = args[i + 1].ToLower().Trim();
                        continue;
                    case "-proj":
                        proj = args[i + 1].ToLower().Trim();
                        continue;
                    case "-path":
                        path = args[i + 1].ToLower().Trim();                        
                        continue;

                    default:
                        goto INVALID;
                }
            }

            if (args.Length == 6 && String.IsNullOrEmpty(proj)) goto INVALID;
            if (args.Length == 6 && proj != "list") goto INVALID;

            API.TFS.Utility.User = user;
            API.TFS.Utility.Pass = pass;

            var prjlist = API.TFS.TeamProject.Project.RetrieveProjectList().OrderBy(prj => prj.Name);

            if ((false==String.IsNullOrEmpty(path)) && (false==Directory.Exists(path)))
            {
                Console.WriteLine("目录不存在："+path);
                return false;
            }
            if (String.IsNullOrEmpty(path))
            {
                path = AppDomain.CurrentDomain.BaseDirectory;
            }
            if (path.EndsWith("\\")) path = path.Substring(0, path.Length - 1);

            if (proj == "list")
            {
                ShowProjectList(prjlist);
                return false;
            }

            string prjname = proj;
            var matchedProjects = prjlist.Where(prj => prj.Name.ToLower() == prjname.ToLower());
            if (prjlist.Count() < 1)
            {
                ShowProjectList(prjlist);
                Console.WriteLine("项目名称不存在。");
                goto INVALID;
            }



            goto SUCCESS;

            INVALID:
            ShowHelp();
            return false;
            SUCCESS:
            return true;
        }

        private static void ShowProjectList(IOrderedEnumerable<API.TFS.TeamProject.ProjectEntity> prjlist)
        {
            Console.WriteLine("可用的项目列表");
            foreach (var prj in prjlist)
            {
                Console.WriteLine("\t|-" + prj.Name);
            }
        }

        private static void ShowVersion()
        {
            Console.WriteLine("TELD (R) TFS Report Tool Version 1.0");
            Console.WriteLine("Written by JuQiang.");
            Console.WriteLine();
        }
        private static void ShowHelp()
        {
            Console.WriteLine("");
            Console.WriteLine("用法 : TRT -user <TFS用户名称> -pass <TFS用户密码> [-type <iteration|quality|month>] -proj <项目名称|list> -path <目录名称>");
            Console.WriteLine("");
            Console.WriteLine("       -type");
            Console.WriteLine("       |-如果-proj是list，则本参数会被忽略，否则本参数必填。");
            Console.WriteLine("       |-iteration：迭代总结报告");
            Console.WriteLine("       |-quality  ：月度经营报告中的质量分析报告");
            Console.WriteLine("       |-month    ：月度经营报告");
            Console.WriteLine("");
            Console.WriteLine("       -proj");
            Console.WriteLine("       |-该参数或者是项目名称，或者是list。后者时，可以列出所有项目列表。");
            Console.WriteLine("       |-当-type是iteration或者quality时，会引用这个参数。");
            Console.WriteLine("       |-当-type是month时，会忽略这个参数。");
            Console.WriteLine("");
            Console.WriteLine("       -path");
            Console.WriteLine("       |-这是一个目录。当-path，则默认输出到程序当前目录下。");
            Console.WriteLine("举例：");
            Console.WriteLine("     trt -user juqiang -pass MyPassword -proj list，获得项目列表");
            Console.WriteLine("     trt -user juqiang -pass MyPassword -proj bdp -type iteration，生成BDP的迭代总结报告");
            Console.WriteLine("     trt -user juqiang -pass MyPassword -proj bdp -type quality，生成BDP的月度经营报告中的质量分析报告");
            Console.WriteLine("     trt -user juqiang -pass MyPassword -type month，生成整个平台的月度经营报告（不包含各二级部门的质量分析报告）");
            Console.WriteLine("     trt -user juqiang -pass MyPassword -proj bdp -type iteration -path c:\\temp，生成BDP的迭代总结报告到c:\\temp目录下");
            Console.WriteLine("");
        }
    }
}
