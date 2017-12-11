using MonthlyReportTool.API.TFS.TeamProject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using MonthlyReportTool.API.Office.PowerPoint.Quality;

namespace MonthlyReportTool.API.Office.PowerPoint
{
    public class Utility
    {
        public static void BuildQualityReport(ProjectEntity project, string yearmonth)
        {
            var pptApplication = new PowerPointInterop.Application();
            // Create the Presentation File
            var pptPresentation = pptApplication.Presentations.Add();// (MsoTriState.msoTrue);

            // Create new Slide
            var slides = pptPresentation.Slides;
            var customLayout = pptPresentation.SlideMaster.CustomLayouts[PowerPointInterop.PpSlideLayout.ppLayoutText];

            List<Tuple<string, Type>> allSlides = new List<Tuple<string, Type>>()
            {
                //Tuple.Create<string, Type>("说明",typeof(IntroductionSlide)),
                //Tuple.Create<string, Type>("目录",typeof(ContentSlide)),
                //Tuple.Create<string, Type>("一、Bug数量及分布情况统计分析",typeof(BugOverviewSlide)),
                //Tuple.Create<string, Type>("一、Bug数量及分布情况统计分析2",typeof(BugCategorySlide)),
                // Tuple.Create<string, Type>("一、Bug数量及分布情况统计分析3",typeof(BugModuleSlide)),
                // Tuple.Create<string, Type>("一、Bug数量及分布情况统计分析4",typeof(BugDeveloperSlide)),
                // Tuple.Create<string, Type>("一、Bug数量及分布情况统计分析5",typeof(BugSeveritySlide)),
                // Tuple.Create<string, Type>("二、Bug修复情况",typeof(BugFixSlide)),
                // Tuple.Create<string, Type>("二、Bug修复情况2",typeof(BugFixDeveloperSlide)),
                 Tuple.Create<string, Type>("三、Bug原因分析",typeof(BugAnalysisCriticalSlide)),
                // Tuple.Create<string, Type>("三、Bug原因分析2",typeof(BugAnalysisReasonSlide)),
                // Tuple.Create<string, Type>("四、预警工单分析",typeof(BugAlertSlide)),
            };

            for (int i = 0; i < allSlides.Count; i++)
            {
                PowerPointInterop.Slide slide = slides.AddSlide(i + 1, customLayout);
                slide.Name = allSlides[i].Item1;

                Console.WriteLine("\t" + slide.Name);

                Type t = allSlides[i].Item2;
                var ci = t.GetConstructor(new Type[] { typeof(PowerPointInterop.Slide) });
                object obj = ci.Invoke(new object[] { slide });
                t.InvokeMember("Build", BindingFlags.InvokeMethod, null, obj, new object[] { project, yearmonth });
            }

        }

        public static string GetPersonName(string fullname)
        {
            if (fullname.Trim().Length < 1) return "";
            return fullname.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

        }
        public static void BuildMonthReport(int year, int month)
        {

        }

        public static string[] GetDateSeriesByFriendlyFormat(List<string> dateseries)
        {
            string[] array = new string[dateseries.Count];
            //显示去冗余用
            dateseries.CopyTo(array);
            bool isSameYear = true;
            string year = array[0].Substring(0, 4);
            for (int i = 1; i < array.Length; i++)
            {
                if (year != array[1].Substring(0, 4))
                {
                    isSameYear = false;
                    break;
                }
            }

            if (isSameYear)
            {
                for (int i = 0; i < array.Length; i++)
                {
                    array[i] = array[i].Substring(5);
                }
            }

            return array;
        }

        public static Tuple<string, string> GetBeginEndDate(string yearmonth)
        {
            int year = Convert.ToInt32(yearmonth.Substring(0, 4));
            int month = Convert.ToInt32(yearmonth.Substring(4, 2));
            DateTime dt = new DateTime(year, month, 1);
            DateTime dt2 = dt.AddMonths(1);
            return Tuple.Create<string, string>(dt.ToString("yyyy-MM-dd 00:00:00.000"), dt2.ToString("yyyy-MM-dd 00:00:00.000"));
        }

        public static Tuple<string, string> GetLast6MonthBeginEndDate(string yearmonth)
        {
            int year = Convert.ToInt32(yearmonth.Substring(0, 4));
            int month = Convert.ToInt32(yearmonth.Substring(4, 2));
            DateTime dt = new DateTime(year, month, 1);
            DateTime dt2 = dt.AddMonths(1);
            dt = dt2.AddMonths(-6);//最近6个月
            return Tuple.Create<string, string>(dt.ToString("yyyy-MM-dd 00:00:00.000"), dt2.ToString("yyyy-MM-dd 00:00:00.000"));
        }
    }
}
