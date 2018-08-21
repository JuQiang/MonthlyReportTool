using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MonthlyReportTool.API.TFS.TeamProject;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Drawing;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using MonthlyReportTool.API.TFS.WorkItem;

namespace MonthlyReportTool.API.Office.PowerPoint.Quality
{
    public class BugAlertSlide : PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public BugAlertSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
        {
            var dateinfo = Utility.GetBeginEndDate(yearmonth);
            var teambugs = TFS.WorkItem.Bug.GetAlertByDate(project.Name, dateinfo.Item1, dateinfo.Item2);

            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "四、预警工单分析";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 24;

            var subframe = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 110f, 12.0f, 6.0f).TextFrame;
            // Add title
            var subtitle = subframe.TextRange;
            subtitle.Text = "1、预警工单分析";
            subtitle.Font.NameFarEast = "微软雅黑";
            subtitle.Font.Bold = MsoTriState.msoTrue;
            subtitle.Font.Color.RGB = 0x00C07000;
            subtitle.Font.Size = 16;

            

            var subframe2 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 330f, 200.0f, 50.0f).TextFrame;
            // Add title
            var subtitle2 = subframe2.TextRange;
            subtitle2.Text = "分析说明：";
            subtitle2.Font.NameFarEast = "微软雅黑";
            subtitle2.Font.Bold = MsoTriState.msoTrue;
            subtitle2.Font.Color.RGB = 0x00C07000;
            subtitle2.Font.Size = 16;

            List<Tuple<string, int>> list = new List<Tuple<string, int>>();

            list.Clear();
            var list1 = teambugs.GroupBy(bug => Utility.GetPersonName(bug.Principal)).OrderByDescending(bug => bug.Count());
            foreach (var bug in list1)
            {
                list.Add(Tuple.Create<string,int>(bug.Key, bug.Count()));
            }
            DrawChart("责任人", list, 60,150);

            list.Clear();
            var list2 = teambugs.GroupBy(bug => bug.WarningGrade).OrderBy(bug => bug.Key);
            foreach (var bug in list2)
            {
                list.Add(Tuple.Create<string, int>(bug.Key, bug.Count()));
            }
            DrawChart("预警级别", list, 260, 150);

            list.Clear();
            var list3 = teambugs.GroupBy(bug => bug.State).OrderBy(bug => bug.Key);
            foreach (var bug in list3)
            {
                list.Add(Tuple.Create<string, int>(bug.Key, bug.Count()));
            }
            DrawChart("状态", list, 460, 150);

        }

        private void DrawChart(string colname, List<Tuple<string,int>> allbugs,int left, int top)
        {
            #region copy
            var chartShape = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered, Left: left, Top: top,Width:200.0f,Height:100.0f);
            var chart = chartShape.Chart;
            chart.ShowDataLabelsOverMaximum = true;
            //input data
            chart.ChartData.Activate();

            ExcelInterop.Workbook workbook = chart.ChartData.Workbook;
            ExcelInterop.Worksheet sheet = workbook.Worksheets["Sheet1"];
            sheet.Cells.Clear();

            

            int num = allbugs.Count();
            var data = new object[num+1, 2];

            data[0, 0] = colname;
            data[0, 1] = "工单个数";

            int row = 1;
            foreach (var bug in allbugs)
            {
                data[row, 0] = bug.Item1;
                data[row, 1] = bug.Item2;
                row++;
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[num + 1, 2]];
            range.Value = data;

            //chart.SetSourceData("'Sheet1'!$A$2:$B$" + (num + 1));
            chart.SetSourceData("'Sheet1'!" + range.Address, PowerPointInterop.XlRowCol.xlColumns);
            chart.HasTitle = true;
            chart.ApplyDataLabels(PowerPointInterop.XlDataLabelsType.xlDataLabelsShowValue);
            chart.ChartTitle.Text = "按"+colname+"统计";

            chart.Refresh();
            workbook.Close();

            GC.Collect();

            #endregion copy
        }
    }
}
