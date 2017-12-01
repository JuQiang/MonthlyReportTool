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
    public class BugDeveloperSlide : PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public BugDeveloperSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
        {
            var dateinfo = Utility.GetBeginEndDate(yearmonth);
            var teambugs = TFS.WorkItem.Bug.GetAllByDate(project.Name, dateinfo.Item1, dateinfo.Item2);

            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "一、Bug数量及分布情况统计分析 - 总体";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 24;

            var subframe = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 110f, 12.0f, 6.0f).TextFrame;
            // Add title
            var subtitle = subframe.TextRange;
            subtitle.Text = "4、Bug按开发人员统计分析";
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

            DrawChart(teambugs);
            
        }

        private void DrawChart(List<BugEntity> allbugs)
        {
            #region copy
            var chartShape = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered, Left: 60, Top: 150, Width: 800.0f, Height: 200.0f);
            var chart = chartShape.Chart;
            chart.ShowDataLabelsOverMaximum = true;
            //input data
            chart.ChartData.Activate();

            ExcelInterop.Workbook workbook = chart.ChartData.Workbook;
            ExcelInterop.Worksheet sheet = chart.ChartData.Workbook.Worksheets["Sheet1"];
            sheet.Cells.Clear();

            ExcelInterop.Range range;
            object[] objHeaders = { "开发人员", "BUG个数" };
            range = sheet.get_Range("A1", "B1");
            range.Value = objHeaders;

            var buglist = allbugs.GroupBy(bug => Utility.GetPersonName(bug.AssignedTo)).OrderByDescending(bug => bug.Count());
            int num = buglist.Count();
            var data = new object[num, 2];
            int row = 0;
            foreach (var bug in buglist)
            {
                data[row, 0] = bug.Key;
                data[row, 1] = bug.Count();
                row++;
            }

            range = sheet.get_Range("A2", "B" + (num + 1));
            range.Value = data;
            //sheet.get_Range("B1").Value = title;
            chart.SetSourceData("'Sheet1'!$A$2:$B$" + (num + 1));
            chart.SeriesCollection(1).Name = "开发人员";
            chart.HasTitle = true;
            chart.ApplyDataLabels(PowerPointInterop.XlDataLabelsType.xlDataLabelsShowValue);
            chart.ChartTitle.Text = "BUG按开发人员分布情况";
            chart.Refresh();
            workbook.Close();

            #endregion copy
        }
    }
}
