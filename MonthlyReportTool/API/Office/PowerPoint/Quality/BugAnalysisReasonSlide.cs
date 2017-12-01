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

namespace MonthlyReportTool.API.Office.PowerPoint.Quality
{
    public class BugAnalysisReasonSlide : PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public BugAnalysisReasonSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
        {
            var dateinfo = Utility.GetBeginEndDate(yearmonth);
            var teambugs = TFS.WorkItem.Bug.GetCriticalByDate(project.Name, dateinfo.Item1, dateinfo.Item2);

            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "三、Bug原因分析";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 24;

            var subframe = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 110f, 12.0f, 6.0f).TextFrame;
            // Add title
            var subtitle = subframe.TextRange;
            subtitle.Text = "2、程序错误Bug主要原因归纳分析（生产环境）";
            subtitle.Font.NameFarEast = "微软雅黑";
            subtitle.Font.Bold = MsoTriState.msoTrue;
            subtitle.Font.Color.RGB = 0x00C07000;
            subtitle.Font.Size = 16;

            

            var bugtbl = slide.Shapes.AddTable(teambugs.Count+1, 5);

            
            var cell1 = bugtbl.Table.Cell(1, 1).Shape.TextFrame.TextRange;
            cell1.Text = "ID"; cell1.Font.Size = 9;
            var cell2 = bugtbl.Table.Cell(1, 2).Shape.TextFrame.TextRange;
            cell2.Text = "问题模块"; cell2.Font.Size = 9;
            var cell3 = bugtbl.Table.Cell(1, 3).Shape.TextFrame.TextRange;
            cell3.Text = "标题"; cell3.Font.Size = 9;
            var cell4 = bugtbl.Table.Cell(1, 4).Shape.TextFrame.TextRange;
            cell4.Text = "严重级别"; cell4.Font.Size = 9;
            var cell5 = bugtbl.Table.Cell(1, 5).Shape.TextFrame.TextRange;
            cell5.Text = "BUG产生原因"; cell5.Font.Size = 9;

            for (int row = 2; row <= teambugs.Count + 1; row++)
            {
                var c1 = bugtbl.Table.Cell(row, 1).Shape.TextFrame.TextRange;
                c1.Text =teambugs[row-2].Id.ToString(); cell3.Font.Size = 9;
                var c2 = bugtbl.Table.Cell(row, 2).Shape.TextFrame.TextRange;
                c2.Text = teambugs[row - 2].ModulesName; cell3.Font.Size = 9;
                var c3 = bugtbl.Table.Cell(row, 3).Shape.TextFrame.TextRange;
                c3.Text = teambugs[row - 2].Title; cell3.Font.Size = 9;
                var c4 = bugtbl.Table.Cell(row, 4).Shape.TextFrame.TextRange;
                c4.Text = teambugs[row - 2].Severity; cell3.Font.Size = 9;
                var c5 = bugtbl.Table.Cell(row, 5).Shape.TextFrame.TextRange;
                c5.Text = ""; cell3.Font.Size = 9;
            }

            var subframe2 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 330f, 200.0f, 50.0f).TextFrame;
            // Add title
            var subtitle2 = subframe2.TextRange;
            subtitle2.Text = "分析说明：";
            subtitle2.Font.NameFarEast = "微软雅黑";
            subtitle2.Font.Bold = MsoTriState.msoTrue;
            subtitle2.Font.Color.RGB = 0x00C07000;
            subtitle2.Font.Size = 16;
        }

        private void DrawChart(List<string> dateseries, PowerPointInterop.Table table)
        {
            #region copy
            var chartShape = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered, Left: 500, Top: 120, Width: 400.0f, Height: 200.0f);
            var chart = chartShape.Chart;
            chart.ShowDataLabelsOverMaximum = true;
            //input data
            chart.ChartData.Activate();
            
            ExcelInterop.Workbook workbook = chart.ChartData.Workbook;
            ExcelInterop.Worksheet sheet = chart.ChartData.Workbook.Worksheets["Sheet1"];
            sheet.Cells.Clear();

            ExcelInterop.Range range;
            object[] objHeaders = { "时间","项目库", "维护库" };
            range = sheet.get_Range("A1", "C1");
            range.Value = objHeaders;

            int num = dateseries.Count();
            var data = new object[num, 3];
            foreach (int n in Enumerable.Range(0, num))
            {
                data[n, 0] = dateseries[n];
                data[n, 1] = Convert.ToInt32(table.Cell(3 + n, 2).Shape.TextFrame.TextRange.Text);
                data[n, 2] = Convert.ToInt32(table.Cell(3 + n, 4).Shape.TextFrame.TextRange.Text);
            }

            range = sheet.get_Range("A2", "C" + (num + 1));
            range.Value = data;
            //sheet.get_Range("B1").Value = title;
            chart.SetSourceData("'Sheet1'!$A$2:$C$"+(num+1));
            chart.SeriesCollection(1).Name = "项目库";
            chart.SeriesCollection(2).Name = "维护库";
            chart.HasTitle = true;
            chart.ApplyDataLabels(PowerPointInterop.XlDataLabelsType.xlDataLabelsShowValue);
            chart.ChartTitle.Text = "项目库和维护库BUG个数分布";

            workbook.Close();

            #endregion copy
        }

        private void DrawChart2(List<string> dateseries, PowerPointInterop.Table table)
        {
            #region copy
            var chartShape = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered, Left: 500, Top: 320, Width: 400.0f, Height: 200.0f);
            var chart = chartShape.Chart;
            chart.ShowDataLabelsOverMaximum = true;
            //input data
            chart.ChartData.Activate();

            ExcelInterop.Workbook workbook = chart.ChartData.Workbook;
            ExcelInterop.Worksheet sheet = chart.ChartData.Workbook.Worksheets["Sheet1"];
            sheet.Cells.Clear();

            ExcelInterop.Range range;
            object[] objHeaders = { "时间", "项目库", "维护库" };
            range = sheet.get_Range("A1", "C1");
            range.Value = objHeaders;

            int num = dateseries.Count();
            var data = new object[num, 3];
            foreach (int n in Enumerable.Range(0, num))
            {
                data[n, 0] = dateseries[n];
                data[n, 1] = table.Cell(3 + n, 3).Shape.TextFrame.TextRange.Text;
                data[n, 2] = table.Cell(3 + n, 5).Shape.TextFrame.TextRange.Text;
            }

            range = sheet.get_Range("A2", "C" + (num + 1));
            range.Value = data;
            //sheet.get_Range("B1").Value = title;
            chart.SetSourceData("'Sheet1'!$A$2:$C$" + (num + 1));
            chart.SeriesCollection(1).Name = "项目库";
            chart.SeriesCollection(2).Name = "维护库";
            chart.HasTitle = true;
            chart.ApplyDataLabels(PowerPointInterop.XlDataLabelsType.xlDataLabelsShowValue);
            chart.ChartTitle.Text = "项目库和维护库BUG占比分析";

            workbook.Close();

            #endregion copy
        }
    }
}
