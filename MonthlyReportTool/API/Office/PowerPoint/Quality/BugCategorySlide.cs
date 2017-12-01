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
    public class BugCategorySlide : PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public BugCategorySlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
        {
            var dateinfo = Utility.GetLast6MonthBeginEndDate(yearmonth);
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
            subtitle.Text = "2、Bug项目库、维护库分布情况分析";
            subtitle.Font.NameFarEast = "微软雅黑";
            subtitle.Font.Bold = MsoTriState.msoTrue;
            subtitle.Font.Color.RGB = 0x00C07000;
            subtitle.Font.Size = 16;

            var bugsdate = teambugs.GroupBy(wi => DateTime.Parse(wi.CreatedDate).ToString("yyyy-MM"));
            List<string> dateseries = new List<string>();
            foreach (var date in bugsdate)
            {
                dateseries.Add(date.Key);
            }


            var bugstype = teambugs.GroupBy(wi => wi.Type);

            var bugtbl = slide.Shapes.AddTable(dateseries.Count + 2, 6);

            bugtbl.Table.Cell(1, 1).Merge(bugtbl.Table.Cell(2, 1));
            var month = bugtbl.Table.Cell(1, 1).Shape.TextFrame.TextRange;
            month.Text = "月份"; month.Font.Size = 9;

            int row = 3;
            int col = 2;

            bugtbl.Table.Cell(1, col).Merge(bugtbl.Table.Cell(1, col + 1));
            var tr1 = bugtbl.Table.Cell(1, col).Shape.TextFrame.TextRange;
            tr1.Text = "项目库"; tr1.Font.Size = 9;

            var tr2 = bugtbl.Table.Cell(2, col).Shape.TextFrame.TextRange;
            tr2.Text = "数量"; tr2.Font.Size = 9;
            var tr3 = bugtbl.Table.Cell(2, col + 1).Shape.TextFrame.TextRange;
            tr3.Text = "占比"; tr3.Font.Size = 9;

            bugtbl.Table.Cell(1, col + 2).Merge(bugtbl.Table.Cell(1, col + 3));
            var tr11 = bugtbl.Table.Cell(1, col + 2).Shape.TextFrame.TextRange;
            tr11.Text = "维护库"; tr1.Font.Size = 9;

            var tr21 = bugtbl.Table.Cell(2, col + 2).Shape.TextFrame.TextRange;
            tr21.Text = "数量"; tr21.Font.Size = 9;
            var tr31 = bugtbl.Table.Cell(2, col + 3).Shape.TextFrame.TextRange;
            tr31.Text = "占比"; tr31.Font.Size = 9;

            foreach (string ym in dateseries)
            {
                var dt = bugtbl.Table.Cell(row, 1).Shape.TextFrame.TextRange;
                dt.Text = ym;
                dt.Font.Size = 9;

                col = 2;

                int count = teambugs
                    .Where(bug => DateTime.Parse(bug.CreatedDate).ToString("yyyy-MM") == ym)
                    .Where(bug => bug.TeamProject.ToLower() != "bugs").Count();
                var data = bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange;
                data.Text = count.ToString();
                data.Font.Size = 9;
                data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;

                count = teambugs
                    .Where(bug => DateTime.Parse(bug.CreatedDate).ToString("yyyy-MM") == ym)
                    .Where(bug => bug.TeamProject.ToLower() == "bugs").Count();
                data = bugtbl.Table.Cell(row, col + 2).Shape.TextFrame.TextRange;
                data.Text = count.ToString();
                data.Font.Size = 9;
                data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;

                row++;
            }

            bugtbl.Table.Cell(1, 6).Merge(bugtbl.Table.Cell(2, 6));
            var sumr = bugtbl.Table.Cell(1, 6).Shape.TextFrame.TextRange;
            sumr.Text = "合计"; sumr.Font.Size = 9;

            for (row = 3; row < dateseries.Count + 3; row++)
            {
                int sum = 0;
                for (col = 2; col < 6; col += 2)
                {
                    sum += Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
                }
                var data = bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange;
                data.Text = Convert.ToString(sum);
                data.Font.Size = 9;
                data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;


                for (col = 2; col < 6; col += 2)
                {
                    var percent = bugtbl.Table.Cell(row, col + 1).Shape.TextFrame.TextRange;
                    var num = Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
                    percent.Text = Convert.ToString(num * 100 / sum) + "%";
                    percent.Font.Size = 9;
                    percent.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;
                }
            }

            bugtbl.Width = 400.0f;
            bugtbl.Height = 150.0f;

            var subframe2 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 330f, 200.0f, 50.0f).TextFrame;
            // Add title
            var subtitle2 = subframe2.TextRange;
            subtitle2.Text = "分析说明：";
            subtitle2.Font.NameFarEast = "微软雅黑";
            subtitle2.Font.Bold = MsoTriState.msoTrue;
            subtitle2.Font.Color.RGB = 0x00C07000;
            subtitle2.Font.Size = 16;

            #region 画柱状图

            DrawChart(dateseries, bugtbl.Table);
            DrawChart2(dateseries, bugtbl.Table);
            return;
            

            #endregion 画柱状图
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
