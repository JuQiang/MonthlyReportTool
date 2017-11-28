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

            var chartShape = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered,Left:500,Top:150,Width:300.0f,Height:150.0f);
            var chart = chartShape.Chart;
            chart.ShowDataLabelsOverMaximum = true;
            
            var chartdata = chart.ChartData;
            var wb = chartdata.Workbook;

            var ws = wb.Worksheets[1];

            ws.ListObjects("Table1").Resize(ws.Range("A1:C7"));
            ws.Range("Table1[[#Headers],[Series 1]]").Value = "项目库";
            ws.Range("Table1[[#Headers],[Series 2]]").Value = "维护库";
            for (int i = 0; i < dateseries.Count; i++)
            {
                ws.Cells(2 + i, 1).Value = dateseries[i].Substring(5,2);
                ws.Cells(2 + i, 2).Value = Convert.ToInt32(bugtbl.Table.Cell(3 + i, 2).Shape.TextFrame.TextRange.Text);
                ws.Cells(2 + i, 3).Value = Convert.ToInt32(bugtbl.Table.Cell(3 + i, 4).Shape.TextFrame.TextRange.Text);
            }

            var chartShape2 = slide.Shapes.AddChart2(Type: XlChartType.xlColumnClustered, Left: 500, Top: 400, Width: 300.0f, Height: 150.0f);
            var chart2 = chartShape2.Chart;
            chart2.ShowDataLabelsOverMaximum = true;

            var chartdata2 = chart2.ChartData;
            var wb2 = chartdata2.Workbook;

            var ws2 = wb2.Worksheets[2];

            ws2.ListObjects("Table1").Resize(ws.Range("A1:C7"));
            ws2.Range("Table1[[#Headers],[Series 1]]").Value = "项目库";
            ws2.Range("Table1[[#Headers],[Series 2]]").Value = "维护库";
            for (int i = 0; i < dateseries.Count; i++)
            {
                ws2.Cells(2 + i, 1).Value = dateseries[i].Substring(5, 2);
                ws2.Cells(2 + i, 2).Value = Convert.ToInt32(bugtbl.Table.Cell(3 + i, 3).Shape.TextFrame.TextRange.Text);
                ws2.Cells(2 + i, 3).Value = Convert.ToInt32(bugtbl.Table.Cell(3 + i, 5).Shape.TextFrame.TextRange.Text);
            }


            #endregion 画柱状图
        }
    }
}
