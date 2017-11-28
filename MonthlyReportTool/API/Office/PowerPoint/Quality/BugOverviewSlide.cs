using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MonthlyReportTool.API.TFS.TeamProject;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.PowerPoint.Quality
{
    public class BugOverviewSlide :  PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public BugOverviewSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project,string yearmonth)
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

            var subframe = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 110f, 20.0f, 5.0f).TextFrame;
            // Add title
            var subtitle = subframe.TextRange;
            subtitle.Text = "1、Bug数量及类别分布情况分析";
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

            var bugtbl = slide.Shapes.AddTable(dateseries.Count + 2, bugstype.Count() * 2 + 2);

            bugtbl.Table.Cell(1, 1).Merge(bugtbl.Table.Cell(2, 1));
            var month = bugtbl.Table.Cell(1, 1).Shape.TextFrame.TextRange;
            month.Text = "月份"; month.Font.Size = 9;

            int row = 3;
            foreach (string ym in Utility.GetDateSeriesByFriendlyFormat(dateseries))
            {
                var dt = bugtbl.Table.Cell(row++, 1).Shape.TextFrame.TextRange;
                dt.Text = ym;
                dt.Font.Size = 9;
            }

            int col = 2;
            foreach (var bugtype in bugstype)
            {
                bugtbl.Table.Cell(1, col).Merge(bugtbl.Table.Cell(1, col + 1));
                var tr1 = bugtbl.Table.Cell(1, col).Shape.TextFrame.TextRange;
                tr1.Text = bugtype.Key; tr1.Font.Size = 9;

                var tr2 = bugtbl.Table.Cell(2, col).Shape.TextFrame.TextRange;
                tr2.Text = "数量"; tr2.Font.Size = 9;
                var tr3 = bugtbl.Table.Cell(2, col + 1).Shape.TextFrame.TextRange;
                tr3.Text = "占比"; tr3.Font.Size = 9;

                row = 3;
                foreach (string ym in dateseries)
                {
                    var monthbugtype = bugtype.Where(wi => DateTime.Parse(wi.CreatedDate).ToString("yyyy-MM") == ym);
                    var data = bugtbl.Table.Cell(row++, col).Shape.TextFrame.TextRange;
                    data.Text = monthbugtype.Count().ToString();
                    data.Font.Size = 9;
                    data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;
                }

                col += 2;
            }
            bugtbl.Table.Cell(1, col).Merge(bugtbl.Table.Cell(2, col));
            var sumr = bugtbl.Table.Cell(1, col).Shape.TextFrame.TextRange;
            sumr.Text = "合计"; sumr.Font.Size = 9;

            for (row = 3; row < dateseries.Count + 3; row++)
            {
                int sum = 0;
                for (col = 2; col < 2 * bugstype.Count() + 2; col += 2)
                {
                    sum += Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
                }
                var data = bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange;
                data.Text = Convert.ToString(sum);
                data.Font.Size = 9;
                data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;


                for (col = 2; col < 2 * bugstype.Count() + 2; col += 2)
                {
                    var percent = bugtbl.Table.Cell(row, col + 1).Shape.TextFrame.TextRange;
                    var num = Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
                    percent.Text = Convert.ToString(num * 100 / sum) + "%";
                    percent.Font.Size = 9;
                    percent.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;
                }
            }

            bugtbl.Height = 6.0f;

            var subframe2 = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 60.0f, 330f, 200.0f, 50.0f).TextFrame;
            // Add title
            var subtitle2 = subframe2.TextRange;
            subtitle2.Text = "分析说明：";
            subtitle2.Font.NameFarEast = "微软雅黑";
            subtitle2.Font.Bold = MsoTriState.msoTrue;
            subtitle2.Font.Color.RGB = 0x00C07000;
            subtitle2.Font.Size = 16;

        }
    }
}
