using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.TeamProject;

namespace MonthlyReportTool.API.Office.Excel
{
    public class OverviewSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public OverviewSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "J", "项目整体说明");

            #region 标题2
            ExcelInterop.Range title2Range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "J"]];
            Utility.AddNativieResource(title2Range);
            title2Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignLeft;
            title2Range.Merge();

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(tableRange);
            var tableBorder = tableRange.Borders;
            Utility.AddNativieResource(tableBorder);
            tableBorder.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 标题2

            #region 表格
            for (int row = 5; row <= 12; row++)
            {
                ExcelInterop.Range table2Range = sheet.Range[sheet.Cells[row, "C"], sheet.Cells[row, "J"]];
                Utility.AddNativieResource(table2Range);
                var table2Border = table2Range.Borders;
                Utility.AddNativieResource(table2Border);
                table2Border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                table2Range.Merge();
            }

            ExcelInterop.Range leftRange = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(leftRange);
            leftRange.ColumnWidth = 30.67;
            leftRange.RowHeight = 20;
            var leftFont = leftRange.Font;
            Utility.AddNativieResource(leftFont);
            leftFont.Bold = true;
            leftFont.Name = "微软雅黑";
            leftFont.Size = 11;

            sheet.Cells[4, "B"] = "迭代期间及人员情况综述";
            sheet.Cells[5, "B"] = "Sprint期间";
            sheet.Cells[6, "B"] = "项目负责人";
            sheet.Cells[7, "B"] = "开发负责人";
            sheet.Cells[8, "B"] = "开发人员";
            sheet.Cells[9, "B"] = "需求人员";
            sheet.Cells[10, "B"] = "UI人员";
            sheet.Cells[11, "B"] = "测试负责人";
            sheet.Cells[12, "B"] = "测试人员";

            var ie = TFS.Utility.GetBestIteration(project.Name);
            sheet.Cells[5, "C"] = String.Format("{0} - {1}", DateTime.Parse(ie.StartDate).ToLongDateString(), DateTime.Parse(ie.EndDate).ToLongDateString());

            #endregion 表格

            int nextRow = Utility.BuildFormalTable(this.sheet, 14, "迭代燃尽图", "说明：主要针对以下几种异常情况做说明：\r\n1、迭代初期任务安排不饱和\r\n2、迭代进行中，剩余工作偏离理想趋势太多\r\n3、迭代结束，剩余工作还有很多未完成\r\n4、可用容量和理想趋势差别较大",
                "B", "J",
                new List<string>() { "此处对燃尽异常进行分析说明" },
                new List<string>() { "B,J" },
                5
                );

            ExcelInterop.Range range = sheet.Range[sheet.Cells[16, "B"], sheet.Cells[16 + 2 + 5, "J"]];
            Utility.AddNativieResource(range);
            range.Merge();

            string burdownPicturePath = TFS.Utility.GetBurndownPictureFile(project.Name);
            var shapes = sheet.Shapes;
            Utility.AddNativieResource(shapes);
            shapes.AddPicture(burdownPicturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 50, 550, 1248 * 2 / 3, 616 * 2 / 3);

            ExcelInterop.Range colRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(colRange);
            var interior = colRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

        }
    }
}
