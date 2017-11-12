using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class OverviewSheet : ExcelSheetBase,IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public OverviewSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            #region 标题
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "J"]];
            Utility.AddNativieResource(titleRange);
            titleRange.ColumnWidth = 10;
            titleRange.RowHeight = 40;
            titleRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            titleRange.Merge();
            sheet.Cells[2, "B"] = "项目整体说明";
            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
            #endregion 标题

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

            #endregion 表格
        }
    }
}
