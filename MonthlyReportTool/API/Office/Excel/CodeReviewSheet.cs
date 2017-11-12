using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class CodeReviewSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public CodeReviewSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            BuildTitle();

            int startRow = BuildTable();
            BuildAnalyzeTable(startRow);
        }
        private void BuildTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "F"]];
            Utility.AddNativieResource(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "代码审查统计分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
        }

        private int BuildTable()
        {
            int row = 4;
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "F"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.ColumnWidth = 20;
            tableTitleRange.RowHeight = 20;
            sheet.Cells[row, "B"] = "本迭代代码审查效率";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            int codereviewCount = 20;
            string[] cols = new string[] { "审查时间", "审查人数", "评审用时（h）", "发现问题数", "效率（个/h）"};
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
                Tuple.Create<string,string>("F","F"),
            };

            row++;
            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row, colsname[i].Item1], sheet.Cells[row, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[row, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "F"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;

            row++;
            //TODO : 放入GIT
            for (int i = 0; i < codereviewCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row + i, colsname[j].Item1], sheet.Cells[row + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    if (j == cols.Length - 1)
                    {
                        sheet.Cells[row + i, colsname[j].Item1] = String.Format("=E{0}/(C{0}*D{0})", row + i);
                        //=E6/(C6*D6)
                    }
                    else
                    {
                        sheet.Cells[row + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", row + i, j + 1);
                    }

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return row + codereviewCount;
        }

        private void BuildAnalyzeTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "F"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "代码审查分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "M"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[startRow + 1, "B"] = "说明：针对本迭代做的代码审查工作做分析";

            ExcelInterop.Range descRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 10, "F"]];
            descRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignLeft;
            descRange.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignTop;
            Utility.AddNativieResource(descRange);
            descRange.Merge();

            ExcelInterop.Borders descBorder = descRange.Borders;
            Utility.AddNativieResource(descBorder);
            descBorder.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
        }        
    }
}
