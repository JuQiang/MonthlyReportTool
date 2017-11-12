using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class CommitmentSheet : ExcelSheetBase,IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public CommitmentSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            BuildTitle();
            BuildSubTitle();

            BuildTestTable();
            BuildPerformanceTestTable();

            int startRow = BuildFailedTable();
            startRow = BuildFailedReasonTable(startRow);
            BuildExceptionTable(startRow);
        }

        private void BuildTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "G"]];
            Utility.AddNativieResource(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "提交单统计分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;

            ExcelInterop.Range colall = sheet.Range[sheet.Cells[1, "B"], sheet.Cells[1, "I"]];
            Utility.AddNativieResource(colall);
            colall.ColumnWidth = 16;
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "E"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "提交单测试情况（不包含运维SQL）";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

            ExcelInterop.Range range2 = sheet.Range[sheet.Cells[4, "G"], sheet.Cells[4, "I"]];
            Utility.AddNativieResource(range2);
            range2.Merge();
            sheet.Cells[4, "G"] = "提交单性能测试统计";
            var titleFont2 = range2.Font;
            Utility.AddNativieResource(titleFont2);
            titleFont2.Bold = true;
            titleFont2.Size = 12;

        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "O"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Name = "微软雅黑";
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;
        }
        private void BuildTestTable()
        {
            string[,] cols = new string[,]
                        {
                { "分类", "迭代提交单", "Hotfix补丁-紧急需求", "Hotfix补丁-BUG"},
                { "提交单测试通过数", "", "", ""},
                { "遗留提交单数", "", "", ""},
                { "需测试提交单总数", "", "", "" },
                { "通过率", "", "", ""},
                        };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
            };
            for (int row = 0; row < cols.GetLength(0); row++)
            {
                for (int col = 0; col < colsname.Count; col++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[5 + row, colsname[col].Item1], sheet.Cells[5 + row, colsname[col].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[5 + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            BuildTestTableTitle();

            sheet.Cells[9, "C"] = "=IF(C6<>0,C6/C8,\"\")";
            sheet.Cells[9, "D"] = "=IF(D6<>0,D6/D8,\"\")";
            sheet.Cells[9, "E"] = "=IF(E6<>0,E6/E8,\"\")";
        }
        private void BuildTestTableTitle()
        {
            ExcelInterop.Range colRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "E"]];
            Utility.AddNativieResource(colRange);
            colRange.RowHeight = 20;
            colRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = colRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var colFont = colRange.Font;
            Utility.AddNativieResource(colFont);
            colFont.Bold = true;
        }

        private void BuildPerformanceTestTable()
        {
            string[,] cols = new string[,]
                        {
                { "分类", "个数"},
                { "提交单测试通过数", ""},
                { "需要性能测试提交单总数",""},
                { "性能测试发现BUG数", ""},
                { "通过率", ""},
                        };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("G","H"),
                Tuple.Create<string,string>("I","I"),
            };
            for (int row = 0; row < cols.GetLength(0); row++)
            {
                for (int col = 0; col < colsname.Count; col++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[5 + row, colsname[col].Item1], sheet.Cells[5 + row, colsname[col].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[5 + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            BuildPerformanceTestTableTitle();

            sheet.Cells[9, "I"] = "=IF(I7<>0,I6/I7,\"\")";
        }
        private void BuildPerformanceTestTableTitle()
        {
            ExcelInterop.Range colRange = sheet.Range[sheet.Cells[5, "G"], sheet.Cells[5, "I"]];
            Utility.AddNativieResource(colRange);
            colRange.RowHeight = 20;
            colRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = colRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var colFont = colRange.Font;
            Utility.AddNativieResource(colFont);
            colFont.Bold = true;
        }
        private int BuildFailedTable()
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[12, "B"], sheet.Cells[12, "G"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[12, "B"] = "提交单打回次数及一次通过率，按人员统计";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[13, "B"], sheet.Cells[13, "G"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 40;
            sheet.Cells[13, "B"] = "说明：一次通过率=一次通过数/提交单总数\r\n按一次通过率排序";

            int failedCount = 20;
            string[] cols = new string[] { "姓名", "一次通过数", "被打回一次数", "被打回两次数", "被打回两次以上数", "一次通过率"};
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
                Tuple.Create<string,string>("F","F"),
                Tuple.Create<string,string>("G","G"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[14, colsname[i].Item1], sheet.Cells[14, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[14, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14, "G"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < failedCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[15 + i, colsname[j].Item1], sheet.Cells[15 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    if (j == cols.Length - 1)
                    {
                        sheet.Cells[15 + i, colsname[j].Item1] = String.Format("=IF(C{0}<>0,C{0}/(C{0}+D{0}+E{0}+F{0}),\"\")",15+i);
                    }
                    else
                    {
                        sheet.Cells[15 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 19 + i, j + 1);
                    }

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return 14 + failedCount + 2;
        }

        private int BuildFailedReasonTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "H"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "提交单打回原因分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "H"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[startRow + 1, "B"] = "说明：";

            int reasonCount = 10;
            string[] cols = new string[] { "提交单ID", "提交单类型", "打回次数", "打回原因", "功能负责人", "测试负责人" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","F"),
                Tuple.Create<string,string>("G","G"),
                Tuple.Create<string,string>("H","H"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 2, colsname[i].Item1], sheet.Cells[startRow + 2, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "H"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < reasonCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return startRow + 3 + reasonCount + 1;
        }

        private void BuildExceptionTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "H"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "提交单持续时间超过2周的异常分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "H"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 40;
            sheet.Cells[startRow + 1, "B"] = "说明：提交单状态从【提交测试】到【测试通过】这段时间超过2周的\r\n提交日期和测试通过时间比较";

            int exceptionCount = 13;
            string[] cols = new string[] { "提交单ID", "持续时间", "原因分析", "功能负责人","测试负责人" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","F"),
                Tuple.Create<string,string>("G","G"),
                Tuple.Create<string,string>("H","H"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 2, colsname[i].Item1], sheet.Cells[startRow + 2, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "H"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < exceptionCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            //return startRow + 3 + featuresCount + 2;
        }
    }
}
