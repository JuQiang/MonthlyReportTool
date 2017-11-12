using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class BugAnalysisSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public BugAnalysisSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            BuildTitle();
            BuildSubTitle();

            BuildDescription();

            BuildSummaryTable();

            int startRow = BuildFixedRateTable();
            BuildImportantTable(startRow);

        }
        private void BuildTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "I"]];
            Utility.AddNativieResource(range);
            range.ColumnWidth = 16;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "Bug统计分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "I"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "Bug新增及处理情况统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

        }

        private void BuildDescription()
        {
            int row = 5;
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "E"]];
            Utility.AddNativieResource(titleRange);
            titleRange.RowHeight = 80;
            titleRange.Merge();
            sheet.Cells[row, "B"] = "说明：\r\n" +
                                    "      Bug修复率 = 本迭代修复数 /（本迭代遗留数 + 本迭代修复数）\r\n" +
                                    "      Bug遗留率 = 本迭代遗留数）/（本迭代遗留数 + 本迭代修复数）\r\n" +
                                    "      1、2级占比 = (1、2级问题数)/ 合计里面的数\r\n" +
                                    "      平均Bug修复耗时：（关闭日期 - 创建日期）/ 总修复数";

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[row, "F"], sheet.Cells[row, "I"]];
            titleRange2.Merge();
            Utility.AddNativieResource(titleRange2);
            titleRange.Merge();
            sheet.Cells[row, "F"] = "\r\n" +
                                    "本迭代新增数：本迭代新登记的Bug数\r\n" +
                                    "本迭代修复数：本迭代修复的所有的Bug数（包括非本迭代登记的Bug）\r\n" +
                                    "本迭代遗留数：本迭代结束后，遗留未修复的所有Bug数（包括非本迭代登记的Bug）";
        }

        private void BuildSummaryTable()
        {
            int start = 6;
            string[,] cols = new string[,]
                        {
                { "", "1 - 严重", "2 - 高", "3 - 中","4 - 低","5 - 无（建议）","合计","1、2级占比"},
                { "本迭代新增数", "", "", "","", "", "",""},
                { "本迭代遗留数", "", "", "","", "", "",""},
                { "本迭代修复数", "", "", "","", "", "",""},
                { "Bug修复率", "", "", "","", "", "",""},
                        };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
                Tuple.Create<string,string>("F","F"),
                Tuple.Create<string,string>("G","G"),
                Tuple.Create<string,string>("H","H"),
                Tuple.Create<string,string>("I","I"),

            };
            for (int row = 0; row < cols.GetLength(0); row++)
            {
                for (int col = 0; col < colsname.Count; col++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[start + row, colsname[col].Item1], sheet.Cells[start + row, colsname[col].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[start + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[start, "B"], sheet.Cells[start, "I"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;


            sheet.Cells[7, "H"] = "=SUM(C7:G7)"; sheet.Cells[8, "H"] = "=SUM(C8:G8)"; sheet.Cells[9, "H"] = "=SUM(C9:G9)";
            sheet.Cells[7, "I"] = "=(C7+D7)/H7"; sheet.Cells[8, "I"] = "=(C8+D8)/H8"; sheet.Cells[9, "I"] = "=(C9+D9)/H9";
            sheet.Cells[10, "C"] = "这里的公式有问题"; sheet.Cells[10, "D"] = "这里的公式有问题"; sheet.Cells[10, "E"] = "这里的公式有问题";
            sheet.Cells[10, "F"] = "这里的公式有问题"; sheet.Cells[10, "G"] = "这里的公式有问题";

        }

        private int BuildFixedRateTable()
        {
            int row = 12;
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "F"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[row, "B"] = "按人员统计的修复率";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            row++;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "F"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[row, "B"] = "说明：";

            int personCount = 10;
            string[] cols = new string[] { "姓名", "本迭代新增数", "本迭代遗留数", "本迭代修复数", "Bug修复率" };
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
            for (int i = 0; i < personCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row + i, colsname[j].Item1], sheet.Cells[row + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    if (j == cols.Length - 1)
                    {
                        sheet.Cells[row + i, colsname[j].Item1] = String.Format("=E{0}/(D{0}+E{0})", row + i);
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

            sheet.Cells[row + personCount, "B"] = "合计";
            sheet.Cells[row + personCount, "C"] = String.Format("=SUM(C{0}:C{1})", row, row + personCount - 1);
            sheet.Cells[row + personCount, "D"] = String.Format("=SUM(D{0}:D{1})", row, row + personCount - 1);
            sheet.Cells[row + personCount, "E"] = String.Format("=SUM(E{0}:E{1})", row, row + personCount - 1);
            sheet.Cells[row + personCount, "F"] = String.Format("=E{0}/(D{0}+E{0})", row + personCount);


            ExcelInterop.Range sumRange = sheet.Range[sheet.Cells[row + personCount, "B"], sheet.Cells[row + personCount, "F"]];
            Utility.AddNativieResource(sumRange);
            var sumBorder = sumRange.Borders;
            Utility.AddNativieResource(sumBorder);
            sumBorder.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            return row + personCount + 1;
        }

        private void BuildImportantTable(int startRow)
        {
            int row = startRow;
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "J"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[row, "B"] = "Bug产生原因分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            row++;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "J"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[row, "B"] = "说明：主要针对严重级别为1、2级的Bug进行原因分析";

            int importantCount = 10;
            string[] cols = new string[] { "BugID", "问题类别", "严重级别", "Bug标题", "原因分析", "指派给", "测试人员" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","F"),
                Tuple.Create<string,string>("G","H"),
                Tuple.Create<string,string>("I","I"),
                Tuple.Create<string,string>("J","J"),
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

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "J"]];
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
            for (int i = 0; i < importantCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row + i, colsname[j].Item1], sheet.Cells[row + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();

                    sheet.Cells[row + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", row + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }
        }
    }
}
