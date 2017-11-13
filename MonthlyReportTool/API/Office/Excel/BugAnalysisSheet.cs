using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;

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

            List<BugEntity> list = new List<BugEntity>() { new BugEntity(), new BugEntity(), new BugEntity(), new BugEntity(), new BugEntity() };

            int startRow = BuildFixedRateTable(12, list);
            startRow = BuildReasonTable(startRow, list);
            startRow = BuildAddedTable(startRow, list);
            startRow = BuildNoneTable(startRow, list);
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "I", "Bug统计分析");
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

        private int BuildFixedRateTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(sheet, startRow, "按人员统计的修复率", "说明：", "B", "F",
                new List<string>() { "姓名", "本迭代新增数", "本迭代遗留数", "本迭代修复数", "Bug修复率" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F" },
                list.Count
                );
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[startRow + i, "F"] = String.Format("=E{0}/(D{0}+E{0})", startRow + i);
            }

            sheet.Cells[startRow + list.Count, "B"] = "合计";
            sheet.Cells[startRow + list.Count, "C"] = String.Format("=SUM(C{0}:C{1})", startRow, startRow + list.Count - 1);
            sheet.Cells[startRow + list.Count, "D"] = String.Format("=SUM(D{0}:D{1})", startRow, startRow + list.Count - 1);
            sheet.Cells[startRow + list.Count, "E"] = String.Format("=SUM(E{0}:E{1})", startRow, startRow + list.Count - 1);
            sheet.Cells[startRow + list.Count, "F"] = String.Format("=E{0}/(D{0}+E{0})", startRow + list.Count);


            return nextRow + 1;
        }

        private int BuildReasonTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "Bug产生原因分析", "说明：主要针对严重级别为1、2级的Bug进行原因分析", "B", "L",
                new List<string>() { "BugID", "问题类别", "严重级别", "Bug标题", "原因分析", "指派给", "测试人员" },
                new List<string>() { "B,B", "C,C", "D,D", "E,H", "I,J", "K,K", "L,L" },
                10);

            return nextRow;
        }
        private int BuildAddedTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代新增Bug数", "说明：", "B", "J",
                new List<string>() { "BugID", "问题类别", "严重级别", "Bug标题", "指派给", "测试人员" },
                new List<string>() { "B,B", "C,C", "D,D", "E,H", "I,I", "J,J" },
                10);

            return nextRow;
        }
        private int BuildNoneTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代处理的不是错误/不予处理Bug分析", "说明：", "B", "L",
                new List<string>() { "BugID", "关闭原因", "问题类别", "严重级别", "Bug标题", "不是错误/不予处理分析" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,H", "I,L" },
                10);

            return nextRow;
        }
    }
}
