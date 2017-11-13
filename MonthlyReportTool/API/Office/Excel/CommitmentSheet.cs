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

            List<CommitmentEntity> list = new List<CommitmentEntity>() {new CommitmentEntity(), new CommitmentEntity(), new CommitmentEntity(), new CommitmentEntity(), new CommitmentEntity(), new CommitmentEntity() };
            int startRow = BuildFailedTable(12,list);
            startRow = BuildFailedReasonTable(startRow,list);
            BuildExceptionTable(startRow,list);
        }

        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "G", "提交单统计分析");
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
        private int BuildFailedTable(int startRow, List<CommitmentEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单打回次数及一次通过率，按人员统计", "说明：一次通过率=一次通过数/提交单总数\r\n按一次通过率排序", "B", "G",
                new List<string>() { "姓名", "一次通过数", "被打回一次数", "被打回两次数", "被打回两次以上数", "一次通过率" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F", "G,G"},
                list.Count);

            return nextRow;
        }

        private int BuildFailedReasonTable(int startRow, List<CommitmentEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单打回原因分析", "说明：", "B", "H",
                new List<string>() { "提交单ID", "提交单类型", "打回次数", "打回原因", "功能负责人", "测试负责人" },
                new List<string>() { "B,B", "C,C", "D,D", "E,F", "G,G", "H,H"},
                list.Count);

            return nextRow;            

        }

        private int BuildExceptionTable(int startRow, List<CommitmentEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单持续时间超过2周的异常分析", "说明：提交单状态从【提交测试】到【测试通过】这段时间超过2周的\r\n提交日期和测试通过时间比较", "B", "H",
                new List<string>() { "提交单ID", "持续时间", "原因分析", "功能负责人", "测试负责人" },
                new List<string>() { "B,B", "C,C", "D,F", "G,G", "H,H" },
                list.Count);

            return nextRow;
        }
    }
}
