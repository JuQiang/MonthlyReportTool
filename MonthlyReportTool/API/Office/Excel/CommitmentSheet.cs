using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;

namespace MonthlyReportTool.API.Office.Excel
{
    public class CommitmentSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        private List<List<CommitmentEntity>> commitmentList;
        public CommitmentSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.commitmentList = TFS.WorkItem.Commitment.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            BuildTitle();
            BuildSubTitle();

            BuildTestTable();
            BuildPerformanceTestTable();

            int startRow = BuildFailedTable(12, this.commitmentList[0]);
            startRow = BuildSuccessTable(startRow, this.commitmentList[1]);//测试通过提交单Bug数统计
            startRow = BuildFailedReasonTable(startRow, this.commitmentList[5]);
            startRow = BuildRemovedReasonTable(startRow, this.commitmentList[2]);
            BuildExceptionTable(startRow, this.commitmentList[1]);

            var range = sheet.get_Range("J1:M1");
            Utility.AddNativieResource(range);
            range.ColumnWidth = 16;

            sheet.Cells[1, "A"] = "";
        }

        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "I", "提交单统计分析");
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
                { "已移除/中止提交单数", "", "", ""},
                { "提交单总数", "", "", "" },
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

            Utility.SetCellBorder(sheet.Range[sheet.Cells[5, colsname[0].Item1], sheet.Cells[5 + cols.GetLength(0) - 1, colsname[0].Item2]]);

            BuildTestTableTitle();

            sheet.Cells[7, "C"] = "=C9-C6-C8";
            sheet.Cells[7, "D"] = "=D9-D6-D8";
            sheet.Cells[7, "E"] = "=E9-E6-E8";

            sheet.Cells[10, "C"] = "=IF(C6<>0,C6/C9,\"\")";
            sheet.Cells[10, "D"] = "=IF(D6<>0,D6/D9,\"\")";
            sheet.Cells[10, "E"] = "=IF(E6<>0,E6/E9,\"\")";

            var testpassed = this.commitmentList[1];
            var removed = this.commitmentList[2];
            var all = this.commitmentList[0];

            sheet.Cells[6, "C"] = testpassed.Where(commitment => commitment.SubmitType == "迭代提交单").Count();
            sheet.Cells[6, "D"] = testpassed.Where(commitment => commitment.SubmitType == "Hotfix补丁-紧急需求").Count();
            sheet.Cells[6, "E"] = testpassed.Where(commitment => commitment.SubmitType == "Hotfix补丁-BUG").Count();

            sheet.Cells[8, "C"] = removed.Where(commitment => commitment.SubmitType == "迭代提交单").Count();
            sheet.Cells[8, "D"] = removed.Where(commitment => commitment.SubmitType == "Hotfix补丁-紧急需求").Count();
            sheet.Cells[8, "E"] = removed.Where(commitment => commitment.SubmitType == "Hotfix补丁-BUG").Count();

            sheet.Cells[9, "C"] = all.Where(commitment => commitment.SubmitType == "迭代提交单").Count();
            sheet.Cells[9, "D"] = all.Where(commitment => commitment.SubmitType == "Hotfix补丁-紧急需求").Count();
            sheet.Cells[9, "E"] = all.Where(commitment => commitment.SubmitType == "Hotfix补丁-BUG").Count();

            Utility.SetCellPercentFormat(sheet.get_Range("C10:E10"));
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[7, "C"], sheet.Cells[7, "E"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[10, "C"], sheet.Cells[10, "E"]]);
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
                { "性能测试通过提交单数", ""},
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
                }
            }
            Utility.SetCellBorder(sheet.Range[sheet.Cells[5, colsname[0].Item1], sheet.Cells[5 + cols.GetLength(0) - 1, colsname[colsname.Count - 1].Item2]]);

            BuildPerformanceTestTableTitle();

            sheet.Cells[6, "I"] = this.commitmentList[4].Count;
            sheet.Cells[7, "I"] = this.commitmentList[3].Count;
            sheet.Cells[9, "I"] = "=IF(I7<>0,I6/I7,\"\")";
            Utility.SetCellPercentFormat(sheet.Cells[9, "I"]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[9, "I"], sheet.Cells[9, "I"]]);
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
            var commitments = list.GroupBy(commitment => commitment.AssignedTo);
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单打回次数及一次通过率，按人员统计", "说明：一次通过率=一次通过数/提交单总数\r\n         按一次通过率排序", "B", "G",
                new List<string>() { "姓名", "一次通过数", "被打回一次数", "被打回两次数", "被打回两次以上数", "一次通过率" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F", "G,G" },
                commitments.Count());

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按一次通过率排序");
            startRow += 3;

            List<Tuple<string, int, int, int, int, double>> cells = new List<Tuple<string, int, int, int, int, double>>();
            foreach (var commitment in commitments)
            {
                string person = commitment.Key.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

                int succCount = commitment.Where(comm => comm.BackNum == 0).Count();
                int back1Count = commitment.Where(comm => comm.BackNum == 1).Count();
                int back2Count = commitment.Where(comm => comm.BackNum == 2).Count();
                int back3Count = commitment.Where(comm => comm.BackNum >= 3).Count();

                cells.Add(Tuple.Create<string, int, int, int, int, double>(
                    person, succCount, back1Count, back2Count, back3Count, (double)succCount / (double)(succCount + back1Count + back2Count + back3Count)
                    )
                );
            }

            var orderedCells = cells.OrderByDescending(tuple => tuple.Item6).ToList();

            for (int i = startRow; i < startRow + cells.Count; i++)
            {
                sheet.Cells[i, "B"] = orderedCells[i - startRow].Item1;
                sheet.Cells[i, "C"] = orderedCells[i - startRow].Item2;
                sheet.Cells[i, "D"] = orderedCells[i - startRow].Item3;
                sheet.Cells[i, "E"] = orderedCells[i - startRow].Item4;
                sheet.Cells[i, "F"] = orderedCells[i - startRow].Item5;
                sheet.Cells[i, "G"] = orderedCells[i - startRow].Item6;
            }
            if(cells.Count > 0) { 
            Utility.SetCellPercentFormat(sheet.Range[sheet.Cells[startRow, "G"], sheet.Cells[startRow + cells.Count - 1, "G"]]);
            Utility.SetFormatSmaller(sheet.Range[sheet.Cells[startRow, "G"], sheet.Cells[startRow + cells.Count - 1, "G"]], 1.00d);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + cells.Count - 1, "B"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[startRow, "G"], sheet.Cells[startRow + cells.Count - 1, "G"]]);
            Utility.SetCellFontRedColor(sheet.Range[sheet.Cells[startRow, "G"], sheet.Cells[startRow + cells.Count - 1, "G"]]);
            Utility.SetCellDarkGrayColor(sheet.Range[sheet.Cells[startRow + cells.Count, "B"], sheet.Cells[startRow + cells.Count, "B"]]);
            
            FillSummaryData(startRow, cells.Count);
            }
            return nextRow;
        }

        private void FillSummaryData(int startRow, int rowCount)
        {
            int curRow = startRow + rowCount;
            Utility.SetCellBorder(sheet.Range[sheet.Cells[curRow, "B"], sheet.Cells[curRow, "G"]]);

            sheet.Cells[curRow, "B"] = "合计";
            sheet.Cells[curRow, "C"] = String.Format("=sum(C{0}:C{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "D"] = String.Format("=sum(D{0}:D{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "E"] = String.Format("=sum(E{0}:E{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "F"] = String.Format("=sum(F{0}:F{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "G"] = String.Format("=C{0}/(C{0}+D{0}+E{0}+F{0})", curRow - 0);

            Utility.SetCellPercentFormat(sheet.Cells[curRow, "G"]);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[curRow, "B"], sheet.Cells[curRow, "B"]], hAlign: ExcelInterop.XlHAlign.xlHAlignCenter);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[curRow, "C"], sheet.Cells[curRow, "G"]]);
            Utility.SetCellFontRedColor(sheet.Range[sheet.Cells[curRow, "G"], sheet.Cells[curRow, "G"]]);
        }
        private int BuildSuccessTable(int startRow, List<CommitmentEntity> list)
        {

            var commitments = list.OrderByDescending(comm => comm.SubmitType).ThenByDescending(comm=>comm.FindedBugCount).ToList();

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "测试通过提交单Bug数统计", "说明：按提交单类型，发现的Bug数排序。", "B", "J",
                new List<string>() { "提交单ID", "提交单类型", "提交单名称", "状态", "发现的Bug数", "功能负责人", "测试负责人"},
                new List<string>() { "B,B", "C,C", "D,F", "G,G", "H,H", "I,I", "J,J"},
                commitments.Count);
            
            startRow += 3;
            for (int i = 0; i < commitments.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = commitments[i].Id;
                sheet.Cells[i + startRow, "C"] = commitments[i].SubmitType;
                sheet.Cells[i + startRow, "D"] = commitments[i].Title;
                sheet.Cells[i + startRow, "G"] = commitments[i].State;
                sheet.Cells[i + startRow, "H"] = commitments[i].FindedBugCount;
                sheet.Cells[i + startRow, "I"] = Utility.GetPersonName(commitments[i].AssignedTo);
                sheet.Cells[i + startRow, "J"] = Utility.GetPersonName(commitments[i].TestResponsibleMan);
            }
            
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + commitments.Count - 1, "B"]]);

            return nextRow - 1;
        }
        private int BuildFailedReasonTable(int startRow, List<CommitmentEntity> list)
        {

            var commitments = list.OrderByDescending(comm => comm.BackNum).ToList();

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单打回原因分析", "说明：这个表格很长，请右拉把后面列都填写上。", "B", "O",
                new List<string>() { "提交单ID", "提交单类型", "提交单名称", "状态", "打回次数", "打回原因", "功能负责人", "测试负责人", "后续改进措施" },
                new List<string>() { "B,B", "C,C", "D,F", "G,G", "H,H", "I,J", "K,K", "L,L", "M,O" },
                commitments.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "这个表格很长，请右拉把后面列都填写上。");
            

            startRow += 3;
            for (int i = 0; i < commitments.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = commitments[i].Id;
                sheet.Cells[i + startRow, "C"] = commitments[i].SubmitType;
                sheet.Cells[i + startRow, "D"] = commitments[i].Title;
                sheet.Cells[i + startRow, "G"] = commitments[i].State;
                sheet.Cells[i + startRow, "H"] = commitments[i].BackNum;
                sheet.Cells[i + startRow, "I"] = "";
                if (commitments[i].AssignedTo.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "K"] = Utility.GetPersonName(commitments[i].AssignedTo);
                }
                if (commitments[i].TestResponsibleMan.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "L"] = Utility.GetPersonName(commitments[i].TestResponsibleMan);
                }
                sheet.Cells[i + startRow, "M"] = "";
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "I"]);
            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "M"]);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + commitments.Count - 1, "B"]]);

            return nextRow-1;

        }

        private int BuildRemovedReasonTable(int startRow, List<CommitmentEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "移除/中止提交单原因分析", "说明：", "B", "L",
                new List<string>() { "提交单ID", "提交单类型", "提交单名称", "状态", "原因分析", "功能负责人", "测试负责人" },
                new List<string>() { "B,B", "C,C", "D,F", "G,G", "H,J", "K,K", "L,L" },
                list.Count);

            startRow += 3;
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = list[i].Id;
                sheet.Cells[i + startRow, "C"] = list[i].SubmitType;
                sheet.Cells[i + startRow, "D"] = list[i].Title;
                sheet.Cells[i + startRow, "G"] = list[i].State;
                sheet.Cells[i + startRow, "H"] = "";

                if (list[i].AssignedTo.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "K"] = Utility.GetPersonName(list[i].AssignedTo);
                }
                if (list[i].TestResponsibleMan.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "L"] = Utility.GetPersonName(list[i].TestResponsibleMan);
                }
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "H"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);

            return nextRow-1;

        }

        private int BuildExceptionTable(int startRow, List<CommitmentEntity> list)
        {
            var commitments = list.Where(comm => (
            comm.TestFinishedTime.Trim().Length > 0 &&
            comm.SubmitDate.Trim().Length > 0 &&
            (DateTime.Parse(comm.TestFinishedTime) - DateTime.Parse(comm.SubmitDate)).TotalDays >= 14.0d
            )
            ).ToList();

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "提交单持续时间超过2周的异常分析", "说明：提交单状态从【提交测试】到【测试通过】这段时间超过2周的\r\n提交日期和测试通过时间比较", "B", "M",
                new List<string>() { "提交单ID", "提交单类型", "提交单名称", "状态", "持续时间", "原因分析", "功能负责人", "测试负责人" },
                new List<string>() { "B,B", "C,C", "D,F", "G,G", "H,H", "I,K", "L,L", "M,M" },
                commitments.Count);

            startRow += 3;
            for (int i = 0; i < commitments.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = commitments[i].Id;
                sheet.Cells[i + startRow, "C"] = commitments[i].SubmitType;
                sheet.Cells[i + startRow, "D"] = commitments[i].Title;
                sheet.Cells[i + startRow, "G"] = commitments[i].State;

                sheet.Cells[i + startRow, "H"] = (int)((DateTime.Parse(commitments[i].TestFinishedTime) - DateTime.Parse(commitments[i].SubmitDate)).TotalDays) + "天";
                sheet.Cells[i + startRow, "I"] = "";

                if (commitments[i].AssignedTo.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "L"] = Utility.GetPersonName(commitments[i].AssignedTo);
                }
                if (commitments[i].TestResponsibleMan.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "M"] = Utility.GetPersonName(commitments[i].TestResponsibleMan);
                }
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "I"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + commitments.Count - 1, "B"]]);

            return nextRow-1;
        }
    }
}
