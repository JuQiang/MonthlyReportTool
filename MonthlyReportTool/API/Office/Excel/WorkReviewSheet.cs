using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;
using System.Drawing;

namespace MonthlyReportTool.API.Office.Excel
{
    public class WorkReviewSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        private List<List<WorkReviewEntity>> workReviewList;
        public WorkReviewSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            BuildTitle();
            this.workReviewList = TFS.WorkItem.WorkReview.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));

            int startRow = BuildAllReviewBillTable(4, this.workReviewList[0]);//记录单根据类型统计
            startRow = BuildReviewBillDetailTable(startRow, this.workReviewList[0]);//审查记录单明细表
            startRow = BuildBugDetailTable(startRow, this.workReviewList[1]);//审查记录单上所有的Bug列表

            BuildAnalyzeTable(startRow);

            var range = sheet.get_Range("B1:Q1");
            Utility.AddNativieResource(range);
            range.ColumnWidth = 16;

            sheet.Cells[1, "A"] = "";
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "I", "审查工作统计分析");
        }
        private int BuildAllReviewBillTable(int startRow, List<WorkReviewEntity> list)
        {
            var workreviews = list.GroupBy(workreview => workreview.ReviewBillType);
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代记录单统计", "说明：都是以本迭代已关闭了的记录单为统计依据", "B", "E",
                new List<string>() { "记录单类型", "记录单个数", "发现的Bug数", "平均发现的Bug数" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F", "E,E" },
                workreviews.Count());

            startRow += 3;

            List<Tuple<string, int, int, double>> cells = new List<Tuple<string, int, int, double>>();
            foreach (var workreview in workreviews)
            {
                string billtype = workreview.Key.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

                int billcount = workreview.Count();
                int bugsum = workreview.Sum(work=>work.FindedBugCount);

                cells.Add(Tuple.Create<string, int, int, double>(
                     billtype, billcount, bugsum, (double)bugsum / (double)billcount
                     )
                 );
            }

            var orderedCells = cells.OrderByDescending(tuple => tuple.Item4).ToList();

            for (int i = startRow; i < startRow + cells.Count; i++)
            {
                sheet.Cells[i, "B"] = orderedCells[i - startRow].Item1;
                sheet.Cells[i, "C"] = orderedCells[i - startRow].Item2;
                sheet.Cells[i, "D"] = orderedCells[i - startRow].Item3;
                sheet.Cells[i, "E"] = String.Format("=ROUND(D{0}/C{0},2)", i); //orderedCells[i - startRow].Item4;
            }
            if (cells.Count > 0)
            {
                Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + cells.Count - 1, "B"]]);
                Utility.SetFormatSmaller(sheet.Range[sheet.Cells[startRow, "E"], sheet.Cells[startRow + cells.Count - 1, "E"]], 1.00d);
                Utility.SetCellGreenColor(sheet.Range[sheet.Cells[startRow, "E"], sheet.Cells[startRow + cells.Count - 1, "E"]]);
                Utility.SetCellFontRedColor(sheet.Range[sheet.Cells[startRow, "E"], sheet.Cells[startRow + cells.Count - 1, "E"]]);
                Utility.SetCellDarkGrayColor(sheet.Range[sheet.Cells[startRow + cells.Count, "B"], sheet.Cells[startRow + cells.Count, "B"]]);

                FillSummaryData(startRow, cells.Count);
            }
            return nextRow;
        }
        private void FillSummaryData(int startRow, int rowCount)
        {
            int curRow = startRow + rowCount;
            Utility.SetCellBorder(sheet.Range[sheet.Cells[curRow, "B"], sheet.Cells[curRow, "E"]]);

            sheet.Cells[curRow, "B"] = "合计";
            sheet.Cells[curRow, "C"] = String.Format("=sum(C{0}:C{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "D"] = String.Format("=sum(D{0}:D{1}", startRow, curRow - 1);
            sheet.Cells[curRow, "E"] = String.Format("=ROUND(D{0}/C{0},2)", curRow - 0);
            
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[curRow, "B"], sheet.Cells[curRow, "B"]], hAlign: ExcelInterop.XlHAlign.xlHAlignCenter);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[curRow, "C"], sheet.Cells[curRow, "E"]]);
            Utility.SetCellFontRedColor(sheet.Range[sheet.Cells[curRow, "E"], sheet.Cells[curRow, "E"]]);
        }
        private int BuildReviewBillDetailTable(int startRow, List<WorkReviewEntity> list)
        {

            var workviews = list.OrderBy(work => work.KeyApplicationName).OrderBy(work=>work.ModulesName).OrderBy(work=>work.FuncName).ToList();
            //OrderBy(backlog => backlog.KeyApplicationName).ThenBy(backlog => backlog.ModulesName).ThenBy(backlog => backlog.FuncName).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代审查记录单明细", "说明：都是以本迭代已关闭了的记录单为统计依据\r\n         按关键应用、模块、功能排序", "B", "Q",
                new List<string>() { "记录单ID", "关键应用", "模块", "功能","记录单标题", "发现的Bug数", "记录单类型", "评审负责人","指派给","活动发生日期","关闭日期" },
                new List<string>() {     "B,B",      "C,D", "E,F",  "G,H",      "I,K",        "L,L",        "M,M",      "N,N",   "O,O",        "P,P",    "Q,Q" },
                workviews.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            startRow += 3;
            for (int i = 0; i < workviews.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = workviews[i].Id;
                sheet.Cells[i + startRow, "C"] = workviews[i].KeyApplicationName;
                sheet.Cells[i + startRow, "E"] = workviews[i].ModulesName;
                sheet.Cells[i + startRow, "G"] = workviews[i].FuncName;
                sheet.Cells[i + startRow, "I"] = workviews[i].Title;
                sheet.Cells[i + startRow, "L"] = workviews[i].FindedBugCount;
                sheet.Cells[i + startRow, "M"] = workviews[i].ReviewBillType;
                if (workviews[i].ReviewResponsibleMan.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "N"] = Utility.GetPersonName(workviews[i].ReviewResponsibleMan);
                }
                if (workviews[i].AssignedTo.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "O"] = Utility.GetPersonName(workviews[i].AssignedTo);
                }
                if(workviews[i].ActionDate.Trim().Length >0)
                    sheet.Cells[i + startRow, "P"] = DateTime.Parse(workviews[i].ActionDate).AddHours(8).ToString("yyyy-MM-dd");
                if (workviews[i].ClosedDate.Trim().Length > 0)
                    sheet.Cells[i + startRow, "Q"] = DateTime.Parse(workviews[i].ClosedDate).AddHours(8).ToString("yyyy-MM-dd");
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + workviews.Count - 1, "B"]]);

            return nextRow - 1;
        }
        private int BuildBugDetailTable(int startRow, List<WorkReviewEntity> list)
        {
            var workreviews1 = list.FindAll(work => work.workItemType != "记录单").ToList();
            var workreviews = workreviews1.OrderBy(work => work.KeyApplicationName).OrderBy(work => work.ModulesName).OrderBy(work => work.FuncName).ToList();

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代审查发现问题明细", "说明： 按关键应用、模块、功能排序", "B", "Q",
                new List<string>() { "BugID", "关键应用", "模块", "功能", "缺陷发现方式", "问题类别", "严重程度", "Bug标题", "状态", "指派给", "发现人" },
                new List<string>() {   "B,B",     "C,D",  "E,F",  "G,H",         "I,I",     "J,J",     "K,K",     "L,N", "O,O",    "P,P",   "Q,Q" },
                workreviews.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            startRow += 3;
            for (int i = 0; i < workreviews.Count; i++)
            {
                sheet.Cells[i + startRow, "B"] = workreviews[i].Id;
                sheet.Cells[i + startRow, "C"] = workreviews[i].KeyApplicationName;
                sheet.Cells[i + startRow, "E"] = workreviews[i].ModulesName;
                sheet.Cells[i + startRow, "G"] = workreviews[i].FuncName;
                sheet.Cells[i + startRow, "I"] = workreviews[i].DetectionMode;//缺陷发现方式
                sheet.Cells[i + startRow, "J"] = workreviews[i].Type;
                sheet.Cells[i + startRow, "K"] = workreviews[i].Severity;
                sheet.Cells[i + startRow, "L"] = workreviews[i].Title;
                sheet.Cells[i + startRow, "O"] = workreviews[i].State;
                if (workreviews[i].AssignedTo.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "P"] = Utility.GetPersonName(workreviews[i].AssignedTo);
                }
                if (workreviews[i].DiscoveryUser.Trim().Length > 0)
                {
                    sheet.Cells[i + startRow, "Q"] = Utility.GetPersonName(workreviews[i].DiscoveryUser);
                }
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + workreviews.Count - 1, "B"]]);

            return nextRow - 1;
        }

        private void BuildAnalyzeTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "H"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "审查工作分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "P"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[startRow + 1, "B"] = "说明：针对本迭代的审查工作做分析";

            ExcelInterop.Range descRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 10, "H"]];
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
