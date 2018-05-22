﻿using System;
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
    public class CodeReviewSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public CodeReviewSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            BuildTitle();

            List<CodeReviewEntity> list = TFS.WorkItem.CodeReview.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            int startRow = BuildTable(4, list);            

            var bugList = Bug.GetAllByIteration(project.Name, TFS.Utility.GetBestIteration(project.Name));
            startRow = BuildCodeReviewTable(startRow, bugList[5]);

            BuildAnalyzeTable(startRow);

            var range = sheet.get_Range("G1:N1");
            Utility.AddNativieResource(range);
            range.ColumnWidth = 16;

            sheet.Cells[1, "A"] = "";
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "F", "代码审查统计分析");
        }

        private int BuildTable(int startRow, List<CodeReviewEntity> list)
        {
            var allbugs = list.GroupBy(cre => cre.CreatedDate2).OrderBy(bugs => bugs.Key).ToList();

            int nextRow = Utility.BuildFormalTable(sheet, startRow, "本迭代代码审查效率", "", "B", "F",
                new List<string>() { "审查时间", "审查人数", "评审用时（h）", "发现问题数", "效率（个/h）" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F" },
                allbugs.Count
                );

            startRow += 3;
            for (int i = 0; i < allbugs.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = allbugs[i].Key;
                sheet.Cells[startRow + i, "E"] = allbugs[i].Count();
                sheet.Cells[startRow + i, "F"] = String.Format("=E{0}/(C{0}*D{0})", startRow + i);
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow-1, "C"]);
            Utility.SetCellFontRedColor(sheet.Cells[startRow-1, "D"]);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + allbugs.Count - 1, "B"]]);
            if (allbugs.Count > 0) { 
                Utility.SetCellGreenColor(sheet.Range[sheet.Cells[startRow, "F"], sheet.Cells[startRow + allbugs.Count - 1, "F"]]);
                Utility.SetCellFontRedColor(sheet.Range[sheet.Cells[startRow, "F"], sheet.Cells[startRow + allbugs.Count - 1, "F"]]);
            }

            return nextRow -1;
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
            tableTitleFont.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "P"]];
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

        private int BuildCodeReviewTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代代码评审发现问题统计", "说明：", "B", "P",
                new List<string>() { "BugID", "关键应用", "模块", "功能","缺陷发现方式", "问题类别", "严重级别", "Bug标题", "指派给", "发现人" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,I", "J,J","K,K", "L,N", "O,O", "P,P" },
                list.Count);

            var orderedBugs = list.OrderBy(bug => bug.KeyApplicationName).ThenBy(bug => bug.ModulesName).ThenBy(bug => bug.FuncName).ToList();
            startRow += 3;
            object[,] arr = new object[list.Count, 15];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = orderedBugs[i].Id.ToString();
                arr[i, 1] = orderedBugs[i].KeyApplicationName;
                arr[i, 3] = orderedBugs[i].ModulesName;
                arr[i, 5] = orderedBugs[i].FuncName;
                arr[i, 7] = orderedBugs[i].DetectionMode;
                arr[i, 8] = orderedBugs[i].Type;
                arr[i, 9] = orderedBugs[i].Severity;
                arr[i, 10] = orderedBugs[i].Title;
                arr[i, 12] = Utility.GetPersonName(orderedBugs[i].AssignedTo);
                arr[i, 13] = Utility.GetPersonName(orderedBugs[i].DiscoveryUser);
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedBugs.Count - 1, "P"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedBugs.Count - 1, "B"]]);

            return nextRow-1;
        }
    }
}
