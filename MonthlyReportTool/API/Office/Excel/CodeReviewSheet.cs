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

            BuildAnalyzeTable(startRow);

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

            Utility.SetCellRedColor(sheet.Cells[startRow-1, "C"]);
            Utility.SetCellRedColor(sheet.Cells[startRow-1, "D"]);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + allbugs.Count - 1, "B"]]);

            return nextRow;
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
