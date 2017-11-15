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
    public class SuggestionSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public SuggestionSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            BuildTitle();

            int startRow = BuildTable1(4);
            startRow = BuildTable2(startRow);
            startRow = BuildTable3(startRow);
            startRow = BuildTable4(startRow);
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "M", "迭代过程总结", columnWidth : 8);
        }

        private void Format2Columns(int startRow)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow+2, "B"], sheet.Cells[startRow + 6, "C"]];
            Utility.AddNativieResource(range);
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
        }
        private int BuildTable1(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "上迭代改进建议落地情况", "说明：", "B", "M",
                new List<string>() { "序号", "分类", "待改进内容", "负责人", "是否落地", "具体改进说明" },
                new List<string>() { "B,B", "C,C", "D,G", "H,H", "I,I", "J,M" },
                4);

            sheet.Cells[startRow + 3, "B"] = 1; sheet.Cells[startRow + 3, "C"] = "计划";
            sheet.Cells[startRow + 4, "B"] = 2; sheet.Cells[startRow + 4, "C"] = "日常管理";
            sheet.Cells[startRow + 5, "B"] = 3; sheet.Cells[startRow + 5, "C"] = "需求";
            sheet.Cells[startRow + 6, "B"] = 4; sheet.Cells[startRow + 6, "C"] = "跟踪";

            Format2Columns(startRow);

            return 12;
        }

        private int BuildTable2(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "做得好的方面", "说明：迭代过程中好的方面", "B", "M",
                new List<string>() { "序号", "分类", "描述" },
                new List<string>() { "B,B", "C,C", "D,M" },
                4);

            sheet.Cells[startRow + 3, "B"] = 1; sheet.Cells[startRow + 3, "C"] = "计划";
            sheet.Cells[startRow + 4, "B"] = 2; sheet.Cells[startRow + 4, "C"] = "日常管理";
            sheet.Cells[startRow + 5, "B"] = 3; sheet.Cells[startRow + 5, "C"] = "需求";
            sheet.Cells[startRow + 6, "B"] = 4; sheet.Cells[startRow + 6, "C"] = "跟踪";

            Format2Columns(startRow);

            return startRow+8;
        }

        private int BuildTable3(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "待改进的方面", "说明：影响迭代进度、质量等方面的待改进的问题", "B", "M",
                new List<string>() { "序号", "分类", "描述" },
                new List<string>() { "B,B", "C,C", "D,M" },
                4);

            sheet.Cells[startRow + 3, "B"] = 1; sheet.Cells[startRow + 3, "C"] = "计划";
            sheet.Cells[startRow + 4, "B"] = 2; sheet.Cells[startRow + 4, "C"] = "日常管理";
            sheet.Cells[startRow + 5, "B"] = 3; sheet.Cells[startRow + 5, "C"] = "需求";
            sheet.Cells[startRow + 6, "B"] = 4; sheet.Cells[startRow + 6, "C"] = "跟踪";

            Format2Columns(startRow);

            return startRow + 8;
        }

        private int BuildTable4(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "过程改进建议/措施", "说明：待改进方面的建议及改进措施", "B", "M",
                new List<string>() { "序号", "分类", "待改进内容","负责人" },
                new List<string>() { "B,B", "C,C", "D,L","M,M" },
                4);

            sheet.Cells[startRow + 3, "B"] = 1; sheet.Cells[startRow + 3, "C"] = "计划";
            sheet.Cells[startRow + 4, "B"] = 2; sheet.Cells[startRow + 4, "C"] = "日常管理";
            sheet.Cells[startRow + 5, "B"] = 3; sheet.Cells[startRow + 5, "C"] = "需求";
            sheet.Cells[startRow + 6, "B"] = 4; sheet.Cells[startRow + 6, "C"] = "跟踪";

            Format2Columns(startRow);

            return startRow + 8;
        }
    }
}
