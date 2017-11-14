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
    public class PerformanceSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public PerformanceSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(string project)
        {
            BuildTitle();

            int startRow = BuildTable(4);
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "F", "人员考评结果");
        }

        private int BuildTable(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "", "", "B", "F",
                new List<string>() { "姓名", "业绩初评", "加分项", "减分项", "总得分"},
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F"},
                10);

            return 12;
        }
    }
}
