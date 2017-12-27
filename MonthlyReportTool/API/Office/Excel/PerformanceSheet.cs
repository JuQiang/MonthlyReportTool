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
    public class PerformanceSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        private ProjectEntity project;
        public PerformanceSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.project = project;
            BuildTitle();

            int startRow = BuildTable(4);

            sheet.Cells[1, "A"] = "";
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "F", "人员考评结果");
            Utility.SetCellRedColor(sheet.Cells[2, "B"]);
        }

        private int BuildTable(int startRow)
        {
            var persons = API.TFS.TeamProject.Member.RetrieveMemberListByIteration(this.project.Name, API.TFS.Utility.GetBestIteration(this.project.Name).Id);

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "", "", "B", "F",
                new List<string>() { "姓名", "业绩初评", "加分项", "减分项", "总得分"},
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F"},
                persons.Count);

            startRow += 3;
            for (int i = 0; i < persons.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] =Utility.GetPersonName(persons[i].DisplayName);
            }

            return nextRow-1;
        }
    }
}
