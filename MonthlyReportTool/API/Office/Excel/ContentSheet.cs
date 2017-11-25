using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.TeamProject;

namespace MonthlyReportTool.API.Office.Excel
{
    public class ContentSheet : ExcelSheetBase,IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public ContentSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            ExcelInterop.Range allrange = sheet.Range[sheet.Cells[5, "D"], sheet.Cells[40, "J"]];
            Utility.AddNativieResource(allrange);
            allrange.ColumnWidth = 12;
            allrange.RowHeight = 15;
            allrange.Merge();
            allrange.UseStandardHeight = true;
            allrange.WrapText = true;
            allrange.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;

            sheet.Cells[5, "D"] = "这个目录没个鸟用。";

            sheet.Cells[1, "A"] = "";
        }
    }
}
