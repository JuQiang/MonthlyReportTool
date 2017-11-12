using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;


namespace MonthlyReportTool.API.Office.Excel
{
    public abstract class ExcelSheetBase
    {
        public ExcelSheetBase(ExcelInterop.Worksheet sheet)
        {
            Utility.SetSheetFont(sheet);
        }
    }
}
