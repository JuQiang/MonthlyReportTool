using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.Office.Excel
{
    public interface IExcelSheet
    {
        void Build(string project);
    }
}
