using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace MonthlyReportTool.API.Office.PowerPoint
{
    public abstract class PowerPointSlideBase
    {
            public PowerPointSlideBase(PowerPointInterop.Slide slide)
            {
                //Utility.SetSheetFont(sheet);
            }
    }
}
