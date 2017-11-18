using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.Agile
{
    public class IterationEntity
    {
        public string Id;
        public string Name;
        public string Path;
        public string StartDate;
        public string EndDate;

        public override string ToString()
        {
            return String.Format("ID={0}, Name={1}, Path={2}, Start={3}, End={4}", "", Name, Path, StartDate, EndDate);
        }
    }
}
