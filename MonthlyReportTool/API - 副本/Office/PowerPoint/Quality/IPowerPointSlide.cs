using MonthlyReportTool.API.TFS.TeamProject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.Office.PowerPoint.Quality
{
    public interface IPowerPointQualitySlide
    {
        void Build(ProjectEntity project,string yearmonth);
    }
}
