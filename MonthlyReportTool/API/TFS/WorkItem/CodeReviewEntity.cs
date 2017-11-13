using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.WorkItem
{
    public class CodeReviewEntity
    {
        public int SystemId;
        public string SystemAreaPath;
        public string SystemTeamProject;
        public string SystemState;
        public string SystemAssignedTo;
        public string SystemCreatedDate;
        public string SystemCreatedBy;
        public string SystemTitle;
        public string MicrosoftVSTSCommonSeverity;
        public string TeldBugType;
        public string TeldBugHopeFixSubmitTime;
        public string TeldBugVerificator;
        public string TeldBugResolvedReason;
        public string CreatedYearMonth;
    }
}
