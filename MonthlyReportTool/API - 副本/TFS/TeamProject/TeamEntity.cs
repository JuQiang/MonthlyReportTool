using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.TeamProject
{
    public class TeamEntity
    {
        public string Id;
        public string Name;
        public string URL;
        public string Description;
        public string IdentityURL;

        public List<MemberEntity> MemberList;

    }
}
