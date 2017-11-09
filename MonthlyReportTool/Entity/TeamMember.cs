using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.Entity
{
    public class TeamMember
    {
        public string DisplayName { get; set; }
        public string MailAccount { get; set; }
        public string TFSAccount
        {
            get
            {
                return String.Format("{0} <{1}>", this.DisplayName, this.MailAccount);
            }
        }
    }

    

}
