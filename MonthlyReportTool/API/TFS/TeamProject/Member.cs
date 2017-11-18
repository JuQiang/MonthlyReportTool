using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.TeamProject
{
    public class Member
    {
        public static List<MemberEntity> RetrieveMemberList(string prjname,string teamname)
        {
            List<MemberEntity> memberlist = new List<MemberEntity>();
            string url = String.Format("http://{0}:8080/{1}/_apis/projects/{2}/teams/{3}/members?api-version=1.0",
                    "tfs.teld.cn",
                    "tfs/teld",
                    prjname,
                    teamname
                    );

            string responseBody = TFS.Utility.GetHttpResponseByUrl(url);


            foreach (var prj in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                memberlist.Add(new MemberEntity()
                {
                    Id = Convert.ToString(prj["id"]),
                    DisplayName = Convert.ToString(prj["displayName"]),
                    UniqueName = Convert.ToString(prj["uniqueName"]),
                    URL = Convert.ToString(prj["url"]),
                    ImageURL = Convert.ToString(prj["imageUrl"]),
                    FullName= Convert.ToString(prj["displayName"])+" <" + Convert.ToString(prj["uniqueName"])+">",
                }
                );
            }

            return memberlist;
        }
    }
}
