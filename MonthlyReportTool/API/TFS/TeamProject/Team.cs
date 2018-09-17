using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.TeamProject
{
    public class Team
    {
        public static List<TeamEntity> RetrieveTeamList(string prjname)
        {
            List<TeamEntity> teamlist = new List<TeamEntity>();
            string url = String.Format("http://{0}:8080/{1}/_apis/projects/{2}/teams?api-version=4.1",
                    "tfs.teld.cn",
                    "tfs/teld",
                    prjname
                    );

            string responseBody = TFS.Utility.GetHttpResponseByUrl(url);


            foreach (var prj in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                teamlist.Add(new TeamEntity()
                {
                    Id = Convert.ToString(prj["id"]),
                    Name = Convert.ToString(prj["name"]),
                    Description = Convert.ToString(prj["description"]),
                    URL = Convert.ToString(prj["url"]),
                    IdentityURL = Convert.ToString(prj["identityUrl"]),
                }
                );
            }

            return teamlist;
        }

        
        
    }
}
