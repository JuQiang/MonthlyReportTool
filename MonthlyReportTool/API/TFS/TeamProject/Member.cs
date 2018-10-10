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
        public static List<MemberEntity> RetrieveMemberListByTeam(string prjname,string teamname)
        {
            List<MemberEntity> memberlist = new List<MemberEntity>();
            string url = String.Format("{0}/_apis/projects/{1}/teams/{2}/members?api-version=4.1",
                   Utility.BaseUrl,
                    prjname,
                    teamname
                    );

            string responseBody = TFS.Utility.GetHttpResponseByUrl(url);


            foreach (var prj in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                MemberEntity me = new MemberEntity();
                me.Id = Convert.ToString(prj["identity"]["id"]);
                me.DisplayName = Convert.ToString(prj["identity"]["displayName"]);
                me.UniqueName = Convert.ToString(prj["identity"]["uniqueName"]);
                me.URL = Convert.ToString(prj["identity"]["url"]);
                me.ImageURL = Convert.ToString(prj["identity"]["imageUrl"]);
                me.FullName = Convert.ToString(prj["identity"]["displayName"]) + " <" + Convert.ToString(prj["identity"]["uniqueName"]) + ">";
                memberlist.Add(me);
            }

            return memberlist;
        }

        public static List<MemberEntity> RetrieveMemberListByIteration(string prj, string iterationId)
        {
            List<MemberEntity> list = new List<MemberEntity>();

            string ret = TFS.Utility.GetHttpResponseByUrl(
                String.Format("{0}/{1}/_apis/work/teamsettings/iterations/{2}?api-version=v4.1-preview.1", Utility.BaseUrl,prj, iterationId)
            );


            string url = (JsonConvert.DeserializeObject(ret) as JObject)["_links"]["capacity"]["href"].ToString().Trim();
            ret = TFS.Utility.GetHttpResponseByUrl(url);
            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["value"] as JArray;
            foreach (var jo in jarray)
            {
                    list.Add(new MemberEntity()
                    {
                        Id = Convert.ToString(jo["teamMember"]["id"]),
                        DisplayName = Convert.ToString(jo["teamMember"]["displayName"]),
                        UniqueName = Convert.ToString(jo["teamMember"]["uniqueName"]),
                        URL = Convert.ToString(jo["teamMember"]["url"]),
                        ImageURL = Convert.ToString(jo["teamMember"]["imageUrl"]),
                    }
                );
            }

            return list;
        }
    }
}
