using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.TeamProject
{
    public class Project
    {
        public static List<ProjectEntity> RetrieveProjectList()
        {
            List<ProjectEntity> prjlist = new List<ProjectEntity>();
            //http://tfs.teld.cn:8080/tfs/teld/_apis/projects?api-version=2.0
            string url = String.Format("http://{0}:8080/{1}/_apis/projects?api-version=2.0",
                    "tfs.teld.cn",
                    "tfs/teld"
                    );

            string responseBody = TFS.Utility.GetHttpResponseByUrl(url);


            foreach (var prj in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                prjlist.Add(new ProjectEntity()
                {
                    Id = Convert.ToString(prj["id"]),
                    Name = Convert.ToString(prj["name"]),
                    Description = Convert.ToString(prj["description"]),
                    URL = Convert.ToString(prj["url"]),
                    State = Convert.ToString(prj["state"]),
                    Revision = Convert.ToInt32(prj["revision"]),
                }
                );
            }

            return prjlist;
        }
    }
}
