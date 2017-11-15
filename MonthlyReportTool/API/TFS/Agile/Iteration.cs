using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.Agile
{
    public class Iteration
    {
        public static List<IterationEntity> GetProjectIterations(string prj)
        {
            List<IterationEntity> list = new List<IterationEntity>();

            string ret = TFS.Utility.GetHttpResponseByUrl(
                String.Format("http://tfs.teld.cn:8080/tfs/teld/{0}/_apis/work/teamsettings/iterations?api-version=v2.0-preview", prj)
            );

            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["value"] as JArray;
            foreach (var jo in jarray)
            {
                list.Add(
                    new IterationEntity()
                    {
                        ID=Convert.ToString(jo["id"]),
                        Name = Convert.ToString(jo["name"]),
                        Path = Convert.ToString(jo["path"]),
                        StartDate= Convert.ToString(jo["attributes"]["startDate"]),
                        EndDate = Convert.ToString(jo["attributes"]["finishDate"])
                    }
                );
            }
            return list;
        }


    }
}
