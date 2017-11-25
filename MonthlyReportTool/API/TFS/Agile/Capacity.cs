using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonthlyReportTool.API.TFS.Agile
{
    public class Capacity
    {
        public static int GetIterationCapacities(string prj,string iterationId)
        {
            string ret = TFS.Utility.GetHttpResponseByUrl(
                String.Format("http://tfs.teld.cn:8080/tfs/teld/{0}/_apis/work/teamsettings/iterations/{1}?api-version=v2.0-preview", prj, iterationId)
            );

            int capacities = 0;

            string url = (JsonConvert.DeserializeObject(ret) as JObject)["_links"]["capacity"]["href"].ToString().Trim();
            ret = TFS.Utility.GetHttpResponseByUrl(url);
            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["value"] as JArray;
            foreach (var jo in jarray)
            {
                var arr = jo["activities"] as JArray;
                if (arr.Count > 0)
                {
                    capacities += Convert.ToInt32(arr[0]["capacityPerDay"]);
                }
            }

            return capacities;
        }
    }
}
