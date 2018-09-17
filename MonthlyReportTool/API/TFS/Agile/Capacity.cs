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
        public static double GetIterationCapacities(string prj, string iterationId)
        {
            var list = GetIterationCapacitiesForTeamMember(prj, iterationId);
            double capacities = 0;
            foreach (var p in list.Keys)
            {
                capacities += list[p];
            }

            return capacities;
        }

        public static Dictionary<string,double> GetIterationCapacitiesForTeamMember(string prj, string iterationId)
        {
            Dictionary<string, double> list = new Dictionary<string, double>();
            string ret = TFS.Utility.GetHttpResponseByUrl(
                String.Format("http://tfs.teld.cn:8080/tfs/teld/{0}/_apis/work/teamsettings/iterations/{1}?api-version=v4.1-preview.1", prj, iterationId)
            );

            string url = (JsonConvert.DeserializeObject(ret) as JObject)["_links"]["capacity"]["href"].ToString().Trim();
            ret = TFS.Utility.GetHttpResponseByUrl(url);
            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["value"] as JArray;
            foreach (var jo in jarray)
            {
                var arr = jo["activities"] as JArray;
                if (arr.Count > 0)
                {
                    double capacity = 0.0d;
                    foreach (var p in arr)
                    {
                        capacity += Convert.ToDouble(p["capacityPerDay"]);
                        
                    }
                    list.Add(jo["teamMember"]["displayName"].ToString(),capacity);
                }
            }

            return list;
        }
    }
}
