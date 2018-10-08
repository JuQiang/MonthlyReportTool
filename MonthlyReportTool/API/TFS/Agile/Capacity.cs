using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;

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
                String.Format("{0}/{1}/_apis/work/teamsettings/iterations/{2}?api-version=v4.1-preview.1",Utility.BaseUrl, prj, iterationId)
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
