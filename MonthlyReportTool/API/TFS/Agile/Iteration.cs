using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;

namespace MonthlyReportTool.API.TFS.Agile
{
    public class Iteration
    {
        public static List<IterationEntity> GetProjectIterations(string prj)
        {
            List<IterationEntity> list = new List<IterationEntity>();

            string url = String.Format("{0}/{1}/_apis/work/teamsettings/iterations?api-version=v4.1-preview.1",Utility.BaseUrl,prj);

            string ret = TFS.Utility.GetHttpResponseByUrl(url);

            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["value"] as JArray;
            foreach (var jo in jarray)
            {
                list.Add(
                    new IterationEntity()
                    {
                        Id = Convert.ToString(jo["id"]),
                        Name = Convert.ToString(jo["name"]),
                        Path = Convert.ToString(jo["path"]),
                        StartDate = Convert.ToString(jo["attributes"]["startDate"]),
                        EndDate = Convert.ToString(jo["attributes"]["finishDate"])
                    }
                );
            }
            return list;
        }

        public static List<string> GetProjectIterationDaysOff(string prj, string iteration)
        {
            List<string> daysOff = new List<string>();

            string ret = TFS.Utility.GetHttpResponseByUrl(
                String.Format("{0}/{1}/_apis/work/teamsettings/iterations/{2}?api-version=v4.1-preview.1", Utility.BaseUrl,prj,iteration)
            );

            string url = (JsonConvert.DeserializeObject(ret) as JObject)["_links"]["teamDaysOff"]["href"].ToString().Trim();
            ret = TFS.Utility.GetHttpResponseByUrl(url);
            var jarray = (JsonConvert.DeserializeObject(ret) as JObject)["daysOff"] as JArray;
            foreach (var jo in jarray)
            {
                string start = jo["start"].ToString();
                string end = jo["end"].ToString();

                if (start == end)
                {
                    daysOff.Add(start);
                }
                else
                {
                    DateTime startDate = DateTime.Parse(start);
                    DateTime endDate = DateTime.Parse(end);

                    daysOff.Add(start);
                    int days = (int)((endDate - startDate).TotalDays);
                    for (int i = 0; i < days; i++)
                    {
                        daysOff.Add(startDate.AddDays(i + 1).ToString("yyyy/M/d HH:mm:ss"));
                    }
                }
            }

            return daysOff;
        }
    }
}
