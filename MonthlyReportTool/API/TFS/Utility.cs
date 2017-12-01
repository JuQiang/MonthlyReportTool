using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using MonthlyReportTool.API.TFS.TeamProject;
using MonthlyReportTool.API.TFS.WorkItem;
using System.Runtime.InteropServices;
using MonthlyReportTool.API.TFS.Agile;
using System.IO;

namespace MonthlyReportTool.API.TFS
{
    public class Utility
    {
        public static string User = "";
        public static string Pass = "";

        public static string GetHttpResponseByUrl(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", User, Pass))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    return response.Content.ReadAsStringAsync().Result;
                }
            }
        }

        public static string GetBurndownPictureFile(string projectName)
        {
            string url = String.Format("http://tfs.teld.cn:8080/tfs/Teld/{0}/_api/_teamChart/Burndown?chartOptions=%7B%22Width%22%3A1248%2C%22Height%22%3A616%2C%22" +
                "ShowDetails%22%3Atrue%2C%22Title%22%3A%22%22%7D&counter=2&iterationPath={1}&__v=5",
                    projectName,
                    Utility.GetBestIteration(projectName).Path
                    );
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", User, Pass))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    byte[] img = response.Content.ReadAsByteArrayAsync().Result;

                    string fname = Environment.GetEnvironmentVariable("temp") + "\\" + Guid.NewGuid().ToString()+".png";
                    File.WriteAllBytes(fname, img);
                    return fname;
                }
            }
        }

        public static JObject RetrieveWorkItems(string columns, string workitems)
        {
            string url = String.Format("http://{0}:8080/{1}/_apis/wit/workitems?ids={2}&fields={3}&?api-version=1.0",
    "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
    "tfs/teld",
    workitems,
    columns
    );
           
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", User, Pass))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    return JsonConvert.DeserializeObject(responseBody) as JObject;
                }
            }

        }

        public static string ExecuteQueryBySQL(string sql)
        {
            string url = String.Format("http://{0}:8080/{1}/_apis/wit/wiql?api-version=1.0",
                    "tfs.teld.cn",
                    "tfs/teld"
                    );
            string ret = String.Empty;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", User, Pass))));


                var postvalue = new StringContent("{\"query\":\"" + sql.Replace("\\","\\\\") + "\"}", Encoding.UTF8, "application/json");
                var method = new HttpMethod("POST");
                var request = new HttpRequestMessage(method, url) { Content = postvalue };
                var response = client.SendAsync(request).Result;

                string result = String.Empty;
                //这句不要在if里面加，坑爹啊！拿出来，才能看到真正的程序错误！！！
                //这个傻逼错误：你必须在请求正文中传递有效的修补程序文档
                //对应的English version实际上是：
                //{"$id":"1","innerException":null,"message":"You must pass a valid patch document in the body of the request.","typeName":"Microsoft.VisualStudio.Services.Common.VssPropertyValidationException, Microsoft.VisualStudio.Services.Common, Version=14.0.0.0, Culture=neutral, PublicKeyToken=b03ftoken0a3a","typeKey":"VssPropertyValidationException","errorCode":0,"eventId":3000}
                //ref: http://stackoverflow.com/questions/29607416/vso-api-work-item-patch-giving-400-bad-request
                //因为method是patch，所以上面说的，其实就是你的post的数据不对。这是个狗屁说明，毫无帮助。
                ret = response.Content.ReadAsStringAsync().Result;

                if (false == response.IsSuccessStatusCode) ret = String.Empty;

                return ret;

            }
        }        

        public static string ReplacePrjAndDateFromWIQL(string wiql, Tuple<string, string, string> original)
        {
            string prj = original.Item1;
            string date1 = original.Item2;
            string date2 = original.Item3;

            int pos = wiql.IndexOf(prj);
            pos = wiql.IndexOf("'",pos+ prj.Length);
            int pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{0}'" + wiql.Substring(pos2 + 1);

            pos = wiql.IndexOf(date1);
            pos = wiql.IndexOf("'", pos + date1.Length);
            pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{1}'" + wiql.Substring(pos2 + 1);

            pos = wiql.IndexOf(date2);
            pos = wiql.IndexOf("'", pos + date2.Length);
            pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{2}'" + wiql.Substring(pos2 + 1);

            return wiql;
        }

        public static string ReplacePrjAndDateAndPrjFromWIQL(string wiql, Tuple<string, string, string,string> original)
        {
            string prj = original.Item1;
            string date1 = original.Item2;
            string date2 = original.Item3;
            string prj2 = original.Item4;

            wiql = wiql.Replace("@project", "'{0}'");
            int pos = wiql.IndexOf(prj);
            pos = wiql.IndexOf("'", pos + prj.Length);
            int pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{0}'" + wiql.Substring(pos2 + 1);

            if (date1 != "_FQQ_")
            {
                pos = wiql.IndexOf(date1);
                pos = wiql.IndexOf("'", pos + date1.Length);
                pos2 = wiql.IndexOf("'", pos + 1);
                wiql = wiql.Substring(0, pos) + "'{1}'" + wiql.Substring(pos2 + 1);
            }

            if (date2 != "_FQQ_")
            {
                pos = wiql.IndexOf(date2);
                pos = wiql.IndexOf("'", pos + date2.Length);
                pos2 = wiql.IndexOf("'", pos + 1);
                wiql = wiql.Substring(0, pos) + "'{2}'" + wiql.Substring(pos2 + 1);
            }

            if (prj2 != "_FQQ_")
            {
                pos = wiql.IndexOf(prj2);
                pos = wiql.IndexOf("'", pos + prj2.Length);
                pos2 = wiql.IndexOf("'", pos + 1);
                wiql = wiql.Substring(0, pos) + "'{0}'" + wiql.Substring(pos2 + 1);
            }

            return wiql;
        }

        public static string ReplaceProjectAndIterationFromWIQL(string wiql)
        {
            string prj = "[System.TeamProject] =";
            int pos = wiql.IndexOf(prj);
            pos = wiql.IndexOf("'", pos + prj.Length);
            int pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{0}'" + wiql.Substring(pos2 + 1);

            string ite = "[System.IterationPath] =";
            pos = wiql.IndexOf(ite);
            pos = wiql.IndexOf("'", pos + ite.Length);
            pos2 = wiql.IndexOf("'", pos + 1);
            wiql = wiql.Substring(0, pos) + "'{1}'" + wiql.Substring(pos2 + 1);

            return wiql;
        }
        public static string GetQueryClause(string queryID)
        {

            string url = String.Format("http://{0}:8080/{1}/_apis/wit/queries/{2}?$expand=clauses&api-version=1.0",
                "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
                "tfs/teld/orgportal",
                queryID
                );


            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", User, Pass))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    return Convert.ToString((JsonConvert.DeserializeObject(responseBody) as JObject)["wiql"]);
                }
            }
        }
        
        private static Dictionary<string, string> memberCache = new Dictionary<string, string>();
        private void FlushTeamMemberList()
        {
            memberCache = new Dictionary<string, string>();

            //http://tfs.teld.cn:8080/tfs/teld/_apis/projects/2edbf3dd-f4d4-4ff9-9cd8-cadda5bdc21a/teams/be7ae25c-0fc2-429a-8759-ef84675fc028/members?api-version=2.0

            List<ProjectEntity> prjlist = new List<ProjectEntity>();
            //http://tfs.teld.cn:8080/tfs/teld/_apis/projects?api-version=2.0
            string url = String.Format("http://{0}:8080/{1}/_apis/projects/2edbf3dd-f4d4-4ff9-9cd8-cadda5bdc21a/teams/be7ae25c-0fc2-429a-8759-ef84675fc028/members?api-version=2.0",
                    "tfs.teld.cn",
                    "tfs/teld"
                    );


            string responseBody = GetHttpResponseByUrl(url);

            foreach (var person in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                string dispname = Convert.ToString(person["displayName"]);
                string uniquename = Convert.ToString(person["uniqueName"]);
                memberCache.Add(dispname, String.Format("{0} <{1}>", dispname, uniquename));
            }

        }

        private static List<string> testMembers = new List<string>();

        public static List<string> GetTestMembers(bool forceRefresh)
        {
            if ((false==forceRefresh) && (testMembers.Count > 0)) return testMembers;

            testMembers.Clear();
            var list = API.TFS.TeamProject.Member.RetrieveMemberListByTeam("orgportal", "TestManager");
            foreach (var me in list)
            {
                testMembers.Add(me.FullName);
            }

            return testMembers;
        }
        
        public static List<JToken> ConvertWorkitemFlatQueryResult2Array(string responseBody)
        {
            List<JToken> list = new List<JToken>();
            if (String.IsNullOrEmpty(responseBody)) return list;

            var jsonobj = JsonConvert.DeserializeObject(responseBody) as JObject;
            StringBuilder sbrefname = new StringBuilder();
            StringBuilder sbid = new StringBuilder();

            var columns = jsonobj["columns"] as JArray;
            foreach (var column in columns)
            {
                var refname = Convert.ToString(column["referenceName"]);
                sbrefname.Append(refname).Append(",");
                var txtname = Convert.ToString(column["name"]);
            }

            var wiarray = jsonobj["workItems"] as JArray;
            if (wiarray == null || wiarray.Count < 1) return list;

            foreach (var id in wiarray)
            {
                var wiid = Convert.ToString(id["id"]);
                sbid.Append(wiid).Append(",");
            }

            sbrefname.Remove(sbrefname.Length - 1, 1);
            sbid.Remove(sbid.Length - 1, 1);

        //    string detailsUrl = String.Format("http://{0}:8080/{1}/_apis/wit/workitems?ids={2}&fields={3}&api-version=2.0",
        //"tfs.teld.cn",
        //"tfs/teld",
        //sbid.ToString(),
        //sbrefname.ToString()
        //);

        //    responseBody = Utility.GetHttpResponseByUrl(detailsUrl);

            //var wiarray = (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray;

            List<StringBuilder> sbWorkItems = new List<StringBuilder>();

            //Get work items
            //After executing a query, get the work items using the IDs that are returned in the query results response. 
            //You can get up to 200 work items at a time.
            //https://www.visualstudio.com/en-us/docs/integrate/api/wit/wiql
            for (int i = 0; i < (wiarray.Count - 1) / 200 + 1; i++)
            {
                StringBuilder sbTemp = new StringBuilder();
                int limit = (i < (wiarray.Count - 1) / 200) ? 200 : (wiarray.Count - (wiarray.Count - 1) / 200 * 200);
                for (int j = 0; j < limit; j++)
                {
                    sbTemp.Append(Convert.ToString(wiarray[200 * i + j]["id"])).Append(",");
                }
                if (sbTemp.Length > 0) sbTemp.Remove(sbTemp.Length - 1, 1);

                sbWorkItems.Add(sbTemp);
            }

            int total = sbWorkItems.Count;
            for (int i = 0; i < sbWorkItems.Count; i++)// test only
            {
                var buglist = RetrieveWorkItems(sbrefname.ToString(), sbWorkItems[i].ToString())["value"] as JArray;
                foreach (var bug in buglist)
                {
                    list.Add(bug);
                }
            }

            return list;
        }

        public static void ReleaseComObject(object com)
        {
            while (Marshal.ReleaseComObject(com) > 0) ;
            com = null;
        }

        public static int GetStandardWorkingDays(string prjName, IterationEntity ite)
        {
            int standardWorkingDays;
            var daysoff = Iteration.GetProjectIterationDaysOff(prjName, ite.Id);
            standardWorkingDays = (int)((DateTime.Parse(ite.EndDate) - DateTime.Parse(ite.StartDate)).TotalDays) + 1 - daysoff.Count;
            for (DateTime dt = DateTime.Parse(ite.StartDate); dt < DateTime.Parse(ite.EndDate).AddDays(1); dt = dt.AddDays(1))
            {
                bool duplicated = false;
                for (int i = 0; i < daysoff.Count; i++)
                {
                    if (dt.Equals(DateTime.Parse(daysoff[i])))
                    {
                        duplicated = true;
                        break;
                    }
                }
                if (duplicated) continue;//如果在迭代里面又单独设置了休息日，那么要和取出来的daysoff排除掉。
                if (dt.DayOfWeek == DayOfWeek.Sunday) standardWorkingDays--;//再刨掉礼拜天
            }

            return standardWorkingDays;
        }
        public static IterationEntity GetBestIteration(string project)
        {
            DateTime now = DateTime.Now;
            var itelist = API.TFS.Agile.Iteration.GetProjectIterations(project);

            for (int i = 1; i < itelist.Count; i++)
            {
                if (string.IsNullOrEmpty(itelist[i - 1].EndDate) || string.IsNullOrEmpty(itelist[i - 0].EndDate)) continue;
                DateTime last1 = DateTime.ParseExact(itelist[i - 1].EndDate, "yyyy/M/d h:mm:ss", null).AddDays(1);
                DateTime last2 = DateTime.ParseExact(itelist[i - 0].EndDate, "yyyy/M/d h:mm:ss", null).AddDays(1);
                if (now >= last1 && now <= last2)
                {
                    return itelist[i - 1];
                }
            }

            return null;
        }

    }


}
