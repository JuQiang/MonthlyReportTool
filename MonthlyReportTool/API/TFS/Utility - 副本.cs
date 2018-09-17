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
using System.Collections;
using System.Net;

namespace MonthlyReportTool.API.TFS
{
    public class Utility
    {
        public static string User = "";
        public static string Pass = "";
        public static string QueryBaseDirectory = "共享查询%2F迭代总结数据查询%2F{0}";

        public static string GetBurndownPictureFile(string projectName)
        {
            string url = String.Format("http://tfs.teld.cn:8080/tfs/Teld/{0}/_api/_teamChart/Burndown?chartOptions=%7B%22Width%22%3A1248%2C%22Height%22%3A616%2C%22" +
                "ShowDetails%22%3Atrue%2C%22Title%22%3A%22%22%7D&counter=2&iterationPath={1}&__v=5",
                    projectName,
                    Utility.GetBestIteration(projectName).Path
                    );
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.ContentType = "application/json";
            //request.ContentType = "application/x-www-form-urlencoded";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.Method = "GET";
            request.Timeout = 30000;
            request.Credentials = new NetworkCredential(User, Pass); //credential;

            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                using (BinaryReader reader = new BinaryReader(response.GetResponseStream()))
                {
                    byte[] img = reader.ReadBytes(100000);
                    string fname = Environment.GetEnvironmentVariable("temp") + "\\" + Guid.NewGuid().ToString() + ".png";
                    File.WriteAllBytes(fname, img);
                    return fname;
                }
            }
        }

        private static List<string> testMembers = new List<string>();
        public static List<string> GetTestMembers(bool forceRefresh)
        {
            if ((false == forceRefresh) && (testMembers.Count > 0)) return testMembers;

            testMembers.Clear();
            var list = API.TFS.TeamProject.Member.RetrieveMemberListByTeam("orgportal", "TestManager");
            foreach (var me in list)
            {
                testMembers.Add(me.FullName);
            }

            return testMembers;
        }

        //替换查询变量，循环替换，对于多个的也都替换了，并且排除不需要替换的。
        public static string ReplaceInformationFromWIQLByReplaceList(string wiql, List<WiqlReplaceColumnEntity> colvalues)
        {
            int pos = -1, pos2 = -1;
            WiqlReplaceColumnEntity colvalue;
            for (int i = 0; i < colvalues.Count; i++)
            {
                colvalue = colvalues[i];
                //1、先搜索包含column的内容，如果搜索到，再搜索包含排除的内容，如果两者的index一样，则跳过继续下一次搜索
                //2、如果是顺序的话，则找到一个就停止，如果是循环，则继续ｗｈｉｌｅ执行
                //pos = wiql.IndexOf(colvalue.column);
                //pos = wiql.IndexOf("'", pos + colvalue.column.Length);
                //pos2 = wiql.IndexOf("'", pos + 1);
                //wiql = wiql.Substring(0, pos) + "'" + colvalue.replacevalue + "'" + wiql.Substring(pos2 + 1);
                pos = wiql.IndexOf(colvalue.column);
                while (pos >= 0)
                {
                    if (!String.IsNullOrEmpty(colvalue.notinclude))//不为空,判断是否是要排除的,是就继续
                    {
                        int tmppos = wiql.IndexOf(colvalue.notinclude, pos2 + 1);
                        if (tmppos == pos)
                        {
                            pos2 = pos2 + colvalue.notinclude.Length;
                            pos = wiql.IndexOf(colvalue.column, pos2 + 1);
                            continue;
                        }
                    }
                    pos = wiql.IndexOf("'", pos + colvalue.column.Length);
                    pos2 = wiql.IndexOf("'", pos + 1);
                    wiql = wiql.Substring(0, pos) + "'" + colvalue.replacevalue + "'" + wiql.Substring(pos2 + 1);
                    pos = wiql.IndexOf(colvalue.column, pos2 + 1);
                }
            }

            return wiql;
        }

        public static string GetQueryClause(string queryID)
        {
            string url = String.Format("http://{0}:8080/{1}/_apis/wit/queries/{2}?$expand=clauses&api-version=4.1",
                "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
                "tfs/teld/orgportal",
                queryID
                );
            string responseBody = GetHttpResponseByUrl(url);
            return Convert.ToString((JsonConvert.DeserializeObject(responseBody) as JObject)["wiql"]);
        }
        public static string GetHttpResponseByUrl(string url, string queryParameter="", string requestMethod="GET")
        {
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.ContentType = "application/json";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.Method = requestMethod;// "POST","GET";
            request.Timeout = 30000;
            request.Credentials = new NetworkCredential(User, Pass); //credential;

            if (requestMethod == "POST")
            {
                string data = JsonConvert.SerializeObject(queryParameter);
                //转换输入参数的编码类型，获取bytep[]数组 
                byte[] byteArray = Encoding.UTF8.GetBytes(data);
                request.ContentLength = byteArray.Length;
                Stream newStream = request.GetRequestStream();//创建一个Stream,赋值是写入HttpWebRequest对象提供的一个stream里面
                newStream.Write(byteArray, 0, byteArray.Length);
            }

            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string responseMsg = reader.ReadToEnd();
                    return responseMsg;
                }
            }
        }
        public static string ExecuteQueryBySQL(string sql)
        {
            string url = String.Format("http://{0}:8080/{1}/_apis/wit/wiql?api-version=4.1",
                    "tfs.teld.cn",
                    "tfs/teld"
                    );

            string data = JsonConvert.SerializeObject(new Dictionary<string, string>() { { "query", sql.Replace("\\", "\\\\") } });

            return GetHttpResponseByUrl(url, data, "POST");
        }
        public static JObject RetrieveWorkItems(string columns, string workitems)
        {
            string url = String.Format("http://{0}:8080/{1}/_apis/wit/workitems?ids={2}&fields={3}&?api-version=4.1",
                                        "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
                                        "tfs/teld",
                                        workitems,
                                        columns
                                     );

            String responseBody = GetHttpResponseByUrl(url);
            return JsonConvert.DeserializeObject(responseBody) as JObject;
        }

        public static List<JToken> ConvertWorkitemQueryResult2Array(string responseBody, ref Hashtable hs)
        {
            List<JToken> list = new List<JToken>();
            if (String.IsNullOrEmpty(responseBody)) return list;
            var jsonobj = JsonConvert.DeserializeObject(responseBody) as JObject;
            var queryType = Convert.ToString(jsonobj["queryType"]);
            if ((string.Equals(queryType, "tree") == true) || (String.Equals(queryType, "oneHop") == true))
                return ConvertWorkitemTreeQueryResult2Array(responseBody, ref hs);
            else
                return ConvertWorkitemFlatQueryResult2Array(responseBody);
        }
        private static List<JToken> ConvertWorkitemFlatQueryResult2Array(string responseBody)
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

            //    string detailsUrl = String.Format("http://{0}:8080/{1}/_apis/wit/workitems?ids={2}&fields={3}&api-version=4.1",
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
        private static List<JToken> ConvertWorkitemTreeQueryResult2Array(string responseBody, ref Hashtable hs)
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

            var wiarray = jsonobj["workItemRelations"] as JArray;
            if (wiarray == null || wiarray.Count < 1) return list;
            foreach (var id in wiarray)
            {
                var wiid = Convert.ToString(id["target"]["id"]);
                sbid.Append(wiid).Append(",");
                if (id["source"] != null && Convert.ToString(id["source"]) != "")
                {
                    if (hs.ContainsKey(wiid)) continue;
                    hs.Add(wiid, Convert.ToString(id["source"]["id"]));
                }
                else
                    hs.Add(wiid, "");
            }

            sbrefname.Remove(sbrefname.Length - 1, 1);
            sbid.Remove(sbid.Length - 1, 1);

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
                    sbTemp.Append(Convert.ToString(wiarray[200 * i + j]["target"]["id"])).Append(",");
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
