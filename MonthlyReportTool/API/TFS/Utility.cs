using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using MonthlyReportTool.Entity;
using MonthlyReportTool.API.TFS.TeamProject;
using MonthlyReportTool.API.TFS.WorkItem;
using System.Runtime.InteropServices;
using MonthlyReportTool.API.TFS.Agile;
using System.IO;

namespace MonthlyReportTool.API.TFS
{
    public class Utility
    {
        private static string username = "juqiang";
        private static string password = "Password02!";

        public static string GetHttpResponseByUrl(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

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
            string url = String.Format("http://tfs.teld.cn:8080/tfs/Teld/{0}/_api/_teamChart/Burndown?chartOptions=%7B%22Width%22%3A494%2C%22Height%22%3A581%2C%22" +
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
                            string.Format("{0}:{1}", username, password))));

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

        private static string GetUnCompletedDepartmantPlan(int days)
        {
            //54e8a414-fef8-4b5d-bec9-57b477eac320
            //https://fabrikam-fiber-inc.visualstudio.com/DefaultCollection/Fabrikam-Fiber-Git/_apis/wit/wiql/1e4e5b17-f212-4ba2-9c2f-a95600ef50a5?api-version=1.0
            string url = "http://tfs.teld.cn:8080/tfs/teld/_apis/wit/wiql/54e8a414-fef8-4b5d-bec9-57b477eac320?api-version=1.0";
            string responseBody = GetHttpResponseByUrl(url);

            StringBuilder fieldlist = new StringBuilder();
            StringBuilder itemlist = new StringBuilder();

            var jsonColumns = (JsonConvert.DeserializeObject(responseBody) as JObject)["columns"] as JArray;
            foreach (var workitem in (JsonConvert.DeserializeObject(responseBody) as JObject)["columns"] as JArray)
            {
                fieldlist.Append((workitem["referenceName"] as JValue).Value.ToString()).Append(",");
            }
            fieldlist.Remove(fieldlist.Length - 1, 1);

            foreach (var workitem in (JsonConvert.DeserializeObject(responseBody) as JObject)["workItemRelations"] as JArray)
            {
                itemlist.Append((workitem["target"]["id"] as JValue).Value.ToString()).Append(",");
            }
            itemlist.Remove(itemlist.Length - 1, 1);

            string url2 = "http://tfs.teld.cn:8080/tfs/teld/_apis/wit/WorkItems?ids="
                + itemlist.ToString()
                + "&fields="
                + fieldlist.ToString()
                + "&api-version=1.0";
            responseBody = GetHttpResponseByUrl(url2);

            return responseBody;
        }

        private List<string> RetrieveFilesOfCheckInHistory(string startTime, string endTime)
        {
            List<string> filelist = new List<string>();
            string url = String.Format("http://tfs.teld.cn:8080/tfs/teld/_apis/tfvc/changesets?fromDate={0}&toDate={1}&api-version=1.0",
                    startTime,
                    endTime
                    );

            List<string> changesets = new List<string>();

            string responseBody = GetHttpResponseByUrl(url);
            foreach (var changeset in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
            {
                changesets.Add(Convert.ToString(changeset["changesetId"]));
            }

            foreach (string changesetsid in changesets)
            {
                string csurl = String.Format("http://tfs.teld.cn:8080/tfs/teld/_apis/tfvc/changesets/{0}/changes?api-version=1.0",
                    changesetsid);
                string csresponse = GetHttpResponseByUrl(csurl);


                foreach (var changeset in (JsonConvert.DeserializeObject(csresponse) as JObject)["value"] as JArray)
                {
                    filelist.Add(Convert.ToString(changeset["item"]["url"]).Replace("t-bj-tfs-01", "tfs.teld.cn"));
                }
            }

            return filelist;
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
                            string.Format("{0}:{1}", username, password))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    return JsonConvert.DeserializeObject(responseBody) as JObject;
                }
            }

        }
        public static JObject ExecuteQuery(string queryID)
        {

            string url = String.Format("http://{0}:8080/{1}/_apis/wit/wiql/{2}?api-version=1.0",
                "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
                "tfs/teld",
                queryID
                );


            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

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
                            string.Format("{0}:{1}", username, password))));


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
                            string.Format("{0}:{1}", username, password))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    return Convert.ToString((JsonConvert.DeserializeObject(responseBody) as JObject)["wiql"]);
                }
            }
        }
        private static void GetDashboards()
        {
            string url = "http://tfs.teld.cn:8080/tfs/teld/c955f4f8-3b05-4afc-9969-3a54f7b70533/ad0bf755-9688-4d6a-95f9-4e20702a2972/_apis/dashboard/dashboards?api-version=3.0-preview.2";
            var username = "juqiang";
            var password = "Password02!";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    foreach (var person in (JsonConvert.DeserializeObject(responseBody) as JObject)["value"] as JArray)
                    {
                        string dispname = Convert.ToString(person["displayName"]);
                        string uniquename = Convert.ToString(person["uniqueName"]);
                        memberCache.Add(dispname, String.Format("{0} <{1}>", dispname, uniquename));
                    }
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
            var list = API.TFS.TeamProject.Member.RetrieveMemberList("orgportal", "TestManager");
            foreach (var me in list)
            {
                testMembers.Add(me.FullName);
            }

            return testMembers;
        }
        public static void RetrieveTeamMemberList(string prj)
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

        private string GetTFSAccountByName(string name)
        {
            if (memberCache.ContainsKey(name)) return memberCache[name];
            else
            {
                FlushTeamMemberList();
                if (memberCache.ContainsKey(name)) return memberCache[name];
                else throw new Exception("No such account in TFS.");
            }
        }

        public bool CreateBug(string username, string password, string title, string assignedTo, string warningTime = "", string warningHost = "",
            string warningRule = "", string warningCategory = "", string warningLevel = "",
            string warningLimit = "", string warningActualValue = "", string detailDataLink = "", string detailGraphLink = "")
        {

            //如下是官方原始文档
            //https://www.visualstudio.com/en-us/docs/integrate/api/wit/work-items#samples
            //如下是坑
            //http://blog.aitgmbh.de/2014/08/26/how-to-change-work-item-state-using-visual-studio-online-rest-api/
            //如下是讨论，核心大意是：TFS高度灵活，很多参数，和你配置的template有关系。
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/1eeef675-1338-41b2-b4ba-b8854d6c8f82/rest-api-what-are-the-work-item-creation-update-requirements?forum=vsx
            //如下是解决我问题的例子，用fiddler抓包，仔细比对每一个他贴的内容
            //http://stackoverflow.com/questions/29607416/vso-api-work-item-patch-giving-400-bad-request

            bool succ = false;

            string bugurl = String.Format("http://{0}:8080/{1}/_apis/wit/workitems/$Bug?api-version=2.0",
                    "tfs.teld.cn",
                    "tfs/teld/bugs"
                    );

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

                Object[] patchDocument = new Object[9];

                patchDocument[0] = new { op = "add", path = "/fields/System.Title", value = title };

                patchDocument[1] = new { op = "add", path = "/fields/System.AssignedTo", value = GetTFSAccountByName(assignedTo) };
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("<p>告警时间：{0}</p>", warningTime);
                sb.AppendFormat("<p>告警主机：{0}</p>", warningHost);
                sb.AppendFormat("<p>告警维度：{0}</p>", warningCategory);
                sb.AppendFormat("<p>告警级别：{0}</p>", warningLevel);
                sb.AppendFormat("<p>告警规则：{0}</p>", warningRule);
                sb.AppendFormat("<p>告警阈值：{0}</p>", warningLimit);
                sb.AppendFormat("<p>告警实际值：{0}</p>", warningActualValue);
                sb.AppendFormat("<p>OM监控地址：{0}</p>", detailDataLink);
                sb.AppendFormat("<p>Grafana监控地址：{0}</p>", detailGraphLink);
                patchDocument[2] = new { op = "add", path = "/fields/Microsoft.VSTS.TCM.ReproSteps", value = sb.ToString() };
                patchDocument[3] = new { op = "add", path = "/fields/Teld.Bug.Client", value = "WebPC" };
                patchDocument[4] = new { op = "add", path = "/fields/Teld.Bug.FunctionMenu", value = "这里写你的出错的功能" };
                //这里都放充电里面吧！
                patchDocument[5] = new { op = "add", path = "/fields/System.AreaPath", value = "Bugs\\CCP_充电云平台部" };
                patchDocument[6] = new { op = "add", path = "/fields/System.State", value = "新建" };
                patchDocument[7] = new { op = "add", path = "/fields/Microsoft.VSTS.Common.Priority", value = "1" };
                //如下严重之类的不能乱写，必须和template定义的一致
                patchDocument[8] = new { op = "add", path = "/fields/Microsoft.VSTS.Common.Severity", value = "1 - 严重" };
                //patchDocument[9] = new { op = "add", path = "/fields/System.Reason", value = "fuck reason" };

                var patchValue = new StringContent(JsonConvert.SerializeObject(patchDocument), Encoding.UTF8, "application/json-patch+json");
                //.GetEncoding("gb2312")
                var method = new HttpMethod("PATCH");
                var request = new HttpRequestMessage(method, bugurl) { Content = patchValue };
                var response = client.SendAsync(request).Result;

                string result = String.Empty;
                //这句不要在if里面加，坑爹啊！拿出来，才能看到真正的程序错误！！！
                //这个傻逼错误：你必须在请求正文中传递有效的修补程序文档
                //对应的English version实际上是：
                //{"$id":"1","innerException":null,"message":"You must pass a valid patch document in the body of the request.","typeName":"Microsoft.VisualStudio.Services.Common.VssPropertyValidationException, Microsoft.VisualStudio.Services.Common, Version=14.0.0.0, Culture=neutral, PublicKeyToken=b03ftoken0a3a","typeKey":"VssPropertyValidationException","errorCode":0,"eventId":3000}
                //ref: http://stackoverflow.com/questions/29607416/vso-api-work-item-patch-giving-400-bad-request
                //因为method是patch，所以上面说的，其实就是你的post的数据不对。这是个狗屁说明，毫无帮助。
                result = response.Content.ReadAsStringAsync().Result;

                if (response.IsSuccessStatusCode)
                {
                    succ = true;
                    result = response.Content.ReadAsStringAsync().Result;
                }

            }

            return succ;
        }
        public bool CreateBugTest()
        {
            //如下是官方原始文档
            //https://www.visualstudio.com/en-us/docs/integrate/api/wit/work-items#samples
            //如下是坑
            //http://blog.aitgmbh.de/2014/08/26/how-to-change-work-item-state-using-visual-studio-online-rest-api/
            //如下是讨论，核心大意是：TFS高度灵活，很多参数，和你配置的template有关系。
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/1eeef675-1338-41b2-b4ba-b8854d6c8f82/rest-api-what-are-the-work-item-creation-update-requirements?forum=vsx
            //如下是解决我问题的例子，用fiddler抓包，仔细比对每一个他贴的内容
            //http://stackoverflow.com/questions/29607416/vso-api-work-item-patch-giving-400-bad-request

            bool succ = false;

            string bugurl = String.Format("http://{0}:8080/{1}/_apis/wit/workitems/$Bug?api-version=2.0",
                    "tfs.teld.cn",
                    "tfs/teld/bugs"
                    );

            var username = "juqiang";
            var password = "yourpassword";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

                Object[] patchDocument = new Object[9];

                patchDocument[0] = new { op = "add", path = "/fields/System.Title", value = "鞠强测试bug001" };
                //TODO：这个名字要取出来
                patchDocument[1] = new { op = "add", path = "/fields/System.AssignedTo", value = "任丽霞 <TELD\\renlx>" };
                patchDocument[2] = new { op = "add", path = "/fields/Microsoft.VSTS.TCM.ReproSteps", value = "1.在预警代码的发送消息部分，调用此方法。2.修改适当的参数即可。" };
                patchDocument[3] = new { op = "add", path = "/fields/Teld.Bug.Client", value = "WebPC" };
                patchDocument[4] = new { op = "add", path = "/fields/Teld.Bug.FunctionMenu", value = "这里写你的出错的功能" };
                //这里都放充电里面吧！
                patchDocument[5] = new { op = "add", path = "/fields/System.AreaPath", value = "Bugs\\CCP_充电云平台部" };
                patchDocument[6] = new { op = "add", path = "/fields/System.State", value = "新建" };
                patchDocument[7] = new { op = "add", path = "/fields/Microsoft.VSTS.Common.Priority", value = "1" };
                //如下严重之类的不能乱写，必须和template定义的一致
                patchDocument[8] = new { op = "add", path = "/fields/Microsoft.VSTS.Common.Severity", value = "1 - 严重" };
                //patchDocument[9] = new { op = "add", path = "/fields/System.Reason", value = "fuck reason" };

                var patchValue = new StringContent(JsonConvert.SerializeObject(patchDocument), Encoding.UTF8, "application/json-patch+json");
                //.GetEncoding("gb2312")
                var method = new HttpMethod("PATCH");
                var request = new HttpRequestMessage(method, bugurl) { Content = patchValue };
                var response = client.SendAsync(request).Result;

                string result = String.Empty;
                //这句不要在if里面加，坑爹啊！拿出来，才能看到真正的程序错误！！！
                //这个傻逼错误：你必须在请求正文中传递有效的修补程序文档
                //对应的English version实际上是：
                //{"$id":"1","innerException":null,"message":"You must pass a valid patch document in the body of the request.","typeName":"Microsoft.VisualStudio.Services.Common.VssPropertyValidationException, Microsoft.VisualStudio.Services.Common, Version=14.0.0.0, Culture=neutral, PublicKeyToken=b03ftoken0a3a","typeKey":"VssPropertyValidationException","errorCode":0,"eventId":3000}
                //ref: http://stackoverflow.com/questions/29607416/vso-api-work-item-patch-giving-400-bad-request
                //因为method是patch，所以上面说的，其实就是你的post的数据不对。这是个狗屁说明，毫无帮助。
                result = response.Content.ReadAsStringAsync().Result;

                if (response.IsSuccessStatusCode)
                {
                    succ = true;
                    result = response.Content.ReadAsStringAsync().Result;
                }

            }

            return succ;
        }


        private List<ReleasePlan> RetrieveReleasePlanEntityList()
        {
            List<ReleasePlan> list = new List<ReleasePlan>();

            string url = String.Format("http://{0}:8080/{1}/_apis/wit/wiql/{2}?api-version=2.0",
                "tfs.teld.cn", //t-bj-tfs.chinacloudapp.cn
                "tfs/teld",
                "773D53A1-5240-47D2-9E98-3DE5F5247776"
                );

            var username = "juqiang";
            var password = "yourpassword";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(
                    new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(
                        System.Text.ASCIIEncoding.ASCII.GetBytes(
                            string.Format("{0}:{1}", username, password))));

                string detailsUrl = String.Empty;
                using (HttpResponseMessage response = client.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

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

                    var ids = jsonobj["workItems"] as JArray;
                    foreach (var id in ids)
                    {
                        var wiid = Convert.ToString(id["id"]);
                        sbid.Append(wiid).Append(",");
                    }

                    sbrefname.Remove(sbrefname.Length - 1, 1);
                    sbid.Remove(sbid.Length - 1, 1);

                    detailsUrl = String.Format("http://{0}:8080/{1}/_apis/wit/workitems?ids={2}&fields={3}&api-version=2.0",
                "tfs.teld.cn",
                "tfs/teld",
                sbid.ToString(),
                sbrefname.ToString()
                );
                }

                using (HttpResponseMessage response = client.GetAsync(detailsUrl).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;
                    var jsonobj = JsonConvert.DeserializeObject(responseBody) as JObject;
                    var items = jsonobj["value"] as JArray;
                    foreach (var item in items)
                    {
                        var fields = (item["fields"] as JObject);

                        ReleasePlan rpe = new ReleasePlan()
                        {
                            Id = Convert.ToInt32(fields["System.Id"]),
                            TeamProject = Convert.ToString(fields["System.TeamProject"]),
                            WorkItemType = Convert.ToString(fields["System.WorkItemType"]),
                            State = Convert.ToString(fields["System.State"]),
                            AssignedTo = Convert.ToString(fields["System.AssignedTo"]),
                            Title = Convert.ToString(fields["System.Title"]),
                            ReleaseDate = Convert.ToDateTime(fields["Teld.Release.Date"]).AddHours(8).ToString("yyyy-MM-dd HH:mm:ss"),
                            RollbackCount = Convert.ToInt32(fields["Teld.Publish.RollbackCount"]),
                            PublishType = Convert.ToString(fields["Teld.Publish.Type"]),
                            RollbackReason = Convert.ToString(fields["Teld.Publish.RollbackReason"])
                        };

                        int pos = rpe.AssignedTo.IndexOf("<");
                        if (pos > -1)
                        {
                            rpe.AssignedTo = rpe.AssignedTo.Substring(0, pos - 1).Trim();
                        }
                        list.Add(rpe);
                    }
                }
            }

            return list;
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
            if (wiarray.Count < 1) return list;

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
