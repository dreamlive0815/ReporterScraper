using System;
using System.Text;
using System.Collections.Generic;
using System.Net;
using System.Web;
using Util;
using Newtonsoft.Json.Linq;

namespace Scraper
{
    class Reporter
    {
        private HttpClient client;
        private CookieContainer cookies;
        private string relativeUri = "/WebReport/ReportServer";

        private Dictionary<string, string> moduleMap;
        private double _guiqinJingdu = 120.360226;
        private double _guiqinWeidu = 30.316166;

        private string device_id = "140AD8A4B7CCC4B4DWH9X17115W08684";
        private string device_name = "HUAWEI+DIG-AL00";

        public string DeviceID
        {
            get { return device_id; }
        }

        public string DeviceName
        {
            get { return device_name; }
        }

        public void SetDevice(string name, string id)
        {
            device_name = name;
            device_id = id;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="minimum"></param>
        /// <param name="maximum"></param>
        /// <param name="Len">小数点保留位数</param>
        /// <returns></returns>
        public double GetRandomDouble(double minimum, double maximum, int Len)   
        {
            Random random = new Random();
            return Math.Round(random.NextDouble() * (maximum - minimum) + minimum, Len);
        }

        public double GuiqinJingdu
        {
            get { return _guiqinJingdu; }
            set { _guiqinJingdu = value; }
        }

        public double GuiqinWeidu
        {
            get { return _guiqinWeidu; }
            set { _guiqinWeidu = value; }
        }

        public double KeqianJingdu
        {
            get { return GetRandomDouble(120.340, 120.350, 6); }
        }

        public double KeqianWeidu
        {
            get { return GetRandomDouble(30.315, 30.316, 6); }
        }

        public string GuiqinTime
        {
            get { return "21:31:00"; }
            //get { return DateTime.Now.ToString("HH:mm:ss"); }
        }

        protected string SessionID { get; set; }

        public Dictionary<string, string> GetBaseParams()
        {
            var ps = new Dictionary<string, string>() {
                {"__device__", "android"},
                {"__mobileapp__", "yes"},
                {"isMobile", "yes"},
            };
            if (SessionID != null) ps["sessionID"] = SessionID;
            return ps;
        }

        public Reporter()
        {
            cookies = new CookieContainer();
            client = new HttpClient(new Dictionary<string, object>() {
                {"BaseUri", "http://stu.xinxi.zstu.edu.cn"},
                {"CookieContainer", cookies},
                {"ContentType", "application/x-www-form-urlencoded"},
                {"UserAgent", null}
            });
            client.OnGetRequest += (request) => {
                request.Headers.Add("Cookie2", "$Version=1");
            };
            client.OnGetResponseContent += (content) => {
                //Console.WriteLine(content);
            };
            client.ReadEncoding = Encoding.GetEncoding("gbk");
            moduleMap = new Dictionary<string, string>();
        }

        public HttpClient Client
        {
            get { return client; }
        }

        public CookieContainer Cookies
        {
            get { return cookies; }
        }

        protected JToken GetJson(string s)
        {
            var json = JToken.Parse(s);
            if (json.Type == JTokenType.Array) return json;
            var errCode = (int?)json["errorCode"];
            var errMsg = (string)json["errorMsg"] ?? "";
            if (errCode != null) throw new ReporterException((int)errCode, errMsg);
            return json;
        }

        public void Login(string username, string password)
        {
            var formParams = GetBaseParams();
            formParams["cmd"] = "login";
            formParams["devname"] = DeviceName;
            formParams["fr_password"] = password;
            formParams["fr_remember"] = "false";
            formParams["fr_username"] = username;
            formParams["fspassword"] = password;
            formParams["fsremember"] = "false";
            formParams["fsusername"] = username;
            formParams["macaddress"] = DeviceID;
            formParams["op"] = "fs_mobile_main";
            var res = client.POST(relativeUri, formParams);
            var json = GetJson(res);
        }

        public JToken GetIndexPage()
        {
            var formParams = GetBaseParams();
            formParams["cmd"] = "module_getrootreports";
            formParams["id"] = "-1";
            formParams["op"] = "fs_main";
            var json = GetJson(client.POST(relativeUri, formParams));
            GetModules(json, "");
            return json;
        }

        public void CheckGuiqinLocation(double jingdu, double weidu)
        {
            if (jingdu < 120.345 || jingdu > 120.368) throw new Exception("归寝签到经度范围120.345-120.368");
            if (weidu < 30.308 || weidu > 30.330) throw new Exception("归寝签到纬度范围30.308-30.330");
        }

        public void CheckKeqianLocation(double jingdu, double weidu)
        {
            if (jingdu < 120.331 || jingdu > 120.3582) throw new Exception("课前签到经度范围120.331-120.3582");
            if (weidu < 30.309 || weidu > 30.32) throw new Exception("课前签到纬度范围30.309-30.32");
        }

        protected void AssertWriteCellSucceed(string res)
        {
        }

        public void GuiqinReport()
        {
            var moduleId = GetModuleID(".归寝签到.归寝签到");
            if (moduleId == null) throw new Exception("无法获取归寝签到的操作id");
            var formParams = GetBaseParams();
            formParams["cmd"] = "entry_report";
            formParams["op"] = "fs_main";
            formParams["id"] = moduleId;
            var r = client.POST(relativeUri, formParams);
            var json = GetJson(r);
            var sessionId = (string)json["sessionid"];
            if (sessionId == null) throw new Exception("无法获取SessionID");
            SessionID = sessionId;

            try
            {
                formParams = GetBaseParams();
                formParams["cmd"] = "read_by_json";
                formParams["op"] = "fr_write";
                formParams["pn"] = "1";
                r = client.POST(relativeUri, formParams);

                double jingdu = GuiqinJingdu, weidu = GuiqinWeidu;
                CheckGuiqinLocation(jingdu, weidu);

                formParams = GetBaseParams();
                formParams["cmd"] = "cal_write_cell";
                formParams["editReportIndex"] = "0";
                formParams["editcol"] = "2";
                formParams["editrow"] = "3";
                formParams["loadidxs"] = "[5b]0[5d]";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"3\"><O t=\"D\"><![5b]CDATA[5b]{0}[5d][5d]></O></C></CellElementList></Report></WorkBook>", weidu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editrow"] = "2";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"2\"><O t=\"D\"><![5b]CDATA[5b]{0}[5d][5d]></O></C></CellElementList></Report></WorkBook>", jingdu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editcol"] = "5";
                formParams["editrow"] = "3";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"5\" r=\"3\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editcol"] = "5";
                formParams["editrow"] = "2";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"5\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams.Remove("editcol");
                formParams.Remove("editrow");
                formParams["editcells"] = "[5b]{\"row\":2,\"col\":2},{\"row\":3,\"col\":2},{\"row\":3,\"col\":5},{\"row\":2,\"col\":5}[5d]";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b]{0}[5d][5d]></O></C><C c=\"2\" r=\"3\"><O t=\"S\"><![5b]CDATA[5b]{1}[5d][5d]></O></C><C c=\"5\" r=\"3\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C><C c=\"5\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>", jingdu, weidu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                CloseSession();

                var time = GuiqinTime;
                formParams = new Dictionary<string, string>();
                formParams["timetype"] = "1";
                formParams["title"] = "%E6%9A%91%E6%9C%9F%E7%AD%BE%E5%88%B0";
                formParams["time"] = HttpUtility.UrlEncode(time);
                formParams["jingdu"] = jingdu.ToString();
                formParams["weidu"] = weidu.ToString();
                formParams["reportlet"] = "2017%2Fbaodaocheck_enter.cpt";
                formParams["op"] = "write";
                formParams["__replaceview__"] = "true";
                var uri = string.Format("{0}?{1}", relativeUri, HttpClient.BuildQuery(formParams));
                formParams = GetBaseParams();
                r = client.POST(uri, formParams);
                json = GetJson(r);
                sessionId = (string)json["sessionid"];
                if (sessionId == null) throw new Exception("无法获取SessionID");
                SessionID = sessionId;

                formParams = GetBaseParams();
                formParams["cmd"] = "read_by_json";
                formParams["op"] = "fr_write";
                formParams["pn"] = "1";
                r = client.POST(relativeUri, formParams);

                formParams = GetBaseParams();
                formParams["cmd"] = "write_verify";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                json = GetJson(r);
                var msg = (string)json[0]?["message"];
                if(msg != null) throw new Exception(string.Format("归寝签到验证失败:{0}", msg));

                formParams = GetBaseParams();
                formParams["cmd"] = "submit_w_report";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                json = GetJson(r);
                var info = json[0]?["fr_submitinfo"];
                if (info == null || !(bool)info["success"])
                {
                    Console.WriteLine(r);
                    throw new Exception("归寝签到失败");
                }

                Console.WriteLine("归寝签到成功");
            }
            catch (Exception ex)
            {
                CloseSession();
                throw ex;
            }

            CloseSession();
        }

        public static string GetKeqianTimeByTimetype(string timetype)
        {
            int hour = 0, minute = 0;
            switch (timetype)
            {
                case "A1": hour = 8; minute = 0; break;
                case "A2": hour = 9; minute = 50; break;
                case "A3": hour = 9; minute = 50; break;
                case "A4": hour = 13; minute = 20; break;
                case "A5": hour = 13; minute = 20; break;
                case "A6": hour = 15; minute = 05; break;
                case "A7": hour = 18; minute = 20; break;
                case "A8": hour = 18; minute = 20; break;
                default: throw new Exception("未知的时间段");
            }
            var now = DateTime.Now;
            var time = new DateTime(now.Year, now.Month, now.Day, hour, minute, 0);
            var rand = new Random();
            var offset = rand.Next(600);
            time = time.AddSeconds(offset);
            return time.ToString("HH:mm:ss");
        }

        public void KeqianReport(string timetype)
        {
            var moduleId = GetModuleID(".考勤体系.课前签到【测试】");
            if (moduleId == null) throw new Exception("无法获取课前签到的操作id");
            var formParams = GetBaseParams();
            formParams["cmd"] = "entry_report";
            formParams["op"] = "fs_main";
            formParams["id"] = moduleId;
            var r = client.POST(relativeUri, formParams);
            var json = GetJson(r);
            var sessionId = (string)json["sessionid"];
            if (sessionId == null) throw new Exception("无法获取SessionID");
            SessionID = sessionId;

            try
            {
                formParams = GetBaseParams();
                formParams["cmd"] = "read_by_json";
                formParams["op"] = "fr_write";
                formParams["pn"] = "1";
                r = client.POST(relativeUri, formParams);

                formParams = GetBaseParams();
                formParams["__pi__"] = "true";
                formParams["op"] = "write";
                formParams["reportlet"] = string.Format("/tiaoshi/{0}.cpt", timetype);
                r = client.POST(relativeUri, formParams);
                json = GetJson(r);
                sessionId = (string)json["sessionid"];
                if (sessionId == null) throw new Exception("无法获取SessionID");
                SessionID = sessionId;

                formParams = GetBaseParams();
                formParams["cmd"] = "read_by_json";
                formParams["op"] = "fr_write";
                formParams["pn"] = "1";
                r = client.POST(relativeUri, formParams);

                double jingdu = KeqianJingdu, weidu = KeqianWeidu;
                CheckKeqianLocation(jingdu, weidu);

                formParams = GetBaseParams();
                formParams["cmd"] = "cal_write_cell";
                formParams["editReportIndex"] = "0";
                formParams["editcol"] = "2";
                formParams["editrow"] = "1";
                formParams["loadidxs"] = "[5b]0[5d]";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"1\"><O t=\"D\"><![5b]CDATA[5b]{0}[5d][5d]></O></C></CellElementList></Report></WorkBook>", jingdu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editcol"] = "5";
                formParams["editrow"] = "2";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"5\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editcol"] = "2";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"2\"><O t=\"D\"><![5b]CDATA[5b]{0}[5d][5d]></O></C></CellElementList></Report></WorkBook>", weidu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams["editcol"] = "5";
                formParams["editrow"] = "1";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"5\" r=\"1\"><O t=\"S\"><![5b]CDATA[5b][65e0][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                formParams.Remove("editcol");
                formParams.Remove("editrow");
                formParams["editcells"] = "[5b]{\"row\":1,\"col\":2},{\"row\":2,\"col\":2},{\"row\":2,\"col\":5},{\"row\":1,\"col\":5}[5d]";
                formParams["reportXML"] = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList><C c=\"2\" r=\"1\"><O t=\"S\"><![5b]CDATA[5b]{0}[5d][5d]></O></C><C c=\"2\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b]{1}[5d][5d]></O></C><C c=\"5\" r=\"2\"><O t=\"S\"><![5b]CDATA[5b][6709][6548][5d][5d]></O></C><C c=\"5\" r=\"1\"><O t=\"S\"><![5b]CDATA[5b][65e0][6548][5d][5d]></O></C></CellElementList></Report></WorkBook>", jingdu, weidu);
                r = client.POST(relativeUri, formParams);
                AssertWriteCellSucceed(r);

                CloseSession();

                var time = GetKeqianTimeByTimetype(timetype);
                formParams = new Dictionary<string, string>();
                formParams["timetype"] = "0";
                formParams["title"] = "%E8%80%83%E5%8B%A4%E7%AD%BE%E5%88%B0";
                formParams["time"] = HttpUtility.UrlEncode(time);
                formParams["jingdu"] = jingdu.ToString();
                formParams["weidu"] = weidu.ToString();
                formParams["reportlet"] = string.Format("tiaoshi%2F{0}check_enter.cpt", timetype);
                formParams["op"] = "write";
                formParams["__replaceview__"] = "true";
                var uri = string.Format("{0}?{1}", relativeUri, HttpClient.BuildQuery(formParams));
                formParams = GetBaseParams();
                r = client.POST(uri, formParams);
                json = GetJson(r);
                sessionId = (string)json["sessionid"];
                if (sessionId == null) throw new Exception("无法获取SessionID");
                SessionID = sessionId;

                formParams = GetBaseParams();
                formParams["cmd"] = "read_by_json";
                formParams["op"] = "fr_write";
                formParams["pn"] = "1";
                r = client.POST(relativeUri, formParams);

                formParams = GetBaseParams();
                formParams["cmd"] = "write_verify";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                json = GetJson(r);
                var msg = (string)json[0]?["message"];
                if (msg != null) throw new Exception(string.Format("课前签到验证失败:{0}", msg));

                formParams = GetBaseParams();
                formParams["cmd"] = "submit_w_report";
                formParams["op"] = "fr_write";
                formParams["reportXML"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><WorkBook><Version>6.5</Version><Report class=\"com.fr.report.WorkSheet\" name=\"0\"><CellElementList></CellElementList></Report></WorkBook>";
                r = client.POST(relativeUri, formParams);
                json = GetJson(r);
                var info = json[0]?["fr_submitinfo"];
                if (info == null || !(bool)info["success"])
                {
                    Console.WriteLine(r);
                    throw new Exception("课前签到失败");
                }

                Console.WriteLine("课前签到成功");
            }
            catch(Exception ex)
            {
                CloseSession();
                throw ex;
            }

            CloseSession();
            
        }

        public void CloseSession()
        {
            var formParams = GetBaseParams();
            formParams["op"] = "closesessionid";
            var uri = string.Format("{0}?{1}", relativeUri, HttpClient.BuildQuery(formParams));
            var r = client.GET(uri);
            SessionID = null;
        }

        protected void GetModules(JToken json, string path)
        {
            foreach(var token in json)
            {
                var text = (string)token["text"];
                var value = (string)token["value"];
                var realPath = string.Format("{0}.{1}", path, text);
                moduleMap[realPath] = value;
                var children = token["ChildNodes"];
                if(children != null) GetModules(children, realPath);
            }
        }

        protected string GetModuleID(string path)
        {
            if (!moduleMap.ContainsKey(path)) return null;
            return moduleMap[path];
        }
    }


    class ReporterException : Exception
    {
        protected int _code;

        public ReporterException(int code, string message) : base(message)
        {
            _code = code;
        }

        public int Code
        {
            get { return _code; }
        }

        public override string Message
        {
            get
            {
                return string.Format("{{\"errorCode\":{0},\"errorMsg\":\"{1}\"}}", _code, base.Message);
            }
        }
    }

}
