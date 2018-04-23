using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;
using System.Text;

namespace Util
{
    class HttpClient
    {
        public event RequestDelegate OnGetRequest;
        public event ResponseContentDelegate OnGetResponseContent;

        protected Dictionary<string, object> config;
        protected CookieContainer cookies;
        protected Encoding readEncoding = Encoding.UTF8;
        protected Encoding writeEncoding = Encoding.UTF8;

        public HttpClient()
        {
            cookies = new CookieContainer();
            config = new Dictionary<string, object>(){
                {"BaseUri", null},
                {"AllowAutoRedirect", true},
                {"ContentType", null},
                {"CookieContainer", cookies},
                {"Referer", null},
                {"Timeout", 5000},
                {"UserAgent", "Csharp HttpClient"},
            };
        }

        public HttpClient(Dictionary<string, object> config) : this()
        {
            this.config = MergeConfig(this.config, config);
        }

        public static HttpWebRequest GetRequestByConfig(string uri, Dictionary<string, object> config)
        {
            var baseUri = GetConfig(config, "BaseUri") as string;
            if (baseUri != null)
            {
                uri = new Uri(new Uri(baseUri), uri).AbsoluteUri;
            }
            var request = WebRequest.Create(uri) as HttpWebRequest;
            var allowAutoRedirect = GetConfig(config, "AllowAutoRedirect") as bool?;
            if (allowAutoRedirect != null) request.AllowAutoRedirect = (bool)allowAutoRedirect;
            var contentType = GetConfig(config, "ContentType") as string;
            if (contentType != null) request.ContentType = contentType;
            var cookieContainer = GetConfig(config, "CookieContainer") as CookieContainer;
            if (cookieContainer != null) request.CookieContainer = cookieContainer;
            var referer = GetConfig(config, "Referer") as string;
            if (referer != null) request.Referer = referer;
            var timeout = GetConfig(config, "Timeout") as int?;
            if (timeout != null) request.Timeout = (int)timeout;
            var userAgent = GetConfig(config, "UserAgent") as string;
            if (userAgent != null) request.UserAgent = userAgent;
            return request;
        }

        public static Dictionary<string, object> MergeConfig(Dictionary<string, object> a, Dictionary<string, object> b)
        {
            var c = new Dictionary<string, object>(a);
            if (b == null) return c;
            foreach (var key in b.Keys) {
                c[key] = b[key];
            }
            return c;
        }

        public static object GetConfig(Dictionary<string, object> config, string key)
        {
            if (!config.ContainsKey(key)) return null;
            return config[key];
        }

        public static string BuildQuery(Dictionary<string, string> formParams)
        {
            var sb = new StringBuilder();
            bool f = true;
            foreach (var param in formParams)
            {
                if (f) f = false; else sb.Append('&');
                sb.Append(param.Key);
                sb.Append('=');
                sb.Append(HttpUtility.UrlEncode(param.Value));
            }
            return sb.ToString();
        }

        public Encoding WriteEncoding
        {
            get { return writeEncoding; }
            set { writeEncoding = value; }
        }

        public Encoding ReadEncoding
        {
            get { return readEncoding; }
            set { readEncoding = value; }
        }

        public Dictionary<string, object> Config
        {
            get { return config; }
        }

        public object GetConfig(string key)
        {
            return GetConfig(config, key);
        }

        protected HttpWebRequest GetRequest(string uri, Dictionary<string, object> config)
        {
            var request = GetRequestByConfig(uri, MergeConfig(this.config, config));
            OnGetRequest?.Invoke(request);
            return request;
        }

        public string GET(string uri)
        {
            return GET(uri, null);
        }

        public string GET(string uri, Dictionary<string, object> config)
        {
            var request = GetRequest(uri, config);
            return GetResponseBody(request);
        }

        public string POST(string uri, string body)
        {
            return POST(uri, body, null);
        }

        public string POST(string uri, Dictionary<string, string> formParams)
        {
            return POST(uri, formParams, null);
        }

        public string POST(string uri, Dictionary<string, string> formParams, Dictionary<string, object> config)
        {
            return POST(uri, BuildQuery(formParams), config);
        }

        public string POST(string uri, string body, Dictionary<string, object> config)
        {
            var request = GetRequest(uri, config);
            request.Method = "POST";
            byte[] bytes = WriteEncoding.GetBytes(body);
            using (Stream writer = request.GetRequestStream()) {
                writer.Write(bytes, 0, bytes.Length);
            }
            return GetResponseBody(request);
        }

        protected string GetResponseBody(HttpWebRequest request)
        {
            var response = request.GetResponse() as HttpWebResponse;
            string body;
            using (StreamReader reader = new StreamReader(response.GetResponseStream(), ReadEncoding)) {
                body = reader.ReadToEnd();
            }
            response.Close();
            OnGetResponseContent?.Invoke(body);
            return body;
        }

    }

    delegate void RequestDelegate(HttpWebRequest request);
    delegate void ResponseContentDelegate(string content);
}