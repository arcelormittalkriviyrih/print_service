using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace JobOrdersService
{
    /// <summary>
    /// Prepare request data from web service
    /// </summary>
    public static class Requests
    {
        /// <summary>
        /// Create final url for web service
        /// </summary>
        public static string CreateRequest(string webServiceUrl, string queryString)
        {
            string UrlRequest = webServiceUrl + queryString;
            ///http://mssql2014srv/odata_unified_svc/api/Dynamic/
            return UrlRequest;
        }

        /// <summary>
        /// Request data from web service
        /// </summary>
        public static string MakeRequest(string requestUrl)
        {
            string responseText = "";
            HttpWebRequest request = WebRequest.Create(requestUrl) as HttpWebRequest;
            request.Credentials = CredentialCache.DefaultNetworkCredentials;
#if (DEBUG)
            request.Credentials = new NetworkCredential("atokar", "qcAL0ZEV", "ask-ad");
#endif
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                if (response.StatusCode != HttpStatusCode.OK)
                    throw new Exception(String.Format(
                    "Server error (HTTP {0}: {1}).",
                    response.StatusCode,
                    response.StatusDescription));
                var encoding = ASCIIEncoding.ASCII;
                using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
                {
                    responseText = reader.ReadToEnd();
                }
                response.Close();
            }
            return responseText;
        }

        /// <summary>
        /// Update status of the job
        /// </summary>
        public static void updateJobStatus(string webServiceUrl, int aJobOrderID, string aActionState)
        {
            string UpdateStatusUrl = Requests.CreateRequest(webServiceUrl, String.Format("JobOrder({0})",
                                                            aJobOrderID));
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(UpdateStatusUrl);

            string payload = "{" + string.Format(@"""DispatchStatus"":""{0}""", aActionState) + "}";

            byte[] body = Encoding.UTF8.GetBytes(payload);
            request.Method = "PATCH";
            request.ContentLength = body.Length;
            request.ContentType = "application/json";
            request.Credentials = CredentialCache.DefaultNetworkCredentials;
#if (DEBUG)
            request.Credentials = new NetworkCredential("atokar", "qcAL0ZEV", "ask-ad");
#endif

            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(body, 0, body.Length);
                stream.Close();
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                if (response.StatusCode != HttpStatusCode.NoContent)
                    throw new Exception(String.Format(
                    "Server error (HTTP {0}: {1}).",
                    response.StatusCode,
                    response.StatusDescription));

                var encoding = ASCIIEncoding.ASCII;
                using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
                {
                    string responseText = reader.ReadToEnd();
                }

                response.Close();
            }
        }
    }
    /// <summary>
    /// Job orders list
    /// </summary>
    public class JobOrders
    {
        private string webServiceUrl;
        private List<JobOrdersValue> jobOrdersObj = null;

        public List<JobOrdersValue> JobOrdersObj
        {
            get { return jobOrdersObj; }
        }

        public JobOrders(string aWebServiceUrl, string WorkType, string DispatchStatus)
        {
            webServiceUrl = aWebServiceUrl;
            string JobOrdersUrl = Requests.CreateRequest(webServiceUrl, "v_JobOrders?$filter=WorkType%20eq%20%27" + WorkType + "%27%20and%20DispatchStatus%20eq%20%27" + DispatchStatus + "%27&$select=ID,Command,CommandRule");
            string JobOrdersSerial = Requests.MakeRequest(JobOrdersUrl);
            jobOrdersObj = DeserializeJobOrders(JobOrdersSerial);
        }

        public class JobOrdersValue
        {
            public int ID { get; set; }
            public string Command { get; set; }
            public object CommandRule { get; set; }
        }

        private class JobOrdersRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }
            public List<JobOrdersValue> value { get; set; }
        }

        private List<JobOrdersValue> DeserializeJobOrders(string json)
        {
            JobOrdersRoot prRoot = JsonConvert.DeserializeObject<JobOrdersRoot>(json);
            return prRoot.value;
        }
    }
}