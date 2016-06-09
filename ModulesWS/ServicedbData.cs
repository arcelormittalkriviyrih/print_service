using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace PrintWindowsService
{
    /// <summary>
    /// Equipment properties
    /// </summary>
    public class EquipmentPropertyValue
    {
        public string Property { get; set; }
        public object Value { get; set; }
    }
    /// <summary>
    /// Property values class
    /// </summary>
    public class PrintPropertiesValue
    {
        public string TypeProperty { get; set; }
        public string PropertyCode { get; set; }
        public string Value { get; set; }
    }

    /// <summary>
    /// Class for processing of input queue and generation of list of labels for printing
    /// </summary>
    public class ServicedbData
    {
        private string webServiceUrl;

        /// <summary>
        /// Job orders for print
        /// </summary>
        private class JobOrdersValue
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

        /// <summary>
        /// Print job parameters
        /// </summary>
        private class PrintJobParametersValue
        {
            public string Property { get; set; }
            public object Value { get; set; }
        }

        private class PrintJobParametersRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }
            public List<PrintJobParametersValue> value { get; set; }
        }

        private class EquipmentPropertyRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }
            public List<EquipmentPropertyValue> value { get; set; }
        }

        private class PrintPropertiesRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }
            public List<PrintPropertiesValue> value { get; set; }
        }

        /// <summary>
        /// Label template file data
        /// </summary>
        private class LabelTemplateValue
        {
            public byte[] Data { get; set; }
        }

        private class LabelTemplateRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }
            public List<LabelTemplateValue> value { get; set; }
        }

        private List<JobOrdersValue> DeserializeJobOrders(string json)
        {
            JobOrdersRoot prRoot = JsonConvert.DeserializeObject<JobOrdersRoot>(json);
            return prRoot.value;
        }

        private List<PrintJobParametersValue> DeserializePrintJobParameters(string json)
        {
            PrintJobParametersRoot prRoot = JsonConvert.DeserializeObject<PrintJobParametersRoot>(json);
            return prRoot.value;
        }

        private List<EquipmentPropertyValue> DeserializeEquipmentProperty(string json)
        {
            EquipmentPropertyRoot prRoot = JsonConvert.DeserializeObject<EquipmentPropertyRoot>(json);
            return prRoot.value;
        }

        private List<PrintPropertiesValue> DeserializePrintProperties(string json)
        {
            PrintPropertiesRoot ppRoot = JsonConvert.DeserializeObject<PrintPropertiesRoot>(json);
            return ppRoot.value;
        }

        private List<LabelTemplateValue> DeserializeLabelTemplate(string json)
        {
            LabelTemplateRoot ltRoot = JsonConvert.DeserializeObject<LabelTemplateRoot>(json);
            return ltRoot.value;
        }

        /// <summary>
        /// Create final url for web service
        /// </summary>
        private string CreateRequest(string queryString)
        {
            string UrlRequest = webServiceUrl + queryString;
            ///http://mssql2014srv/odata_unified_svc/api/Dynamic/
            return UrlRequest;
        }

        /// <summary>
        /// Request data from web service
        /// </summary>
        static string MakeRequest(string requestUrl)
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

        public ServicedbData(string aWebServiceUrl)
        {
            webServiceUrl = aWebServiceUrl;
        }

        /// <summary>
        /// Return print job parameter value by Property
        /// </summary>
        private string getPrintJobParameter(List<PrintJobParametersValue> aPrintJobParametersObj, string aProperty)
        {
            string ParamValue = "";

            PrintJobParametersValue propertyFind = aPrintJobParametersObj.Find(x => (x.Property == aProperty));
            if (propertyFind != null)
            {
                ParamValue = propertyFind.Value == null ? "" : propertyFind.Value.ToString();
            }

            return ParamValue;
        }

        /// <summary>
        /// Processing of input queue and generation of list of labels for printing
        /// </summary>
        public void fillPrintJobData(List<jobPropsWS> resultData)
        {
            byte[] XlFile = null;
            string JobOrdersUrl = CreateRequest("v_JobOrders?$filter=WorkType%20eq%20%27Print%27%20and%20DispatchStatus%20eq%20%27ToPrint%27&$select=ID,Command,CommandRule");
            string JobOrders = MakeRequest(JobOrdersUrl);
            List<JobOrdersValue> JobOrdersObj = DeserializeJobOrders(JobOrders);

            foreach (JobOrdersValue joValue in JobOrdersObj)
            {
                string PrintJobParametersUrl = CreateRequest(String.Format("v_PrintJobParameters?$filter=JobOrderID%20eq%20{0}%20&$select=Property,Value",
                                                                            joValue.ID));
                string PrintJobParameters = MakeRequest(PrintJobParametersUrl);
                List<PrintJobParametersValue> PrintJobParametersObj = DeserializePrintJobParameters(PrintJobParameters);

                string PrinterID = getPrintJobParameter(PrintJobParametersObj, "PrinterID");
                string MaterialLotID = getPrintJobParameter(PrintJobParametersObj, "MaterialLotID");

                List<EquipmentPropertyValue> EquipmentPropertyObj = null;
                if (PrinterID != "")
                {
                    string EquipmentPropertyUrl = CreateRequest(String.Format("v_EquipmentProperty?$filter=EquipmentID%20eq%20{0}%20&$select=Property,Value",
                                                                               PrinterID));
                    string EquipmentProperty = MakeRequest(EquipmentPropertyUrl);
                    EquipmentPropertyObj = DeserializeEquipmentProperty(EquipmentProperty);
                }

                string PrintPropertiesUrl = CreateRequest(String.Format("v_PrintProperties?$filter=MaterialLotID%20eq%20{0}&$select=TypeProperty,PropertyCode,Value",
                                                                         MaterialLotID));
                string PrintPropertiesResponse = MakeRequest(PrintPropertiesUrl);
                List<PrintPropertiesValue> PrintPropertiesObj = DeserializePrintProperties(PrintPropertiesResponse);

                string TemplateUrl = CreateRequest(String.Format("v_PrintFile?$filter=MaterialLotID%20eq%20{0}%20and%20Property%20eq%20%27{1}%27&$select=Data",
                                                                  MaterialLotID, "TEMPLATE"));
                string TemplateResponse = MakeRequest(TemplateUrl);
                List<LabelTemplateValue> LabelTemplateObj = DeserializeLabelTemplate(TemplateResponse);
                if (LabelTemplateObj.Count > 0)
                {
                    XlFile = LabelTemplateObj[0].Data;
                }

                resultData.Add(new jobPropsWS(joValue.ID,
                                              joValue.Command,
                                              (string)(joValue.CommandRule),
                                              XlFile, 
                                              EquipmentPropertyObj, 
                                              PrintPropertiesObj));
            }
        }

        /// <summary>
        /// Update status of label print
        /// </summary>
        public void updateJobStatus(int aJobOrderID, string aPrintState)
        {
            string UpdateStatusUrl = CreateRequest(String.Format("JobOrder({0})", aJobOrderID));
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(UpdateStatusUrl);

            string payload = "{" + string.Format(@"""DispatchStatus"":""{0}""", aPrintState) + "}";

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
}