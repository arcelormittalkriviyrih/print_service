using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using JobOrdersService;

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
    public class LabeldbData
    {
        private string webServiceUrl;

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

        public LabeldbData(string aWebServiceUrl)
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
        public void fillPrintJobData(List<PrintJobProps> resultData)
        {
            byte[] XlFile = null;
            /*string JobOrdersUrl = CreateRequest("v_JobOrders?$filter=WorkType%20eq%20%27Print%27%20and%20DispatchStatus%20eq%20%27ToPrint%27&$select=ID,Command,CommandRule");
            string JobOrders = MakeRequest(JobOrdersUrl);
            List<JobOrdersValue> JobOrdersObj = DeserializeJobOrders(JobOrders);*/
            JobOrders jobOrders = new JobOrders(webServiceUrl, "Print", "ToPrint");

            foreach (JobOrders.JobOrdersValue joValue in jobOrders.JobOrdersObj)
            {
                string PrintJobParametersUrl = Requests.CreateRequest(webServiceUrl, String.Format("v_PrintJobParameters?$filter=JobOrderID%20eq%20{0}%20&$select=Property,Value",
                                                                      joValue.ID));
                string PrintJobParameters = Requests.MakeRequest(PrintJobParametersUrl);
                List<PrintJobParametersValue> PrintJobParametersObj = DeserializePrintJobParameters(PrintJobParameters);

                string PrinterID = getPrintJobParameter(PrintJobParametersObj, "PrinterID");
                string MaterialLotID = getPrintJobParameter(PrintJobParametersObj, "MaterialLotID");

                List<EquipmentPropertyValue> EquipmentPropertyObj = null;
                if (PrinterID != "")
                {
                    string EquipmentPropertyUrl = Requests.CreateRequest(webServiceUrl, String.Format("v_EquipmentProperty?$filter=EquipmentID%20eq%20{0}%20&$select=Property,Value",
                                                                         PrinterID));
                    string EquipmentProperty = Requests.MakeRequest(EquipmentPropertyUrl);
                    EquipmentPropertyObj = DeserializeEquipmentProperty(EquipmentProperty);
                }

                string PrintPropertiesUrl = Requests.CreateRequest(webServiceUrl, String.Format("v_PrintProperties?$filter=MaterialLotID%20eq%20{0}&$select=TypeProperty,PropertyCode,Value",
                                                                   MaterialLotID));
                string PrintPropertiesResponse = Requests.MakeRequest(PrintPropertiesUrl);
                List<PrintPropertiesValue> PrintPropertiesObj = DeserializePrintProperties(PrintPropertiesResponse);

                string TemplateUrl = Requests.CreateRequest(webServiceUrl, String.Format("v_PrintFile?$filter=MaterialLotID%20eq%20{0}%20and%20Property%20eq%20%27{1}%27&$select=Data",
                                                            MaterialLotID, "TEMPLATE"));
                string TemplateResponse = Requests.MakeRequest(TemplateUrl);
                List<LabelTemplateValue> LabelTemplateObj = DeserializeLabelTemplate(TemplateResponse);
                if (LabelTemplateObj.Count > 0)
                {
                    XlFile = LabelTemplateObj[0].Data;
                }

                resultData.Add(new PrintJobProps(joValue.ID,
                                                 joValue.Command,
                                                 (string)(joValue.CommandRule),
                                                 XlFile,
                                                 EquipmentPropertyObj,
                                                 PrintPropertiesObj));
            }
        }
    }
}
