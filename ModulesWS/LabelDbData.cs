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
        /// <summary>	Gets or sets the property. </summary>
        ///
        /// <value>	The property. </value>
        public string Property { get; set; }

        /// <summary>	Gets or sets the value. </summary>
        ///
        /// <value>	The value. </value>
        public object Value { get; set; }
    }
    /// <summary>
    /// Property values class
    /// </summary>
    public class PrintPropertiesValue
    {
        /// <summary>	Gets or sets the type property. </summary>
        ///
        /// <value>	The type property. </value>
        public string TypeProperty { get; set; }

        /// <summary>	Gets or sets the property code. </summary>
        ///
        /// <value>	The property code. </value>
        public string PropertyCode { get; set; }

        /// <summary>	Gets or sets the value. </summary>
        ///
        /// <value>	The value. </value>
        public string Value { get; set; }
    }

    /// <summary>
    /// Class for processing of input queue and generation of list of labels for printing
    /// </summary>
    public class LabeldbData
    {
        /// <summary>	URL of the web service. </summary>
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

        /// <summary>	An equipment property root. </summary>
        private class EquipmentPropertyRoot
        {
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }

            /// <summary>	Gets or sets the value. </summary>
            ///
            /// <value>	The value. </value>
            public List<EquipmentPropertyValue> value { get; set; }
        }

        /// <summary>	A print properties root. </summary>
        private class PrintPropertiesRoot
        {
            /// <summary>	Gets or sets the metadata. </summary>
            ///
            /// <value>	The metadata. </value>
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }

            /// <summary>	Gets or sets the value. </summary>
            ///
            /// <value>	The value. </value>
            public List<PrintPropertiesValue> value { get; set; }
        }

        /// <summary>
        /// Label template file data
        /// </summary>
        private class LabelTemplateValue
        {
            /// <summary>	Gets or sets the data. </summary>
            ///
            /// <value>	The data. </value>
            public byte[] Data { get; set; }
        }

        /// <summary>	A label template root. </summary>
        private class LabelTemplateRoot
        {
            /// <summary>	Gets or sets the metadata. </summary>
            ///
            /// <value>	The metadata. </value>
            [JsonProperty("odata.metadata")]
            public string Metadata { get; set; }

            /// <summary>	Gets or sets the value. </summary>
            ///
            /// <value>	The value. </value>
            public List<LabelTemplateValue> value { get; set; }
        }

        /// <summary>	Deserialize print job parameters. </summary>
        ///
        /// <param name="json">	The JSON. </param>
        ///
        /// <returns>	A List&lt;PrintJobParametersValue&gt; </returns>
        private List<PrintJobParametersValue> DeserializePrintJobParameters(string json)
        {
            PrintJobParametersRoot prRoot = JsonConvert.DeserializeObject<PrintJobParametersRoot>(json);
            return prRoot.value;
        }

        /// <summary>	Deserialize equipment property. </summary>
        ///
        /// <param name="json">	The JSON. </param>
        ///
        /// <returns>	A List&lt;EquipmentPropertyValue&gt; </returns>
        private List<EquipmentPropertyValue> DeserializeEquipmentProperty(string json)
        {
            EquipmentPropertyRoot prRoot = JsonConvert.DeserializeObject<EquipmentPropertyRoot>(json);
            return prRoot.value;
        }

        /// <summary>	Deserialize print properties. </summary>
        ///
        /// <param name="json">	The JSON. </param>
        ///
        /// <returns>	A List&lt;PrintPropertiesValue&gt; </returns>
        private List<PrintPropertiesValue> DeserializePrintProperties(string json)
        {
            PrintPropertiesRoot ppRoot = JsonConvert.DeserializeObject<PrintPropertiesRoot>(json);
            return ppRoot.value;
        }

        /// <summary>	Deserialize label template. </summary>
        ///
        /// <param name="json">	The JSON. </param>
        ///
        /// <returns>	A List&lt;LabelTemplateValue&gt; </returns>
        private List<LabelTemplateValue> DeserializeLabelTemplate(string json)
        {
            LabelTemplateRoot ltRoot = JsonConvert.DeserializeObject<LabelTemplateRoot>(json);
            return ltRoot.value;
        }

        /// <summary>	Constructor. </summary>
        ///
        /// <param name="webServiceUrl">	URL of the web service. </param>
        public LabeldbData(string webServiceUrl)
        {
            this.webServiceUrl = webServiceUrl;
        }

        /// <summary>
        /// Return print job parameter value by Property
        /// </summary>
        private string getPrintJobParameter(List<PrintJobParametersValue> printJobParametersObj, string property)
        {
            string result = string.Empty;
            PrintJobParametersValue propertyFind = printJobParametersObj.Find(x => (x.Property == property));
            if (propertyFind != null)
            {
                result = propertyFind.Value == null ? string.Empty : propertyFind.Value.ToString();
            }
            return result;
        }

        /// <summary>
        /// Get print jobs to process
        /// </summary>
        public JobOrders getJobsToProcess()
        {
            JobOrders jobOrders = new JobOrders(webServiceUrl, "Print", "ToPrint");
            return jobOrders;            
        }

        /// <summary>
        /// Get job details
        /// </summary>
        public PrintJobProps getJobData(System.Diagnostics.EventLog EventLog, JobOrders.JobOrdersValue joValue)
        {
            string PrintJobParametersUrl = Requests.CreateRequest(webServiceUrl, string.Format("v_PrintJobParameters?$filter=JobOrderID%20eq%20{0}%20&$select=Property,Value",
                                                                                  joValue.ID));
            string PrintJobParameters = Requests.MakeRequest(PrintJobParametersUrl);
            List<PrintJobParametersValue> PrintJobParametersObj = DeserializePrintJobParameters(PrintJobParameters);
            //CommonEventSender.SenderMonitorEvent.sendMonitorEvent(EventLog, "Reading v_PrintJobParameters for job " + joValue.ID, System.Diagnostics.EventLogEntryType.Information);

            string PrinterID = getPrintJobParameter(PrintJobParametersObj, "PrinterID");
            string MaterialLotID = getPrintJobParameter(PrintJobParametersObj, "MaterialLotID");

            List<EquipmentPropertyValue> EquipmentPropertyObj = null;
            if (PrinterID != "")
            {
                string EquipmentPropertyUrl = Requests.CreateRequest(webServiceUrl, string.Format("v_EquipmentProperty?$filter=EquipmentID%20eq%20{0}%20&$select=Property,Value",
                                                                     PrinterID));
                string EquipmentProperty = Requests.MakeRequest(EquipmentPropertyUrl);
                EquipmentPropertyObj = DeserializeEquipmentProperty(EquipmentProperty);
            }
            //CommonEventSender.SenderMonitorEvent.sendMonitorEvent(EventLog, "Reading v_EquipmentProperty for job " + joValue.ID, System.Diagnostics.EventLogEntryType.Information);

            string PrintPropertiesUrl = Requests.CreateRequest(webServiceUrl, string.Format("v_PrintProperties?$filter=MaterialLotID%20eq%20{0}&$select=TypeProperty,PropertyCode,Value",
                                                               MaterialLotID));
            string PrintPropertiesResponse = Requests.MakeRequest(PrintPropertiesUrl);
            List<PrintPropertiesValue> PrintPropertiesObj = DeserializePrintProperties(PrintPropertiesResponse);
            //CommonEventSender.SenderMonitorEvent.sendMonitorEvent(EventLog, "Reading v_PrintProperties for job " + joValue.ID, System.Diagnostics.EventLogEntryType.Information);

            string TemplateUrl = Requests.CreateRequest(webServiceUrl, string.Format("v_PrintFile?$filter=MaterialLotID%20eq%20{0}%20and%20Property%20eq%20%27{1}%27&$select=Data",
                                                        MaterialLotID, "TEMPLATE"));
            //test string TemplateUrl = Requests.CreateRequest(webServiceUrl, "Files?$filter=ID%20eq%2068&$select=Data");
            string TemplateResponse = Requests.MakeRequest(TemplateUrl);
            List<LabelTemplateValue> LabelTemplateObj = DeserializeLabelTemplate(TemplateResponse);
            //CommonEventSender.SenderMonitorEvent.sendMonitorEvent(EventLog, "Reading v_PrintFile for job " + joValue.ID, System.Diagnostics.EventLogEntryType.Information);
            byte[] XlFile = null;

            if (LabelTemplateObj.Count > 0)
            {
                XlFile = LabelTemplateObj[0].Data;
            }

            return new PrintJobProps(joValue.ID,
                                             joValue.Command,
                                             (string)(joValue.CommandRule),
                                             XlFile,
                                             EquipmentPropertyObj,
                                             PrintPropertiesObj);
        }
    }
}
