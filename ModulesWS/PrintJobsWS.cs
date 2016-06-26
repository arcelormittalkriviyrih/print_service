using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Reflection;
using CommonEventSender;
using JobOrdersService;
using JobPropsService;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for the management of processing of input queue on printing of labels
    /// </summary>
    public class PrintJobs
    {
        #region Const

        private const string cServiceTitle = "Сервис печати этикеток";
        /// <summary>
        /// The name of the system event source used by this service.
        /// </summary>
        private const string cSystemEventSourceName = "ArcelorMittal.PrintService.EventSource";

        /// <summary>
        /// The name of the system event log used by this service.
        /// </summary>
        private const string cSystemEventLogName = "AM.PrintService.ArcelorMittal.Log";

        /// <summary>
        /// The name of the configuration parameter for the print task frequency in seconds.
        /// </summary>
        private const string cPrintTaskFrequencyName = "PrintTaskFrequency";

        /// <summary>
        /// The name of the configuration parameter for the ping timeout in seconds.
        /// </summary>
        private const string cPingTimeoutName = "PingTimeout";

        /// <summary>
        /// The name of the configuration parameter for the Odata service url.
        /// </summary>
        private const string cOdataService = "OdataServiceUri";

        /// <summary>
        /// The name of the configuration parameter for the XlsConverter path
        /// </summary>
        private const string cXlsConverterPath = "XlsConverterPath";

        /// <summary>
        /// The name of the configuration parameter for the Ghost Script path
        /// </summary>
        private const string cGhostScriptPath = "GhostScriptPath";

        /// <summary>
        /// The name of the configuration parameter for the SMTP host
        /// </summary>
        private const string cSMTPHost = "SMTPHost";

        /// <summary>
        /// The name of the configuration parameter for the SMTP port
        /// </summary>
        private const string cSMTPPort = "SMTPPort";

        #endregion

        #region Fields

        /// <summary>
        /// Time interval for checking print tasks
        /// </summary>
        private System.Timers.Timer printTimer;

        private PrintServiceProductInfo wmiProductInfo;
        private bool jobStarted = false;
        private string odataServiceUrl;

		/// <summary>	The event log. </summary>
		private EventLog eventLog;

        #endregion

        #region Property

        /// <summary>
        /// Gets the event log which is used by the service.
        /// </summary>
        public EventLog EventLog
        {
            get
            {
                lock (this)
                {
                    if (eventLog == null)
                    {
                        string lSystemEventLogName = cSystemEventLogName;
                        eventLog = new EventLog();
                        if (!System.Diagnostics.EventLog.SourceExists(cSystemEventSourceName))
                        {
                            System.Diagnostics.EventLog.CreateEventSource(cSystemEventSourceName, lSystemEventLogName);
                        }
                        else
                        {
                            lSystemEventLogName = EventLog.LogNameFromSourceName(cSystemEventSourceName, ".");
                        }
                        eventLog.Source = cSystemEventSourceName;
                        eventLog.Log = lSystemEventLogName;
                        PrintLabelWS.eventLog = eventLog;

                        WindowsIdentity identity = WindowsIdentity.GetCurrent();
                        WindowsPrincipal principal = new WindowsPrincipal(identity);
                        if (principal.IsInRole(WindowsBuiltInRole.Administrator))
                        {
                            eventLog.ModifyOverflowPolicy(OverflowAction.OverwriteAsNeeded, 7);
                        }
                    }
                    return eventLog;
                }
            }
        }

		/// <summary>
		/// Status of processing of queue
		/// </summary>
		public bool JobStarted
		{
			get
			{
				return jobStarted;
			}
		}

        #endregion

        #region Constructor

        /// <summary>	Default constructor. </summary>
        public PrintJobs()
        {
            // Set up a timer to trigger every print task frequency.
            int printTaskFrequencyInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPrintTaskFrequencyName]);
            //dbConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[cConnectionStringName].ConnectionString;
            odataServiceUrl = System.Configuration.ConfigurationManager.AppSettings[cOdataService];

            PrintLabelWS.pingTimeoutInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPingTimeoutName]);
            PrintLabelWS.ExcelTemplateFile = Path.GetTempPath() + "Label.xlsx";
            PrintLabelWS.PDFTemplateFile = Path.GetTempPath() + "Label.pdf";
            PrintLabelWS.xlsConverterPath = System.Configuration.ConfigurationManager.AppSettings[cXlsConverterPath];
            PrintLabelWS.ghostScriptPath = System.Configuration.ConfigurationManager.AppSettings[cGhostScriptPath];
            PrintLabelWS.SMTPHost = System.Configuration.ConfigurationManager.AppSettings[cSMTPHost];
            PrintLabelWS.SMTPPort = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cSMTPPort]);

            wmiProductInfo = new PrintServiceProductInfo(cServiceTitle,
                                                         Environment.MachineName,
                                                         Assembly.GetExecutingAssembly().GetName().Version.ToString(),
                                                         DateTime.Now,
                                                         printTaskFrequencyInSeconds,
                                                         PrintLabelWS.pingTimeoutInSeconds,
                                                         odataServiceUrl);

            printTimer = new System.Timers.Timer();
            printTimer.Interval = printTaskFrequencyInSeconds * 1000; // seconds to milliseconds
            printTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnPrintTimer);

            SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Print Task Frequncy = {0}", printTaskFrequencyInSeconds), EventLogEntryType.Information);
        }

        #endregion

        #region Destructor

        /// <summary>
        /// Constructor that prevents a default instance of this class from being created.
        /// </summary>
        ~ PrintJobs()
        {
            if (eventLog != null)
            {
                eventLog.Close();
                eventLog.Dispose();
            }
        }
        #endregion

        #region Methods

        /// <summary>
        /// Start of processing of input queue
        /// </summary>
        public void StartJob()
        {
            /*if (printLabel.lTemplate == null)
            {
                try
                {
                    printLabel.lTemplate = new LabelTemplate(printLabel.templateFile);
                }
                catch (Exception ex)
                {
                    string lLastError = "Error of Excel start: " + ex.ToString();
                    SenderMonitorEvent.sendMonitorEvent(vpEventLog, lLastError, EventLogEntryType.Error);
                    wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                    wmiProductInfo.PublishInfo();
                }
            }*/

            SenderMonitorEvent.sendMonitorEvent(EventLog, "Starting print service...", EventLogEntryType.Information);

            printTimer.Start();

            SenderMonitorEvent.sendMonitorEvent(EventLog, "Print service has been started", EventLogEntryType.Information);
            jobStarted = true;
        }

        /// <summary>
        /// Stop of processing of input queue
        /// </summary>
        public void StopJob()
        {
            SenderMonitorEvent.sendMonitorEvent(EventLog, "Stopping print service...", EventLogEntryType.Information);

            //stop timers if working
            if (printTimer.Enabled)
                printTimer.Stop();

            SenderMonitorEvent.sendMonitorEvent(EventLog, "Print service has been stopped", EventLogEntryType.Information);
            jobStarted = false;
        }

        /// <summary>
        /// Processing of input queue
        /// </summary>
        public void OnPrintTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            SenderMonitorEvent.sendMonitorEvent(EventLog, "Monitoring the print activity", EventLogEntryType.Information);
            printTimer.Stop();

            string lLastError = string.Empty;
            List<PrintJobProps> JobData = new List<PrintJobProps>();
            try
            {
                string printState;
                LabeldbData lDbData = new LabeldbData(odataServiceUrl);
                lDbData.fillPrintJobData(JobData);

                foreach (PrintJobProps job in JobData)
                {
                    if (job.isExistsTemplate)
                    {
                        job.prepareTemplate(PrintLabelWS.ExcelTemplateFile);
                        if (job.Command == "Print")
                        {
                            if (PrintLabelWS.PrintTemplate(job))
                            {
                                printState = "Done";
                                wmiProductInfo.LastActivityTime = DateTime.Now;
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            lLastError = string.Format("JobOrderID: {0}. Print to: {1}. Status: {2}", job.JobOrderID, job.PrinterName, printState);
                        }
                        else
                        {
                            if (PrintLabelWS.EmailTemplate(job))
                            {
                                printState = "Done";
                                wmiProductInfo.LastActivityTime = DateTime.Now;
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            lLastError = string.Format("JobOrderID: {0}. Mail to: {1}. Status: {2}", job.JobOrderID, job.CommandRule, printState);
                        }
                        SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, printState == "Failed" ? EventLogEntryType.FailureAudit : EventLogEntryType.SuccessAudit);
                        if (printState == "Failed")
                        {
                            wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                        }
                    }
                    else
                    {
                        printState = "Failed";
                        lLastError = "Excel template is empty";
                        SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                        wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                    }

                    if (printState == "Done")
                    {
                        Requests.updateJobStatus(odataServiceUrl, job.JobOrderID, printState);
                    }
                }
            }
            catch (Exception ex)
            {
                lLastError = "Get data from DB. Error: " + ex.ToString();
                SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
            }
            wmiProductInfo.PrintedLabelsCount += JobData.Count;
            wmiProductInfo.PublishInfo();
            SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Print is done. {0} tasks", JobData.Count), EventLogEntryType.Information);

            printTimer.Start();
        }
        #endregion
    }

    /// <summary>
    /// Class of label for print
    /// </summary>
    public class PrintJobProps : JobProps
    {
        private byte[] xlFile;
        private List<EquipmentPropertyValue> tableEquipmentProperty;
        private List<PrintPropertiesValue> tableLabelProperty;

        /// <summary>
        /// Printer name for print label
        /// </summary>
        public string PrinterName
        {
            get { return getEquipmentProperty("PRINTER_NAME"); }
        }
        /// <summary>
        /// IP of printer
        /// </summary>
        public string IpAddress
        {
            get { return getEquipmentProperty("PRINTER_IP"); }
        }
        /// <summary>
        /// Is exists template of label
        /// </summary>
        public bool isExistsTemplate
        {
            get { return xlFile.Length > 0; }
        }

        /// <summary>	Constructor. </summary>
        ///
        /// <param name="jobOrderID">			  	Identifier for the job order. </param>
        /// <param name="command">				  	The command. </param>
        /// <param name="commandRule">			  	The command rule. </param>
        /// <param name="xlFile">				  	The xl file. </param>
        /// <param name="tableEquipmentProperty">	The table equipment property. </param>
        /// <param name="tableLabelProperty">	  	The table label property. </param>
        public PrintJobProps(int jobOrderID,
                             string command,
                             string commandRule,
                             byte[] xlFile,
                             List<EquipmentPropertyValue> tableEquipmentProperty,
                             List<PrintPropertiesValue> tableLabelProperty) : base(jobOrderID, 
                                                                                    command, 
                                                                                    commandRule)
        {
            this.xlFile = xlFile;
            this.tableEquipmentProperty = tableEquipmentProperty;
            this.tableLabelProperty = tableLabelProperty;
        }
        /// <summary>
        /// Prepare template for print
        /// </summary>
        public void prepareTemplate(string excelTemplateFile)
        {
            if (xlFile.Length > 0)
            {
                using (FileStream fs = new FileStream(excelTemplateFile, FileMode.Create))
                {
                    fs.Write(xlFile, 0, xlFile.Length);
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// Return label parameter value by TypeProperty and PropertyCode
        /// </summary>
        public string getLabelParameter(string typeProperty, string propertyCode)
        {
            string result = string.Empty;
            if (tableLabelProperty != null)
            {
                PrintPropertiesValue propertyFind = tableLabelProperty.Find(x => (x.TypeProperty == typeProperty) & (x.PropertyCode == propertyCode));
                if (propertyFind != null)
                {
                    result = propertyFind.Value;
                }
            }

            return result;
        }
        /// <summary>
        /// Return equipment property value by Property
        /// </summary>
        public string getEquipmentProperty(string property)
        {
            string result = string.Empty;
            if (tableEquipmentProperty != null)
            {
                EquipmentPropertyValue propertyFind = tableEquipmentProperty.Find(x => (x.Property == property));
                if (propertyFind != null)
                {
                    result = propertyFind.Value == null ? string.Empty : propertyFind.Value.ToString();
                }
            }
        
			return result;
        }
    }
}
