using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Reflection;
using CommonEventSender;
using JobOrdersService;
using JobPropsService;
using Newtonsoft.Json;
using System.Collections.Concurrent;
using System.Threading;
using System.Linq;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for the management of processing of input queue on printing of labels
    /// </summary>
    public sealed class PrintJobs : IDisposable
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
        /// The name of the configuration parameter for the Odata service url.
        /// </summary>
        private const string cOdataService = "OdataServiceUri";

        ///// <summary>
        ///// The name of the configuration parameter for the Ghost Script path
        ///// </summary>
        //private const string cGhostScriptPath = "GhostScriptPath";

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

        private ConcurrentDictionary<string, Thread> printThreadConcurrentDictionary = new ConcurrentDictionary<string, Thread>();

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
            SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("ODataServiceUrl = {0}", odataServiceUrl), EventLogEntryType.Information);

            //PrintLabelWS.ghostScriptPath = System.Configuration.ConfigurationManager.AppSettings[cGhostScriptPath];
            PrintLabelWS.SMTPHost = System.Configuration.ConfigurationManager.AppSettings[cSMTPHost];
            PrintLabelWS.SMTPPort = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cSMTPPort]);
            SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("SMTP config = {0}:{1}", PrintLabelWS.SMTPHost, PrintLabelWS.SMTPPort), EventLogEntryType.Information);


            try
            {
                wmiProductInfo = new PrintServiceProductInfo(cServiceTitle,
                                                         Environment.MachineName,
                                                         Assembly.GetExecutingAssembly().GetName().Version.ToString(),
                                                         DateTime.Now,
                                                         odataServiceUrl);
            }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
            catch (Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
            {
                //SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Failed to initialize WMI = {0}", ex.ToString()), EventLogEntryType.Error);
            }

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
        ~PrintJobs()
        {
            if (eventLog != null)
            {
                eventLog.Close();
                eventLog.Dispose();
            }

            if (printTimer != null)
            {
                printTimer.Close();
                printTimer.Dispose();
            }
        }

        public void Dispose()
        {
            if (eventLog != null)
            {
                eventLog.Close();
                eventLog.Dispose();
            }

            if (printTimer != null)
            {
                printTimer.Close();
                printTimer.Dispose();
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
            string lLastError = string.Empty;
            printTimer.Stop();
            SenderMonitorEvent.sendMonitorEvent(EventLog, "Monitoring the print activity", EventLogEntryType.Information);

            try
            {
                RemoveAllNotAliveThreads();
                var printerAddresesToIgnore = printThreadConcurrentDictionary.Select(t => t.Key);

                LabeldbData lDbData = new LabeldbData(odataServiceUrl);
                JobOrders jobsToProcess = lDbData.getJobsToProcess();
                int CountJobsToProcess = jobsToProcess.JobOrdersObj.Count;
                SenderMonitorEvent.sendMonitorEvent(EventLog, "Jobs to process: " + CountJobsToProcess, EventLogEntryType.Information);

                if (CountJobsToProcess > 0)
                {
                    var groupedJobsToProcess = jobsToProcess.JobOrdersObj.GroupBy(j => j.PrinterIP);
                    foreach (var jobVal in groupedJobsToProcess)
                    {
                        try
                        {
                            JobOrders.JobOrdersValue[] printerJobArray = jobVal.OrderBy(o => o.ID).ToArray();

                            if (string.IsNullOrEmpty(jobVal.Key))
                            {
                                PrintJobProps job = lDbData.getJobData(EventLog, printerJobArray.First());
                                throw new Exception(string.Format("Printer IP address missing for printer {0}.", job.PrinterNo));
                            }
                            else
                            {
                                Thread printThread = new Thread(DoPrintWork);
                                if (printThreadConcurrentDictionary.TryAdd(jobVal.Key, printThread))
                                    printThread.Start(printerJobArray);
                                else
                                    SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Print Thread already exists, will be printed later, PrinterIP={0}.", jobVal.Key), EventLogEntryType.Information);
                                //throw new Exception(string.Format("Print Thread can't be created for printer IP {0}.", jobVal.Key));
                            }
                        }
                        catch (Exception ex)
                        {
                            string details = GetWebExceptionDetails(ex);
                            lLastError = "Error: " + ex.ToString() + " Details: " + details;
                            SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                            if (wmiProductInfo != null)
                                wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    string details = GetWebExceptionDetails(ex);
                    lLastError = "Error getting jobs: " + ex.ToString() + " Details: " + details;
                    SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                    if (wmiProductInfo != null)
                        wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                }
                catch (Exception exc)
                {
                    SenderMonitorEvent.sendMonitorEvent(EventLog, exc.Message, EventLogEntryType.Error);
                }
            }
            finally
            {
                printTimer.Start();
            }
        }

        private string GetWebExceptionDetails(Exception ex)
        {
            string details = string.Empty;
            if (ex is System.Net.WebException)
            {
                var resp = new StreamReader((ex as System.Net.WebException).Response.GetResponseStream()).ReadToEnd();

                try
                {
                    dynamic obj = JsonConvert.DeserializeObject(resp);
                    details = obj.error.message;
                }
                catch
                {
                    details = resp;
                }
            }
            return details;
        }

        private void DoPrintWork(object data)
        {
            if (data is JobOrders.JobOrdersValue[])
            {
                int CountJobsToProcess = 0;
                string lLastError = string.Empty;

                try
                {
                    string lPrintState;
                    int lLastJobID = 0;
                    string lFactoryNumber = string.Empty;
                    LabeldbData lDbData = new LabeldbData(odataServiceUrl);
                    JobOrders.JobOrdersValue[] jobValues = data as JobOrders.JobOrdersValue[];
                    //CountJobsToProcess = jobValues.Length;

                    foreach (JobOrders.JobOrdersValue jobValue in jobValues)
                    {
                        CountJobsToProcess++;

                        try
                        {
                            PrintJobProps job = lDbData.getJobData(EventLog, jobValue);
                            lLastJobID = job.JobOrderID;
                            lFactoryNumber = job.getLabelParameter("FactoryNumber", "FactoryNumber");

                            if (job.isExistsTemplate)
                            {
                                string randomFileName = Path.GetRandomFileName().Replace(".", "");
                                //PrintLabelWS.ExcelTemplateFile = Path.GetTempPath() + "Label.xlsx";
                                //PrintLabelWS.PDFTemplateFile = Path.GetTempPath() + "Label.pdf";
                                //PrintLabelWS.BMPTemplateFile = Path.GetTempPath() + "Label.bmp";
                                PrintLabelWS printLabelWS = new PrintLabelWS()
                                {
                                    BMPTemplateFile = Path.GetTempPath() + randomFileName + ".bmp",
                                    ExcelTemplateFile = Path.GetTempPath() + randomFileName + ".xlsx",
                                    PDFTemplateFile = Path.GetTempPath() + randomFileName + ".pdf"
                                };                                 

                                if (job.Command == "Print")
                                {
                                    string printerStatus = printLabelWS.getPrinterStatus(job.IpAddress, job.PrinterNo);
                                    Requests.updatePrinterStatus(odataServiceUrl, job.PrinterNo, printerStatus);
                                    if (!printerStatus.Equals("OK"))
                                    {
                                        throw new Exception(string.Format("Cannot print to {0}. Not valid printer status: {1}", job.PrinterNo, printerStatus));
                                    }
                                }
                                job.prepareTemplate(printLabelWS.ExcelTemplateFile);
                                if (job.Command == "Print")
                                {
                                    if (printLabelWS.PrintTemplate(job))
                                    {
                                        lPrintState = "Done";
                                        if (wmiProductInfo != null)
                                            wmiProductInfo.LastActivityTime = DateTime.Now;
                                    }
                                    else
                                    {
                                        lPrintState = "Failed";
                                    }
                                    lLastError = string.Format("JobOrderID: {0}. FactoryNumber: {3}. Print to: {1}. Status: {2}", job.JobOrderID, job.PrinterName, lPrintState, lFactoryNumber);
                                }
                                else
                                {
                                    if (printLabelWS.EmailTemplate(job))
                                    {
                                        lPrintState = "Done";
                                        if (wmiProductInfo != null)
                                            wmiProductInfo.LastActivityTime = DateTime.Now;
                                    }
                                    else
                                    {
                                        lPrintState = "Failed";
                                    }
                                    lLastError = string.Format("JobOrderID: {0}. FactoryNumber: {3}. Mail to: {1}. Status: {2}", job.JobOrderID, job.CommandRule, lPrintState, lFactoryNumber);
                                }
                                SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, lPrintState == "Failed" ? EventLogEntryType.Error : EventLogEntryType.Information);
                                if (lPrintState == "Failed")
                                {
                                    if (wmiProductInfo != null)
                                        wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                                }

                                //Clear All PrintLabelWS Temp Files
                                if (File.Exists(printLabelWS.BMPTemplateFile))
                                    File.Delete(printLabelWS.BMPTemplateFile);
                                if (File.Exists(printLabelWS.PDFTemplateFile))
                                    File.Delete(printLabelWS.PDFTemplateFile);
                                if (File.Exists(printLabelWS.ExcelTemplateFile))
                                    File.Delete(printLabelWS.ExcelTemplateFile);
                            }
                            else
                            {
                                lPrintState = "Failed";
                                lLastError = string.Format("Excel template is empty. JobOrderID: {0}. FactoryNumber: {1}.", job.JobOrderID, lFactoryNumber);
                                SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                                if (wmiProductInfo != null)
                                    wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                            }

                            if (lPrintState == "Done")
                            {
                                Requests.updateJobStatus(odataServiceUrl, job.JobOrderID, lPrintState);
                            }
                            else if (lPrintState == "Failed")
                            {
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            string details = GetWebExceptionDetails(ex);
                            lLastError = "JobOrderID: " + lLastJobID + ". FactoryNumber: " + lFactoryNumber + " Error: " + ex.ToString() + " Details: " + details;
                            SenderMonitorEvent.sendMonitorEvent(EventLog, lLastError, EventLogEntryType.Error);
                            if (wmiProductInfo != null)
                                wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);

                            break;
                        }
                    }
                }
                finally
                {
                    try
                    {
                        if (wmiProductInfo != null)
                        {
                            wmiProductInfo.PrintedLabelsCount += CountJobsToProcess;
                            wmiProductInfo.PublishInfo();
                        }
                        SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Print is done. {0} tasks", CountJobsToProcess), EventLogEntryType.Information);
                    }
                    catch (Exception exc)
                    {
                        SenderMonitorEvent.sendMonitorEvent(EventLog, exc.Message, EventLogEntryType.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Removes all done printer Threads
        /// </summary>
        private void RemoveAllNotAliveThreads()
        {
            var lvThreadsToRemove = printThreadConcurrentDictionary.Where(t => t.Value.IsAlive == false).Select(t => t.Key);
            foreach (var lvThreadToRemove in lvThreadsToRemove)
            {
                Thread lvRemovedThread;
                if (printThreadConcurrentDictionary.TryRemove(lvThreadToRemove, out lvRemovedThread) == false)
                    SenderMonitorEvent.sendMonitorEvent(EventLog, string.Format("Thread Printer IP:{0} Can't be removed.", lvThreadToRemove), EventLogEntryType.Warning);
            }
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
        /// Paper width in pixels
        /// </summary>
        public string PaperWidth
        {
            get { return getEquipmentProperty("PAPER_WIDTH"); }
        }

        /// <summary>
        /// Paper height in pixels
        /// </summary>
        public string PaperHeight
        {
            get { return getEquipmentProperty("PAPER_HEIGHT"); }
        }

        /// <summary>
        /// Printer NO
        /// </summary>
        public string PrinterNo
        {
            get { return getEquipmentProperty("PRINTER_NO"); }
        }

        /// <summary>
        /// Is exists template of label
        /// </summary>
        public bool isExistsTemplate
        {
            get { return (xlFile == null ? false : xlFile.Length > 0); }
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
            if (isExistsTemplate)
            {
                using (FileStream fs = new FileStream(excelTemplateFile, FileMode.Create))
                {
                    fs.Write(xlFile, 0, xlFile.Length);
                    //fs.Close();
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
