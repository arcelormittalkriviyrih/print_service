using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Reflection;
using CommonEventSender;

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
        private const string cSystemEventLogName = "ArcelorMittal.PrintService.Log";

        /// <summary>
        /// The name of the configuration parameter for the print task frequency in seconds.
        /// </summary>
        private const string cPrintTaskFrequencyName = "PrintTaskFrequency";

        /// <summary>
        /// The name of the configuration parameter for the ping timeout in seconds.
        /// </summary>
        private const string cPingTimeoutName = "PingTimeout";

        /*        /// <summary>
                /// The name of the configuration parameter for the DB connection string.
                /// </summary>
                private const string cConnectionStringName = "DBDataSource";*/

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
        private System.Timers.Timer m_PrintTimer;

        private PrintServiceProductInfo wmiProductInfo;
        private bool fJobStarted = false;
        //private string dbConnectionString;
        private string OdataServiceUrl;
        #endregion

        #region vpEventLog

        /// <summary>
        /// The value of the vpEventLog property.
        /// </summary>
        private EventLog m_EventLog;

        /// <summary>
        /// Gets the event log which is used by the service.
        /// </summary>
        public EventLog vpEventLog
        {
            get
            {
                lock (this)
                {
                    if (m_EventLog == null)
                    {
                        string lSystemEventLogName = cSystemEventLogName;
                        m_EventLog = new EventLog();
                        if (!System.Diagnostics.EventLog.SourceExists(cSystemEventSourceName))
                        {
                            System.Diagnostics.EventLog.CreateEventSource(cSystemEventSourceName, lSystemEventLogName);
                        }
                        else
                        {
                            lSystemEventLogName = EventLog.LogNameFromSourceName(cSystemEventSourceName, ".");
                        }
                        m_EventLog.Source = cSystemEventSourceName;
                        m_EventLog.Log = lSystemEventLogName;
                        printLabelWS.vpEventLog = m_EventLog;

                        WindowsIdentity identity = WindowsIdentity.GetCurrent();
                        WindowsPrincipal principal = new WindowsPrincipal(identity);
                        if (principal.IsInRole(WindowsBuiltInRole.Administrator))
                        {
                            m_EventLog.ModifyOverflowPolicy(OverflowAction.OverwriteAsNeeded, 7);
                        }
                    }
                    return m_EventLog;
                }
            }
        }

        #endregion
        /// <summary>
        /// Status of processing of queue
        /// </summary>
        public bool JobStarted
        {
            get
            {
                return fJobStarted;
            }
        }

        #region Constructor

        public PrintJobs()
        {
            // Set up a timer to trigger every print task frequency.
            int printTaskFrequencyInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPrintTaskFrequencyName]);
            //dbConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[cConnectionStringName].ConnectionString;
            OdataServiceUrl = System.Configuration.ConfigurationManager.AppSettings[cOdataService];

            printLabelWS.pingTimeoutInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPingTimeoutName]);
            printLabelWS.ExcelTemplateFile = Path.GetTempPath() + "Label.xlsx";
            printLabelWS.PDFTemplateFile = Path.GetTempPath() + "Label.pdf";
            printLabelWS.xlsConverterPath = System.Configuration.ConfigurationManager.AppSettings[cXlsConverterPath];
            printLabelWS.ghostScriptPath = System.Configuration.ConfigurationManager.AppSettings[cGhostScriptPath];
            printLabelWS.SMTPHost = System.Configuration.ConfigurationManager.AppSettings[cSMTPHost];
            printLabelWS.SMTPPort = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cSMTPPort]);

            wmiProductInfo = new PrintServiceProductInfo(cServiceTitle,
                                                         Environment.MachineName,
                                                         Assembly.GetExecutingAssembly().GetName().Version.ToString(),
                                                         DateTime.Now,
                                                         printTaskFrequencyInSeconds,
                                                         printLabelWS.pingTimeoutInSeconds,
                                                         OdataServiceUrl);

            m_PrintTimer = new System.Timers.Timer();
            m_PrintTimer.Interval = printTaskFrequencyInSeconds * 1000; // seconds to milliseconds
            m_PrintTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnPrintTimer);

            senderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("Print Task Frequncy = {0}", printTaskFrequencyInSeconds), EventLogEntryType.Information);
        }

        #endregion

        #region Destructor

        ~ PrintJobs()
        {
            if (m_EventLog != null)
            {
                m_EventLog.Close();
                m_EventLog.Dispose();
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
                    senderMonitorEvent.sendMonitorEvent(vpEventLog, lLastError, EventLogEntryType.Error);
                    wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                    wmiProductInfo.PublishInfo();
                }
            }*/

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Starting print service...", EventLogEntryType.Information);

            m_PrintTimer.Start();

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print service has been started", EventLogEntryType.Information);
            fJobStarted = true;
        }

        /// <summary>
        /// Stop of processing of input queue
        /// </summary>
        public void StopJob()
        {
            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Stopping print service...", EventLogEntryType.Information);

            //stop timers if working
            if (m_PrintTimer.Enabled)
                m_PrintTimer.Stop();

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print service has been stopped", EventLogEntryType.Information);
            fJobStarted = false;
        }

        /// <summary>
        /// Processing of input queue
        /// </summary>
        public void OnPrintTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Monitoring the print activity", EventLogEntryType.Information);
            m_PrintTimer.Stop();

            /*
            //временно для тестирования
            if (printLabel.lTemplate == null)
            {
                printLabel.lTemplate = new LabelTemplate();
            }
            //временно для тестирования*/

            string lLastError = "";
            List<PrintJobProps> JobData = new List<PrintJobProps>();
            try
            {
                string printState;
                //labelDbData lDbData = new labelDbData(dbConnectionString);
                LabeldbData lDbData = new LabeldbData(OdataServiceUrl);
                lDbData.fillPrintJobData(JobData);

                foreach (PrintJobProps job in JobData)
                {
                    if (job.isExistsTemplate)
                    {
                        job.prepareTemplate(printLabelWS.ExcelTemplateFile);
                        if (job.Command == "Print")
                        {
                            if (printLabelWS.printTemplate(job))
                            {
                                printState = "Done";
                                wmiProductInfo.LastActivityTime = DateTime.Now;
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            lLastError = String.Format("JobOrderID: {0}. Print to: {1}. Status: {2}", job.JobOrderID, job.PrinterName, printState);
                        }
                        else
                        {
                            if (printLabelWS.emailTemplate(job))
                            {
                                printState = "Done";
                                wmiProductInfo.LastActivityTime = DateTime.Now;
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            lLastError = String.Format("JobOrderID: {0}. Mail to: {1}. Status: {2}", job.JobOrderID, job.CommandRule, printState);
                        }
                        senderMonitorEvent.sendMonitorEvent(vpEventLog, lLastError, printState == "Failed" ? EventLogEntryType.FailureAudit : EventLogEntryType.SuccessAudit);
                        if (printState == "Failed")
                        {
                            wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                        }
                    }
                    else
                    {
                        printState = "Failed";
                        lLastError = "Excel template is empty";
                        senderMonitorEvent.sendMonitorEvent(vpEventLog, lLastError, EventLogEntryType.Error);
                        wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
                    }

                    if (printState == "Done")
                    {
                        lDbData.updateJobStatus(job.JobOrderID, printState);
                    }
                }
            }
            catch (Exception ex)
            {
                lLastError = "Get data from DB. Error: " + ex.ToString();
                senderMonitorEvent.sendMonitorEvent(vpEventLog, lLastError, EventLogEntryType.Error);
                wmiProductInfo.LastServiceError = string.Format("{0}. On {1}", lLastError, DateTime.Now);
            }
            wmiProductInfo.PrintedLabelsCount += JobData.Count;
            wmiProductInfo.PublishInfo();
            senderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("Print is done. {0} tasks", JobData.Count), EventLogEntryType.Information);

            m_PrintTimer.Start();
        }
        #endregion
    }
}
