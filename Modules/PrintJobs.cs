using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Reflection;

//[assembly:Instrumented("root\\PrintWindowsService")]
//using Aspose.Cells;
//using Aspose.Cells.Rendering;
//using System.Drawing;
//using System.Drawing.Imaging;
//using System.Runtime.InteropServices;

//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;

namespace PrintWindowsService
{
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

        /// <summary>
        /// The name of the configuration parameter for the print task frequency in seconds.
        /// </summary>
        private const string cConnectionStringName = "DBDataSource";

        #endregion

        #region Fields

        /// <summary>
        /// Time interval for checking print tasks
        /// </summary>
        private System.Timers.Timer m_PrintTimer;
        private bool fJobStarted = false;
        private string dbConnectionString;
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
                        printLabel.vpEventLog = m_EventLog;

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
            dbConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings[cConnectionStringName].ConnectionString;

            printLabel.pingTimeoutInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPingTimeoutName]);
            printLabel.templateFile = Path.GetTempPath() + "Label.xls";;

            if (printLabel.xl == null)
            {
                try
                {
                    printLabel.xl = new ExcelApplication();
                }
                catch (Exception ex)
                {
                    senderMonitorEvent.sendMonitorEvent(vpEventLog, "Error of Excel start: " + ex.ToString(), EventLogEntryType.Error);
                }
            }

            m_PrintTimer = new System.Timers.Timer();
            m_PrintTimer.Interval = printTaskFrequencyInSeconds * 1000; // seconds to milliseconds
            m_PrintTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnPrintTimer);

            ProductInfo wmiProductInfo = new ProductInfo(Environment.MachineName,
                                                         cServiceTitle,
                                                         Assembly.GetExecutingAssembly().GetName().Version.ToString(),
                                                         DateTime.Now,
                                                         printTaskFrequencyInSeconds,
                                                         printLabel.pingTimeoutInSeconds,
                                                         dbConnectionString);

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

        public void StartJob()
        {
            if (printLabel.xl == null)
            {
                try
                {
                    printLabel.xl = new ExcelApplication();
                }
                catch (Exception ex)
                {
                    senderMonitorEvent.sendMonitorEvent(vpEventLog, "Error of Excel start: " + ex.ToString(), EventLogEntryType.Error);
                }
            }

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Starting print service...", EventLogEntryType.Information);

            m_PrintTimer.Start();

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print service has been started", EventLogEntryType.Information);
            fJobStarted = true;
        }

        public void StopJob()
        {
            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Stopping print service...", EventLogEntryType.Information);

            //stop timers if working
            if (m_PrintTimer.Enabled)
                m_PrintTimer.Stop();

            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print service has been stopped", EventLogEntryType.Information);
            fJobStarted = false;
        }

        public void OnPrintTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            senderMonitorEvent.sendMonitorEvent(vpEventLog, "Monitoring the print activity", EventLogEntryType.Information);
            m_PrintTimer.Stop();

            //временно для тестирования
            if (printLabel.xl == null)
            {
                //работаем со своим экземпляром Excel
                printLabel.xl = new ExcelApplication();
            }
            //временно для тестирования

            List<jobProps> JobData = new List<jobProps>();
            try
            {
                string printState;
                labelDbData lDbData = new labelDbData(dbConnectionString);
                lDbData.fillJobData(ref JobData);

                foreach (jobProps job in JobData)
                {
                    if (job.isExistsTemplate)
                    {
                        job.prepareTemplate();
                        if (printLabel.printTemplate(job))
                        {
                            printState = "Printed";
                        }
                        else
                        {
                            printState = "Failed";
                        }
                        senderMonitorEvent.sendMonitorEvent(vpEventLog, String.Format("ProductionResponseID: {0}. Print to: {1}. Status: {2}", job.ProductionResponseID, job.PrinterName, printState), printState == "Failed" ? EventLogEntryType.FailureAudit : EventLogEntryType.SuccessAudit);
                    }
                    else
                    {
                        printState = "Failed";
                        senderMonitorEvent.sendMonitorEvent(vpEventLog, "Excel template is empty", EventLogEntryType.Error);
                    }

                    lDbData.updateJobStatus(job.ProductionResponseID, printState);
                }
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Get data from DB. Error: " + ex.ToString(), EventLogEntryType.Error);
            }
            senderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("Print is done. {0} tasks", JobData.Count), EventLogEntryType.Information);

            m_PrintTimer.Start();
        }
        #endregion
    }
}
