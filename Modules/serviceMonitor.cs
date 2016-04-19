﻿using System;
using System.Diagnostics;
using System.Management.Instrumentation;

[assembly: Instrumented("root\\PrintWindowsService")]

namespace PrintWindowsService
{
    [InstrumentationClass(InstrumentationType.Instance)]
    /// <summary>
    /// Class for a grant in WMI
    /// </summary>
    public class ProductInfo
    {
        private string prAppName;
        private string prComputerName;
        private string prVersion;
        private DateTime prStartTime;
        private int prPrintTaskFrequencyInSeconds;
        private int prPingTimeoutInSeconds;
        private string prDBConnectionString;
        private DateTime prLastActivityTime;
        private string prLastServiceError;
        private int prPrintedLabelsCount;

        /// <summary>
        /// Application name
        /// </summary>
        public string AppName
        {
            get { return prAppName; }
        }
        /// <summary>
        /// Computer name
        /// </summary>
        public string ComputerName
        {
            get { return prComputerName; }
        }
        /// <summary>
        /// Version
        /// </summary>
        public string Version
        {
            get { return prVersion; }
        }
        /// <summary>
        /// Time of app start
        /// </summary>
        public DateTime StartTime
        {
            get { return prStartTime; }
        }
        /// <summary>
        /// Print task frequency in seconds
        /// </summary>
        public int PrintTaskFrequencyInSeconds
        {
            get { return prPrintTaskFrequencyInSeconds; }
        }
        /// <summary>
        /// Ping timeout in seconds
        /// </summary>
        public int PingTimeoutInSeconds
        {
            get { return prPingTimeoutInSeconds; }
        }
        /// <summary>
        /// DB connection string
        /// </summary>
        public string DBConnectionString
        {
            get { return prDBConnectionString; }
        }
        /// <summary>
        /// Time from the moment of start
        /// </summary>
        public TimeSpan TimeFromStart
        {
            get { return DateTime.Now - prStartTime; }
        }
        /// <summary>
        /// Time of the last activity of service
        /// </summary>
        public DateTime LastActivityTime
        {
            get { return prLastActivityTime; }
            set { prLastActivityTime = value; }
        }
        /// <summary>
        /// Last error of service
        /// </summary>
        public string LastServiceError
        {
            get { return prLastServiceError; }
            set { prLastServiceError = value; }
        }
        /// <summary>
        /// Count of the printed labels
        /// </summary>
        public int PrintedLabelsCount
        {
            get { return prPrintedLabelsCount; }
            set { prPrintedLabelsCount = value; }
        }

        public ProductInfo(string cAppName,
                           string cComputerName,
                           string cVersion,
                           DateTime cStartTime,
                           int cPrintTaskFrequencyInSeconds,
                           int cPingTimeoutInSeconds,
                           string cDBConnectionString)
        {
            prAppName = cAppName;
            prComputerName = cComputerName;
            prVersion = cVersion;
            prStartTime = cStartTime;
            prPrintTaskFrequencyInSeconds = cPrintTaskFrequencyInSeconds;
            prPingTimeoutInSeconds = cPingTimeoutInSeconds;
            prDBConnectionString = cDBConnectionString;

            LastActivityTime = new DateTime(0);
            LastServiceError = "";
            PrintedLabelsCount = 0;

            PublishInfo();
        }

        public void PublishInfo()
        {
            Instrumentation.Publish(this);
        }
    }
        
    [InstrumentationClass(InstrumentationType.Event)]
    /// <summary>
    /// Event for a grant in WMI
    /// </summary>
    public class senderMonitorEvent
    {
        private string message;
        private EventLogEntryType eventType;
        private DateTime eventTime;
        
        /// <summary>
        /// Text of event massage
        /// </summary>
        public string Message
        {
            get { return message; }
        }
        /// <summary>
        /// Type of event massage
        /// </summary>
        public string EventTypeName
        {
            get { return eventType.ToString(); }
        }
        /// <summary>
        /// Time of event massage
        /// </summary>
        public DateTime EventTime
        {
            get { return eventTime; }
        }

        public senderMonitorEvent(EventLog cEventLog, string cMessage, EventLogEntryType cEventType)
        {
            message = cMessage;
            eventType = cEventType;
            eventTime = DateTime.Now;
            if (cEventLog != null)
            {
                cEventLog.WriteEntry(cMessage, cEventType);
            }
        }

        /// <summary>
        /// Create and fire event
        /// </summary>
        public static void sendMonitorEvent(EventLog cEventLog, string cMessage, EventLogEntryType cEventType)
        {
            senderMonitorEvent MonitorEvent = new senderMonitorEvent(cEventLog, cMessage, cEventType);
            Instrumentation.Fire(MonitorEvent);
        }
    }
}