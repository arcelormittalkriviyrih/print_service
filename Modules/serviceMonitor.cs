using System;
using System.Diagnostics;
using System.Management.Instrumentation;

[assembly: Instrumented("root\\PrintWindowsService")]

namespace PrintWindowsService
{
    [InstrumentationClass(InstrumentationType.Instance)]
    public class ProductInfo
    {
        private string prName;
        private string prComputerName;
        private string prVersion;
        private DateTime prStartTime;
        private int prPrintTaskFrequencyInSeconds;
        private int prPingTimeoutInSeconds;
        private string prDBConnectionString;

        public string Name
        {
            get { return prName; }
        }
        public string ComputerName
        {
            get { return prComputerName; }
        }
        public string Version
        {
            get { return prVersion; }
        }
        public DateTime StartTime
        {
            get { return prStartTime; }
        }
        public int PrintTaskFrequencyInSeconds
        {
            get { return prPrintTaskFrequencyInSeconds; }
        }
        public int PingTimeoutInSeconds
        {
            get { return prPingTimeoutInSeconds; }
        }
        public string DBConnectionString
        {
            get { return prDBConnectionString; }
        }

        public ProductInfo(string cName,
                           string cComputerName,
                           string cVersion,
                           DateTime cStartTime,
                           int cPrintTaskFrequencyInSeconds,
                           int cPingTimeoutInSeconds,
                           string cDBConnectionString)
        {
            prName = cName;
            prComputerName = cComputerName;
            prVersion = cVersion;
            prStartTime = cStartTime;
            prPrintTaskFrequencyInSeconds = cPrintTaskFrequencyInSeconds;
            prPingTimeoutInSeconds = cPingTimeoutInSeconds;
            prDBConnectionString = cDBConnectionString;
            Instrumentation.Publish(this);
        }
    }
        
    [InstrumentationClass(InstrumentationType.Event)]
    public class senderMonitorEvent
    {
        private string message;
        private EventLogEntryType eventType;
        private DateTime eventTime;
        // Определяем свойства
        public string Message
        {
            get { return message; }
        }
        public string EventTypeName
        {
            get { return eventType.ToString(); }
        }
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

        public static void sendMonitorEvent(EventLog cEventLog, string cMessage, EventLogEntryType cEventType)
        {
            senderMonitorEvent MonitorEvent = new senderMonitorEvent(cEventLog, cMessage, cEventType);
            Instrumentation.Fire(MonitorEvent);
        }
    }
}