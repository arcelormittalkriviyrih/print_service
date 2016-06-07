using System;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for initialising of parameters label and printing of the set label
    /// </summary>
    public static class printLabel
    {
        public static int pingTimeoutInSeconds;
        public static EventLog vpEventLog;
        public static string templateFile;
        public static string xlsConverterPath;
        public static string ghostScriptPath;

        /// <summary>
        /// Printing of the prepared label
        /// </summary>
        public static bool printTemplate(jobPropsWS aJobProps)
        {
            //перед печатью если задан IP сделать пинг
            if ((pingTimeoutInSeconds > 0) & (aJobProps.IpAddress != ""))
            {
                System.Net.NetworkInformation.Ping printerPing = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply printerReply = printerPing.Send(aJobProps.IpAddress, pingTimeoutInSeconds);
                if (printerReply.Status != System.Net.NetworkInformation.IPStatus.Success)
                {
                    senderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("Printer {0}  {1}  ping timeout status {2}", aJobProps.PrinterName, aJobProps.IpAddress, printerReply.Status), EventLogEntryType.Warning);
                    return false;
                }
            }

            Boolean boolPrintLabel = false;
            Boolean boolConvertLabel = false;
            //LabelTemplate.vpEventLog = vpEventLog;
            LabelTemplate lTemplate = new LabelTemplate(templateFile);
            try
            {
                lTemplate.FillParamValues(aJobProps);
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not prepare label template. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            try
            {
                boolConvertLabel = convertToPDF();
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not convert label template to pdf. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            if (boolConvertLabel)
            {
                boolPrintLabel = PrintPDF(aJobProps.PrinterName);
            }
            else
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error);
                return false;
            }

            return boolPrintLabel;
        }

        private static bool convertToPDF()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = "\"" + templateFile + "\" \"" + Path.GetTempPath() + "Label.pdf" + "\"";
            startInfo.FileName = xlsConverterPath;
            startInfo.UseShellExecute = false;

            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            Process process = null;
            process = Process.Start(startInfo);
            process.WaitForExit(30000);
            if (process.HasExited == false)
                process.Kill();
            int exitcode = process.ExitCode;
            process.Close();
            return exitcode == 0;
        }

        private static bool PrintPDF(string printerName)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = " -sDEVICE=mswinpr2 -dLastPage=1 -dBATCH -dNOPAUSE -dPrinted -dNOSAFER -dNOPROMPT -dQUIET -sOutputFile=\"\\\\spool\\" + printerName + "\" \"" + Path.GetTempPath() + "Label.pdf" + "\" ";
            startInfo.FileName = ghostScriptPath;
            startInfo.UseShellExecute = false;

            startInfo.RedirectStandardError = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            Process process = null;
            process = Process.Start(startInfo);
            process.WaitForExit(30000);
            if (process.HasExited == false)
                process.Kill();
            int exitcode = process.ExitCode;
            process.Close();
            return exitcode == 0;
        }
    }
}
