using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using CommonEventSender;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for initialising of parameters label and printing of the set label
    /// </summary>
    public static class printLabelWS
    {
        public static int pingTimeoutInSeconds;
        public static EventLog vpEventLog;
        public static string ExcelTemplateFile;
        public static string PDFTemplateFile;
        public static string xlsConverterPath;
        public static string ghostScriptPath;
        public static string SMTPHost;
        public static int SMTPPort;

        /// <summary>
        /// Printing of the prepared label
        /// </summary>
        public static bool printTemplate(PrintJobProps aJobProps)
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
            if (preparePDF(aJobProps))
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
            File.Delete(PDFTemplateFile);
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.Arguments = "\"" + ExcelTemplateFile + "\" \"" + PDFTemplateFile + "\"";
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
            startInfo.Arguments = " -sDEVICE=mswinpr2 -dLastPage=1 -dBATCH -dNOPAUSE -dPrinted -dNOSAFER -dNOPROMPT -dQUIET -sOutputFile=\"\\\\spool\\" + printerName + "\" \"" + PDFTemplateFile + "\" ";
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

        private static bool preparePDF(PrintJobProps aJobProps)
        {
            Boolean boolConvertLabel = false;
            LabelTemplate lTemplate = new LabelTemplate(ExcelTemplateFile);
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

            return boolConvertLabel;
        }
        private static bool EmailPDF(string emailAddresses)
        {
            MailMessage mail;
            try
            {
                using (mail = new MailMessage())
                {
                    string mailFrom = "";
                    String[] mailtoList = emailAddresses.Split(',');
                    foreach (var mailTo in mailtoList)
                    {
                        if (mailTo != "")
                        {
                            mail.To.Add(new MailAddress(mailTo));
                            mailFrom = mailTo;
                        }
                    }
                    mail.From = new MailAddress(mailFrom);
                    mail.Subject = "Label";
                    mail.Body = "Label";
                    mail.Attachments.Add(new Attachment(PDFTemplateFile));
                    SmtpClient client = new SmtpClient();
                    client.Host = SMTPHost;
                    client.Port = SMTPPort;
                    //client.EnableSsl = true;
                    client.Credentials = CredentialCache.DefaultNetworkCredentials;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Send(mail);
                }
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not send email with label template. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Email of the prepared label
        /// </summary>
        public static bool emailTemplate(PrintJobProps aJobProps)
        {
            Boolean boolEmailLabel = false;
            if (preparePDF(aJobProps))
            {
                boolEmailLabel = EmailPDF(aJobProps.CommandRule);
            }
            else
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error);
                return false;
            }

            return boolEmailLabel;
        }
    }
}
