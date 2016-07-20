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
    public static class PrintLabelWS
    {
        public static int pingTimeoutInSeconds;
        public static EventLog eventLog;
        public static string ExcelTemplateFile;
        public static string PDFTemplateFile;
        public static string BMPTemplateFile;
        public static string xlsConverterPath;
        public static string ghostScriptPath;
        public static string SMTPHost;
        public static int SMTPPort;

        /// <summary>
        /// Printing of the prepared label
        /// </summary>
        public static bool PrintTemplate(PrintJobProps jobProps)
        {
            //перед печатью если задан IP сделать пинг
            if ((pingTimeoutInSeconds > 0) && (jobProps.IpAddress != ""))
            {
                System.Net.NetworkInformation.Ping printerPing = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply printerReply = printerPing.Send(jobProps.IpAddress, pingTimeoutInSeconds);
                if (printerReply.Status != System.Net.NetworkInformation.IPStatus.Success)
                {
                    SenderMonitorEvent.sendMonitorEvent(eventLog, string.Format("Printer {0}  {1}  ping timeout status {2}", jobProps.PrinterName, jobProps.IpAddress, printerReply.Status), EventLogEntryType.Warning);
                    return false;
                }
            }

            Boolean boolPrintLabel = false;
            if (PrepareTemplate(jobProps, false))
            {
                boolPrintLabel = PrintBMP(jobProps.PrinterName);
            }
            else
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error);
                return false;
            }

            return boolPrintLabel;
        }

        /// <summary>	Converts this object to a PDF. </summary>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool ConvertToPDF()
        {
            //File.Delete(PDFTemplateFile);
            //ProcessStartInfo startInfo = new ProcessStartInfo();
            //startInfo.Arguments = "\"" + ExcelTemplateFile + "\" \"" + PDFTemplateFile + "\"";
            //startInfo.FileName = xlsConverterPath;
            //startInfo.UseShellExecute = false;

            //startInfo.RedirectStandardError = true;
            //startInfo.RedirectStandardOutput = true;
            //startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            //Process process = null;
            //process = Process.Start(startInfo);
            //process.WaitForExit(30000);
            //if (process.HasExited == false)
            //    process.Kill();
            //int exitcode = process.ExitCode;
            //process.Close();
            //return exitcode == 0;
            xlsConverter.Program.Convert(ExcelTemplateFile, PDFTemplateFile);
            return File.Exists(BMPTemplateFile);
        }

        /// <summary>	Converts this object to a BMP. </summary>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool ConvertToBMP()
        {
            //File.Delete(PDFTemplateFile);
            //ProcessStartInfo startInfo = new ProcessStartInfo();
            //startInfo.Arguments = "\"" + ExcelTemplateFile + "\" \"" + BMPTemplateFile + "\"";
            //startInfo.FileName = xlsConverterPath;
            //startInfo.UseShellExecute = false;

            //startInfo.RedirectStandardError = true;
            //startInfo.RedirectStandardOutput = true;
            //startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            //Process process = null;
            //process = Process.Start(startInfo);
            //process.WaitForExit(30000);
            //if (process.HasExited == false)
            //    process.Kill();
            //int exitcode = process.ExitCode;
            //process.Close();
            //return exitcode == 0;
            xlsConverter.Program.Convert(ExcelTemplateFile, BMPTemplateFile);
            return File.Exists(BMPTemplateFile);
        }

        /// <summary>	Print PDF. </summary>
        ///
        /// <param name="printerName">	Name of the printer. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
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

        /// <summary>	Print BMP. </summary>
        ///
        /// <param name="printerName">	Name of the printer. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool PrintBMP(string printerName)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo(BMPTemplateFile);
                info.Arguments = "\"" + printerName + "\"";
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;
                info.UseShellExecute = true;
                info.Verb = "PrintTo";
                Process process = null;
                process = Process.Start(info);

                process.WaitForExit(30000);
                if (process.HasExited == false)
                    process.Kill();
                int exitcode = process.ExitCode;
                process.Close();
                return exitcode == 0;
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Print BMP error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }
        }

        /// <summary>	Prepare document for print. </summary>
        ///
        /// <param name="jobProps">	The job properties. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool PrepareTemplate(PrintJobProps jobProps, bool isPDF)
        {
            Boolean boolConvertLabel = false;
            LabelTemplate lTemplate = new LabelTemplate(ExcelTemplateFile);
            try
            {
                lTemplate.FillParamValues(jobProps);
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not prepare label template. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            try
            {
                if (isPDF)
                   boolConvertLabel = ConvertToPDF();
                else
                   boolConvertLabel = ConvertToBMP();
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert label template to pdf. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            return boolConvertLabel;
        }

        /// <summary>	Email PDF. </summary>
        ///
        /// <param name="emailAddresses">	The email addresses. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool EmailPDF(string emailAddresses)
        {
            MailMessage mail;
            try
            {
                using (mail = new MailMessage())
                {
                    string mailFrom = "";
                    string[] mailtoList = emailAddresses.Split(',');
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
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not send email with label template. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Email of the prepared label
        /// </summary>
        public static bool EmailTemplate(PrintJobProps jobProps)
        {
            Boolean boolEmailLabel = false;
            if (PrepareTemplate(jobProps, true))
            {
                boolEmailLabel = EmailPDF(jobProps.CommandRule);
            }
            else
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error);
                return false;
            }

            return boolEmailLabel;
        }
    }
}
