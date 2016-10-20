using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using CommonEventSender;
using ZSDK_API.Comm;
using ZSDK_API.Printer;
using JobOrdersService;

namespace PrintWindowsService
{
	/// <summary>
	/// Class for initialising of parameters label and printing of the set label
	/// </summary>
	public static class PrintLabelWS
	{
		//public static int pingTimeoutInSeconds;
		public static EventLog eventLog;
		public static string ExcelTemplateFile;
		public static string PDFTemplateFile;
		public static string BMPTemplateFile;		
		//public static string ghostScriptPath;
		public static string SMTPHost;
		public static int SMTPPort;

		/// <summary>
		/// Printing of the prepared label
		/// </summary>
		public static bool PrintTemplate(PrintJobProps jobProps)
		{
            //перед печатью если задан IP сделать пинг
   //         if ((pingTimeoutInSeconds > 0) && (jobProps.IpAddress != ""))
			//{
			//	System.Net.NetworkInformation.Ping printerPing = new System.Net.NetworkInformation.Ping();
			//	System.Net.NetworkInformation.PingReply printerReply = printerPing.Send(jobProps.IpAddress, pingTimeoutInSeconds);
			//	if (printerReply.Status != System.Net.NetworkInformation.IPStatus.Success)
			//	{
			//		SenderMonitorEvent.sendMonitorEvent(eventLog, string.Format("JobOrderID: {0} Printer {1}  {2}  ping timeout status {3}", jobProps.JobOrderID, jobProps.PrinterNo, jobProps.IpAddress, printerReply.Status), EventLogEntryType.Warning);
			//		return false;
			//	}
			//}
            
            Boolean boolPrintLabel = false;
			if (PrepareTemplate(jobProps, false))
			{
                boolPrintLabel = PrintZebra(jobProps.IpAddress, jobProps.PaperWidth, jobProps.PaperHeight, jobProps.JobOrderID, jobProps.PrinterNo);//PrintBMP(jobProps.PrinterName);            
            }
			else
			{
				//SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error);
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
            float dpi = 203f;
            if (!float.TryParse(System.Configuration.ConfigurationManager.AppSettings["ZebraPrinterDPI"], out dpi))
            {
                throw new Exception("Printer DPI is missing in config.");
            }
            xlsConverter.Program.Convert(ExcelTemplateFile, BMPTemplateFile, dpi, dpi, true);
			return File.Exists(BMPTemplateFile);
		}

		/*
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
		}*/
        

        public static string getPrinterStatus(string printerIpAddress, string printerNo)
        {
            
            ZebraPrinterConnection connection = null;
            try
            {
                int port = 9100;
                if (!int.TryParse(System.Configuration.ConfigurationManager.AppSettings["ZebraPrinterPort"], out port))
                {
                    throw new Exception("Printer port is missing in config.");
                }

                connection = new TcpPrinterConnection(printerIpAddress, port);
                connection.Open();
                ZebraPrinter printer = ZebraPrinterFactory.GetInstance(connection);

                PrinterStatus printerStatus = printer.GetCurrentStatus();
                if (printerStatus.IsReadyToPrint)
                {                    
                    return "OK";
                }
                else if (printerStatus.IsPaused)
                {
                    return "Paused";                    
                }
                else if (printerStatus.IsHeadOpen)
                {
                    return  "Head is open";                    
                }
                else if (printerStatus.IsPaperOut)
                {
                    return  "Out of paper";                    
                }
                else if (printerStatus.IsRibbonOut)
                {
                    return  "Ribbon is out";                    
                }
                else
                {
                    return  "Invalid status: "+ printerStatus.ToString();                    
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (connection != null && connection.IsConnected())
                    connection.Close();
            }
        }

        /// <summary>	Print BMP. </summary>
        ///
        /// <param name="printerName">	Name of the printer. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private static bool PrintZebra(string printerIpAddress, string width, string height, int JobOrderId, string printerNo)
		{
            bool result = true;
			ZebraPrinterConnection connection = null;
			try
			{
				if (string.IsNullOrEmpty(printerIpAddress))
					throw new Exception(string.Format("Printer IP address missing for printer {0}.", printerNo));
				if (string.IsNullOrEmpty(width))
					throw new Exception(string.Format("Paper width is null for printer {0}.", printerNo));
				if (string.IsNullOrEmpty(height))
					throw new Exception(string.Format("Paper height is null for {0}.", printerNo));

				int port = 9100;
				if (!int.TryParse(System.Configuration.ConfigurationManager.AppSettings["ZebraPrinterPort"], out port))
				{
					throw new Exception("Printer port is missing in config.");
				}
                
                int paperWidth = 0;
				if (!int.TryParse(width, out paperWidth))
				{
					throw new Exception(string.Format("Paper width is not integer for {0}.", printerNo));
				}

				int paperHeight = 0;
				if (!int.TryParse(height, out paperHeight))
				{
					throw new Exception(string.Format("Paper height is not integer for {0}.", printerNo));
				}

				connection = new TcpPrinterConnection(printerIpAddress, port);
				connection.Open();
				ZebraPrinter printer = ZebraPrinterFactory.GetInstance(connection);

				PrinterStatus printerStatus = printer.GetCurrentStatus();
				if (printerStatus.IsReadyToPrint)
				{
                    //printer.GetGraphicsUtil().PrintImage(BMPTemplateFile, 0, 0);
                    printer.GetGraphicsUtil().PrintImage(BMPTemplateFile, 0, 0, paperWidth, paperHeight, false);
				}
				else if (printerStatus.IsPaused)
				{
					throw new Exception(string.Format("Cannot Print because the printer {0} is paused.", printerNo));
				}
				else if (printerStatus.IsHeadOpen)
				{
					throw new Exception(string.Format("Cannot Print because the printer {0} head is open.", printerNo));
				}
				else if (printerStatus.IsPaperOut)
				{
					throw new Exception(string.Format("Cannot Print because the paper is out for printer {0}.", printerNo));
				}
				else if (printerStatus.IsRibbonOut)
				{
					throw new Exception(string.Format("Cannot Print because the ribbon is out for printer {0}.", printerNo));
				}
				else
				{
					throw new Exception(string.Format("Cannot print to {0}. Not valid printer status: {1}", printerNo, printerStatus.ToString()));
				}
			}
			catch (Exception ex)
			{
				SenderMonitorEvent.sendMonitorEvent(eventLog, "JobOrderId: "+ JobOrderId +". Print Zebra error: " + ex.ToString(), EventLogEntryType.Error);
				result = false;
			}
			finally
			{
				if (connection != null && connection.IsConnected())
					connection.Close();
			}

			return result;
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
