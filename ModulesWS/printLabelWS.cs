using CommonEventSender;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Mail;
using Zebra.Sdk.Comm;
using Zebra.Sdk.Printer;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for initialising of parameters label and printing of the set label
    /// </summary>
    public class PrintLabelWS : IDisposable
    {
        //public static int pingTimeoutInSeconds;
        public EventLog eventLog;
        public string ExcelTemplateFile;
        public string PDFTemplateFile;
        public string BMPTemplateFile;
        //public static string ghostScriptPath;
        public static string SMTPHost;
        public static int SMTPPort;

        /// <summary>
        /// Printing of the prepared label
        /// </summary>
        public bool PrintTemplate(PrintJobProps jobProps)
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
        private bool ConvertToPDF()
        {
            //// Open a template excel file
            //using (Workbook book = new Workbook(ExcelTemplateFile))
            //{
            //    book.CalculateFormula();

            //    // Make all sheets invisible except first worksheet
            //    for (int i = 1; i < book.Worksheets.Count; i++)
            //    {
            //        book.Worksheets[i].IsVisible = false;
            //    }

            //    book.Save(PDFTemplateFile, SaveFormat.Pdf);
            //}

            xlsConverter.Program.ConvertNoRotate(ExcelTemplateFile, PDFTemplateFile);
            return File.Exists(PDFTemplateFile);
        }

        /// <summary>	Converts this object to a BMP. </summary>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private bool ConvertToBMP()
        {
            float dpi = 203f;
            if (!float.TryParse(System.Configuration.ConfigurationManager.AppSettings["ZebraPrinterDPI"], out dpi))
            {
                throw new Exception("Printer DPI is missing in config.");
            }

            //// Open a template excel file
            //using (Workbook book = new Workbook(ExcelTemplateFile))
            //{
            //    book.CalculateFormula();

            //    // Get the first worksheet.
            //    Worksheet sheet = book.Worksheets[0];

            //    // Define ImageOrPrintOptions
            //    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //    // Specify the image format
            //    imgOptions.ImageType = Aspose.Cells.Drawing.ImageType.Bmp;
            //    imgOptions.OnlyArea = true;
            //    //imgOptions.OnePagePerSheet = true;
            //    //imgOptions.IsCellAutoFit = true;
            //    imgOptions.HorizontalResolution = (int)dpi;
            //    imgOptions.VerticalResolution = (int)dpi;
            //    // Render the sheet with respect to specified image/print options
            //    SheetRender sr = new SheetRender(sheet, imgOptions);
            //    // Render the image for the sheet
            //    using (Bitmap bitmap = sr.ToImage(0))
            //    {
            //        bool rotate = true;
            //        using (Image croppedImage = AutoCrop(bitmap))
            //        {
            //            if (rotate)
            //            {
            //                croppedImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
            //            }
            //            croppedImage.Save(BMPTemplateFile);
            //        }

            //        //bitmap.Save(BMPTemplateFile);
            //    }
            //}

            xlsConverter.Program.Convert(ExcelTemplateFile, BMPTemplateFile, dpi, dpi, true);
            return File.Exists(BMPTemplateFile);
        }

        private static Image AutoCrop(Bitmap bmp)
        {
            if (Image.GetPixelFormatSize(bmp.PixelFormat) != 32)
                throw new InvalidOperationException("Autocrop currently only supports 32 bits per pixel images.");

            // Initialize variables
            var cropColor = System.Drawing.Color.White;

            var bottom = 0;
            var left = bmp.Width; // Set the left crop point to the width so that the logic below will set the left value to the first non crop color pixel it comes across.
            var right = 0;
            var top = bmp.Height; // Set the top crop point to the height so that the logic below will set the top value to the first non crop color pixel it comes across.

            var bmpData = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);

            unsafe
            {
                var dataPtr = (byte*)bmpData.Scan0;

                for (var y = 0; y < bmp.Height; y++)
                {
                    for (var x = 0; x < bmp.Width; x++)
                    {
                        var rgbPtr = dataPtr + (x * 4);

                        var b = rgbPtr[0];
                        var g = rgbPtr[1];
                        var r = rgbPtr[2];
                        var a = rgbPtr[3];

                        // If any of the pixel RGBA values don't match and the crop color is not transparent, or if the crop color is transparent and the pixel A value is not transparent
                        if ((cropColor.A > 0 && (b != cropColor.B || g != cropColor.G || r != cropColor.R || a != cropColor.A)) || (cropColor.A == 0 && a != 0))
                        {
                            if (x < left)
                                left = x;

                            if (x >= right)
                                right = x + 1;

                            if (y < top)
                                top = y;

                            if (y >= bottom)
                                bottom = y + 1;
                        }
                    }

                    dataPtr += bmpData.Stride;
                }
            }

            bmp.UnlockBits(bmpData);

            if (left < right && top < bottom)
                //return bmp.Clone(new Rectangle(left, top, right - left, bottom - top), bmp.PixelFormat);
                return bmp.Clone(new Rectangle(/*left*/0, /*top*/0, right, bottom), bmp.PixelFormat);

            return null; // Entire image should be cropped, so just return null
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


        public string getPrinterStatus(string printerIpAddress, string printerNo)
        {
            Connection connection = null;
            try
            {
                int port = 9100;
                if (!int.TryParse(System.Configuration.ConfigurationManager.AppSettings["ZebraPrinterPort"], out port))
                {
                    throw new Exception("Printer port is missing in config.");
                }

                connection = new TcpConnection(printerIpAddress, port);
                connection.Open();
                ZebraPrinter printer = ZebraPrinterFactory.GetInstance(connection);

                PrinterStatus printerStatus = printer.GetCurrentStatus();
                if (printerStatus.isReadyToPrint)
                {
                    return "OK";
                }
                else if (printerStatus.isPaused)
                {
                    return "Paused";
                }
                else if (printerStatus.isHeadOpen)
                {
                    return "Head is open";
                }
                else if (printerStatus.isPaperOut)
                {
                    return "Out of paper";
                }
                else if (printerStatus.isRibbonOut)
                {
                    return "Ribbon is out";
                }
                else
                {
                    return "Invalid status: " + printerStatus.ToString();
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (connection != null && connection.Connected)
                    connection.Close();
            }
        }

        /// <summary>	Print BMP. </summary>
        ///
        /// <param name="printerName">	Name of the printer. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private bool PrintZebra(string printerIpAddress, string width, string height, int JobOrderId, string printerNo)
        {
            bool result = true;
            Connection connection = null;
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

                connection = new TcpConnection(printerIpAddress, port);
                connection.Open();
                ZebraPrinter printer = ZebraPrinterFactory.GetInstance(connection);

                PrinterStatus printerStatus = printer.GetCurrentStatus();
                if (printerStatus.isReadyToPrint)
                {
                    //printer.GetGraphicsUtil().PrintImage(BMPTemplateFile, 0, 0);
                    printer.PrintImage(BMPTemplateFile, 0, 0, paperWidth, paperHeight, false);
                    //result = false;
                }
                else if (printerStatus.isPaused)
                {
                    throw new Exception(string.Format("Cannot Print because the printer {0} is paused.", printerNo));
                }
                else if (printerStatus.isHeadOpen)
                {
                    throw new Exception(string.Format("Cannot Print because the printer {0} head is open.", printerNo));
                }
                else if (printerStatus.isPaperOut)
                {
                    throw new Exception(string.Format("Cannot Print because the paper is out for printer {0}.", printerNo));
                }
                else if (printerStatus.isRibbonOut)
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
                SenderMonitorEvent.sendMonitorEvent(eventLog, "JobOrderId: " + JobOrderId + ". Print Zebra error: " + ex.ToString(), EventLogEntryType.Error, 4);
                result = false;
            }
            finally
            {
                if (connection != null && connection.Connected)
                    connection.Close();
            }

            return result;
        }

        /// <summary>	Print BMP. </summary>
        ///
        /// <param name="printerName">	Name of the printer. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private bool PrintBMP(string printerName)
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
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Print BMP error: " + ex.ToString(), EventLogEntryType.Error, 4);
                return false;
            }
        }

        /// <summary>	Prepare document for print. </summary>
        ///
        /// <param name="jobProps">	The job properties. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private bool PrepareTemplate(PrintJobProps jobProps, bool isPDF)
        {
            Boolean boolConvertLabel = false;
            LabelTemplate lTemplate = new LabelTemplate(ExcelTemplateFile);
            try
            {
                lTemplate.FillParamValues(jobProps);
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not prepare label template. Error: " + ex.ToString(), EventLogEntryType.Error, 4);
                return false;
            }
            try
            {
                //try
                //{
                //    Aspose.Cells.License lvLicense = new Aspose.Cells.License();
                //    lvLicense.SetLicense("Aspose.Cells.lic");
                //}
                //catch (Exception ex)
                //{
                //    SenderMonitorEvent.sendMonitorEvent(eventLog, "Aspose.Cells License Error: " + ex.ToString(), EventLogEntryType.Warning);
                //}

                if (isPDF)
                    boolConvertLabel = ConvertToPDF();
                else
                    boolConvertLabel = ConvertToBMP();
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert Excel label template to PDF/BMP. Error: " + ex.ToString(), EventLogEntryType.Error, 4);
                return false;
            }
            return boolConvertLabel;
        }

        /// <summary>	Email PDF. </summary>
        ///
        /// <param name="emailAddresses">	The email addresses. </param>
        ///
        /// <returns>	true if it succeeds, false if it fails. </returns>
        private bool EmailPDF(string emailAddresses)
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

                    try
                    {
                        using (SmtpClient client = new SmtpClient())
                        {
                            client.Host = SMTPHost;
                            client.Port = SMTPPort;
                            //client.EnableSsl = true;
                            client.Credentials = CredentialCache.DefaultNetworkCredentials;
#if (DEBUG)
                            client.Credentials = new NetworkCredential("ochekmez", "Arcelor1", "ask-ad");
#endif
                            client.DeliveryMethod = SmtpDeliveryMethod.Network;
                            client.Send(mail);

                            client.ServicePoint.CloseConnectionGroup(client.ServicePoint.ConnectionName);
                        }
                    }
                    finally
                    {
                        foreach (Attachment attachment in mail.Attachments)
                        {
                            attachment.Dispose();
                        }
                        mail.Attachments.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not send email with label template. Error: " + ex.ToString(), EventLogEntryType.Error, 4);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Email of the prepared label
        /// </summary>
        public bool EmailTemplate(PrintJobProps jobProps)
        {
            Boolean boolEmailLabel = false;
            if (PrepareTemplate(jobProps, true))
            {
                boolEmailLabel = EmailPDF(jobProps.CommandRule);
            }
            else
            {
                SenderMonitorEvent.sendMonitorEvent(eventLog, "Can not convert label template to pdf. Process failed", EventLogEntryType.Error, 4);
                return false;
            }

            return boolEmailLabel;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                eventLog = null;
            }
        }
    }
}
