using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
//using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Security.Principal;
using System.Reflection;

namespace PrintWindowsService
{
    public class PrintJobs
    {
        #region Const

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

        int pingTimeoutInSeconds;
        Excel.Application xl = null;
        System.Globalization.CultureInfo oldCI, newCI;
        string tmpExcelFile;
        DataTable tableLabelProperty;

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
            //InitializeComponent();

            // Set up a timer to trigger every print task frequency.
            int printTaskFrequencyInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPrintTaskFrequencyName]);
            pingTimeoutInSeconds = int.Parse(System.Configuration.ConfigurationManager.AppSettings[cPingTimeoutName]);
            m_PrintTimer = new System.Timers.Timer();
            m_PrintTimer.Interval = printTaskFrequencyInSeconds * 1000; // seconds to milliseconds
            m_PrintTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnPrintTimer);
            tmpExcelFile = Path.GetTempPath() + "Label.xls"; //Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Label.xls";
            //vpEventLog.WriteEntry(string.Format("File - {0}", tmpExcelFile));
            vpEventLog.WriteEntry(string.Format("Print Task Frequncy = {0}", printTaskFrequencyInSeconds));
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
            /*try
            {// Присоединение к открытому приложению Excel (если оно открыто)
                xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                xl = new Excel.Application();// Если нет, то создаём новое приложение
            }*/
            try
            {
                if (xl == null)
                {
                    xl = new Excel.Application();
                }

                //обход ошибки: System.Runtime.InteropServices.COMException (0x80028018): Использован старый формат, либо библиотека имеет неверный тип
                oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                newCI = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentCulture = newCI;

                //xl.UserControl = false;
                xl.DisplayAlerts = false;
                //xl.Interactive = false;
                //xl.Visible = true;
            }
            catch (Exception ex)
            {
                vpEventLog.WriteEntry("Error of Excel start: " + ex.ToString(), EventLogEntryType.Error);
            }

            vpEventLog.WriteEntry("Starting print service...");

            m_PrintTimer.Start();

            vpEventLog.WriteEntry("Print service has been started");
            fJobStarted = true;
        }

        public void StopJob()
        {
            if (xl != null)
            {
                if (xl.Workbooks.Count > 0)
                {
                    xl.ActiveWorkbook.Close(false);
                }
                xl.Quit();
                System.Threading.Thread.CurrentThread.CurrentCulture = newCI;
                xl.DisplayAlerts = true;
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                //xl = null;
                //GC.GetTotalMemory(true);
            }

            vpEventLog.WriteEntry("Stopping print service...");

            //stop timers if working
            if (m_PrintTimer.Enabled)
                m_PrintTimer.Stop();

            vpEventLog.WriteEntry("Print service has been stopped");
            fJobStarted = false;
        }

        public void OnPrintTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            // TODO: Insert print logic here.
            vpEventLog.WriteEntry("Monitoring the print activity.", EventLogEntryType.Information);
            m_PrintTimer.Stop();
            int RequestCount = 0;

            //временно
            if (xl == null)
            {
                /*
                try
                {// Присоединение к открытому приложению Excel (если оно открыто)
                    xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    xl = new Excel.Application();// Если нет, то создаём новое приложение
                }
                */
                //работаем со своим экземпляром Excel
                xl = new Excel.Application();

                if (newCI == null)
                {
                    newCI = new System.Globalization.CultureInfo("en-US");
                }
                System.Threading.Thread.CurrentThread.CurrentCulture = newCI;
                xl.UserControl = false;
            }

            SqlConnection dbConnection = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings[cConnectionStringName].ConnectionString);
            //SqlTransaction dbTransactionRead;

            try
            {
                dbConnection.Open();
                //dbTransactionRead = dbConnection.BeginTransaction();

                SqlCommand selectCommandProdResponse = new SqlCommand("SELECT ID, ResponseState, ProductionRequestID, EquipmentID, EquipmentClassID, ProductSegmentID, ProcessSegmentID\n" +
                      "FROM v_ProductionResponse\n" +
                      "WHERE (ResponseState = @State)\n" +
                      "  AND (EquipmentClassID = @EquipmentClassID)", dbConnection);
                selectCommandProdResponse.Parameters.AddWithValue("@State", "ToPrint");
                selectCommandProdResponse.Parameters.AddWithValue("@EquipmentClassID", "/2/");

                /* количество и параметры принтера читаю из selectLabelProperty
                SqlCommand selectCommandLot = new SqlCommand("SELECT ID, Quantity, ProductionRequest\n" +
                      "FROM v_MaterialLot_Request\n" +
                      "WHERE (ProductionRequest = @ProductionRequestID)", dbConnection);
                selectCommandLot.Parameters.AddWithValue("@ProductionRequestID", null);
                */

                SqlCommand selectLabelProperty = new SqlCommand("SELECT TypeProperty, ClassPropertyID, ValueProperty\n" +
                      "FROM v_PrintProperties\n" +
                      "WHERE (ProductionResponse = @ProductionResponse)", dbConnection);
                selectLabelProperty.Parameters.AddWithValue("@ProductionResponse", null);

                SqlCommand CommandUpdateStatus = new SqlCommand("BEGIN TRANSACTION T1; UPDATE ProductionResponse SET ResponseState = @State WHERE ID = @ProductionResponseID; COMMIT TRANSACTION T1;", dbConnection);
                CommandUpdateStatus.Parameters.AddWithValue("@State", null);
                CommandUpdateStatus.Parameters.AddWithValue("@ProductionResponseID", null);

                /*все параметры в selectLabelProperty
                SqlCommand selectCommandPrinterProp = new SqlCommand("SELECT EquipmentProperty.ClassPropertyID, EquipmentProperty.Value\n" +
                      "FROM v_EquipmentProperty EquipmentProperty\n" +
                      "WHERE EquipmentProperty.EquipmentID = @EquipmentID\n" +
                      "  AND EquipmentProperty.ClassPropertyID IN (2, 3)", dbConnection);
                selectCommandPrinterProp.Parameters.AddWithValue("@EquipmentID", null);
                */

                SqlCommand selectCommandFiles = new SqlCommand("SELECT pf.Data\n" +
                      "FROM v_ProductionParameter_Files pf\n" +
                      "WHERE pf.ProductSegmentID = @ProductSegmentID\n" +
                      "  AND pf.ProcessSegmentID = @ProcessSegmentID\n" +
                      "  AND pf.PropertyType = @PropertyType"
                      //"  AND pf.FileType = @FileType" ???
                      , dbConnection);
                selectCommandFiles.Parameters.AddWithValue("@ProductSegmentID", null);
                selectCommandFiles.Parameters.AddWithValue("@ProcessSegmentID", null);
                selectCommandFiles.Parameters.AddWithValue("@PropertyType", 1);
                //selectCommandFiles.Parameters.AddWithValue("@FileType", null);
                tableLabelProperty = new DataTable();

                string QuantityParam = "", ToPrinterName = "", IpAddress = "", printState = "";
                
                using (SqlDataReader dbReaderProdResponse = selectCommandProdResponse.ExecuteReader())
                {
                    while (dbReaderProdResponse.Read())
                    {
                        /* количество и параметры принтера читаю из selectLabelProperty
                        //чтение количества для печати шаблона
                        selectCommandLot.Parameters["@ProductionRequestID"].Value = dbReaderProdResponse["ProductionRequestID"];
                        using (SqlDataReader dbReaderLot = selectCommandLot.ExecuteReader())
                        {
                            dbReaderLot.Read();
                            lnQuantity = dbReaderLot.GetInt32(1);
                            dbReaderLot.Close();
                        }

                        //чтение параметров принтера
                        selectCommandPrinterProp.Parameters["@EquipmentID"].Value = dbReaderProdResponse["EquipmentID"];
                        using (SqlDataReader dbReaderPrinterProp = selectCommandPrinterProp.ExecuteReader())
                        {
                            while (dbReaderPrinterProp.Read())
                            {
                                if (dbReaderPrinterProp.GetInt32(0) == 2)
                                {
                                    ToPrinterName = dbReaderPrinterProp.GetSqlString(1);
                                }
                                else
                                {
                                    IpAddress = dbReaderPrinterProp.GetSqlString(1);
                                }
                            }
                            dbReaderPrinterProp.Close();
                        }
                        */

                        //чтение параметров для шаблона и печати
                        selectLabelProperty.Parameters["@ProductionResponse"].Value = dbReaderProdResponse["ID"];
                        tableLabelProperty.Clear();
                        using (SqlDataAdapter adapterLabelProp = new SqlDataAdapter(selectLabelProperty))
                        {
                            adapterLabelProp.Fill(tableLabelProperty);
                        }

                        QuantityParam = GetParamaterFromDb("Weight", "0");
                        ToPrinterName = GetParamaterFromDb("EquipmentProperty", "2");
                        IpAddress = GetParamaterFromDb("EquipmentProperty", "3");

                        //чтение шаблона для печати этикетки
                        selectCommandFiles.Parameters["@ProductSegmentID"].Value = dbReaderProdResponse["ProductSegmentID"];
                        selectCommandFiles.Parameters["@ProcessSegmentID"].Value = dbReaderProdResponse["ProcessSegmentID"];
                        byte[] XlFile;
                        using (SqlDataReader dbReaderFiles = selectCommandFiles.ExecuteReader())
                        {
                            dbReaderFiles.Read();
                            XlFile = (byte[])dbReaderFiles["Data"];
                            dbReaderFiles.Close();
                        }

                        if (XlFile.Length > 0)
                        {
                            using (FileStream fs = new FileStream(tmpExcelFile, FileMode.Create))
                            {
                                fs.Write(XlFile, 0, XlFile.Length);
                                fs.Close();
                            }

                            if (PrintRange(ToPrinterName, IpAddress, QuantityParam))
                            {
                                printState = "Printed";
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            vpEventLog.WriteEntry(String.Format("ProductionResponseID: {0}. Print to: {1}. Status: {2}", dbReaderProdResponse["ID"], ToPrinterName, printState), printState == "Failed"? EventLogEntryType.FailureAudit : EventLogEntryType.SuccessAudit);
                        }
                        else
                        {
                            vpEventLog.WriteEntry("Excel template is empty", EventLogEntryType.Error);
                            printState = "Failed";
                        }

                        CommandUpdateStatus.Parameters["@ProductionResponseID"].Value = dbReaderProdResponse["ID"];
                        CommandUpdateStatus.Parameters["@State"].Value = printState;
                        //dbTransactionWrite = dbConnection.BeginTransaction();
                        //CommandUpdateStatus.Transaction = dbTransactionWrite;
                        CommandUpdateStatus.ExecuteNonQuery();
                        //dbTransactionRead.Commit();
                        RequestCount++;
                    }
                    dbReaderProdResponse.Close();
                }
            }
            catch (Exception ex)
            {
                vpEventLog.WriteEntry("Get data from DB. Error: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                dbConnection.Close();
            }
            vpEventLog.WriteEntry(string.Format("Print is done. {0} tasks", RequestCount));
            //временно
            //xl.Quit();
            //System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

            m_PrintTimer.Start();
        }

        /*
        public static class myPrinters
        {
            [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool SetDefaultPrinter(string Name);

            [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int pcchBuffer);

            private const int ERROR_FILE_NOT_FOUND = 2;

            private const int ERROR_INSUFFICIENT_BUFFER = 122;

            public static String getDefaultPrinter()
            {

                int pcchBuffer = 0;
                if (GetDefaultPrinter(null, ref pcchBuffer))
                {
                    return null;
                }
                int lastWin32Error = Marshal.GetLastWin32Error();
                if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
                {
                    StringBuilder pszBuffer = new StringBuilder(pcchBuffer);
                    if (GetDefaultPrinter(pszBuffer, ref pcchBuffer))
                    {
                        return pszBuffer.ToString();
                    }
                    lastWin32Error = Marshal.GetLastWin32Error();
                }
                if (lastWin32Error == ERROR_FILE_NOT_FOUND)
                {
                    return null;
                }
                throw new Exception("Marshal.GetLastWin32Error()");

            }
        }
        */

        //чтение параметров из бд
        public string GetParamaterFromDb(string aTypeProperty, string aClassPropertyID)
        {
            string ParamValue = "";

            DataRow[] foundRows;
            foundRows = tableLabelProperty.Select("TypeProperty = '" + aTypeProperty + "' AND ClassPropertyID = " + aClassPropertyID);
            if (foundRows.Length > 0)
            {
                ParamValue = foundRows[0]["ValueProperty"].ToString();
            }

            return ParamValue;
        }

        //печать области на заданный принтер
        public bool PrintRange(string toPrinterName, string IpAdress, string printQuantity)
        {
            //перед печатью если задан IP сделать пинг
            if ((pingTimeoutInSeconds > 0) & (IpAdress != ""))
            {
                System.Net.NetworkInformation.Ping printerPing = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply printerReply = printerPing.Send(IpAdress, pingTimeoutInSeconds);
                if (printerReply.Status != System.Net.NetworkInformation.IPStatus.Success)
                {
                    vpEventLog.WriteEntry(string.Format("Printer {0}  {1}  ping timeout status {2}", toPrinterName, IpAdress, printerReply.Status), EventLogEntryType.Warning);
                    return false;
                }
            }

            /*
            в сервисе не работает System.Drawing.Printing 
            System.Drawing.Printing.PrintDocument pd = new System.Drawing.Printing.PrintDocument();

            //String pkInstalledPrinters = "";
            //System.Drawing.Printing.PrinterSettings.StringCollection sc = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
            //for (int i = 0; i < sc.Count; i++)
            //{
            //    pkInstalledPrinters += "\n" + sc[i];
            //}

            pd.PrinterSettings.PrinterName = toPrinterName;

            if (!pd.PrinterSettings.IsValid)
            {
                vpEventLog.WriteEntry(string.Format("Printer {0} is not valid", toPrinterName));
                return false;
            }
            */

            //сделать вызов Excel один раз и закрыть с остановкой сервиса
            //Excel.Application xl = new Excel.Application();
            //xl.UserControl = true;
            //System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = newCI;
            try
            {
                xl.Workbooks.Add(tmpExcelFile);//@"D:\template.xls");
            }
            catch (Exception ex)
            {
                vpEventLog.WriteEntry("Can not open file. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            Excel.Worksheet WsFirst = (Excel.Worksheet)xl.ActiveWorkbook.ActiveSheet; // get_Item(1); //(Excel.Worksheet)lWb.ActiveSheet; //

            Excel.Range FindParamValue;
            Excel.Worksheet WsParams;
            Boolean boolPrintLabel = false;
            try
            {
                //количество всегда на второй закладке в ячейке A2
                WsParams = (Excel.Worksheet)xl.Sheets.get_Item(2);
                FindParamValue = (Excel.Range)WsParams.Cells[1, 3];
                FindParamValue.Value = printQuantity;

                int iRow = 2;
                while (((Excel.Range)WsParams.Cells[iRow, 1]).Value != null)
                {
                    ((Excel.Range)WsParams.Cells[iRow, 3]).Value = GetParamaterFromDb(((Excel.Range)WsParams.Cells[iRow, 1]).Value.ToString(), ((Excel.Range)WsParams.Cells[iRow, 2]).Value.ToString());
                    iRow++;
                }
            }
            catch (Exception ex)
            {
                vpEventLog.WriteEntry("Parameters sheet is not found. Error: " + ex.ToString(), EventLogEntryType.Warning);
            }

            /* установка значений свойств для этикетки
            foreach (DataRow rowProp in labelProp.Rows)
            {
                FindParamName = null;
                FindParamName = WsParams.get_Range("A1", "Z1").Find(rowProp[0], Type.Missing,
                                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false);

                if (FindParamName != null)
                {
                    FindParamValue = (Excel.Range)WsParams.Cells[2, FindParamName.Column];
                    FindParamValue.Value = rowProp[1];
                }
            }
            */

            /*Excel.Range editCell = lWs.get_Range("C3");
            editCell.Value = 123.45;*/
            //WsFirst.Protect(Contents: false);

            /*
            если нужна установки нужного размера страницы
            System.Drawing.Printing.PrintDocument pdoc = new System.Drawing.Printing.PrintDocument();
            pdoc.PrinterSettings.PrinterName = toPrinterName;
            WsFirst.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperUser; //???
            */

            // пример на Aspose.Cell
            //            string designerFile = @"D:\template.xls";
            //            Workbook workbook = new Workbook(designerFile);
            //            Worksheet sheet = workbook.Worksheets[0];
            //            sheet.SelectRange(1, 1, 5, 14, false);
            //            workbook.Save(@"D:\1.tiff", SaveFormat.TIFF);
            //второй вариант
            //ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            //Specify the image format
            //imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
            //Only one page for the whole sheet would be rendered
            //imgOptions.OnePagePerSheet = true;

            //Render the sheet with respect to specified image/print options
            //SheetRender sr = new SheetRender(sheet, imgOptions);
            //Render the image for the sheet
            //Bitmap bitmap = sr.ToImage(0);

            //Save the image file specifying its image format.
            //bitmap.Save(@"d:\1.bmp");


            //Process.Start(@"D:\TMP\printExcel.exe");

            try
            {
                //myPrinters.SetDefaultPrinter(toPrinterName);
                xl.PrintCommunication = false;
                WsFirst.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                WsFirst.PageSetup.CenterHorizontally = false;
                WsFirst.PageSetup.CenterVertically = false;
                WsFirst.PageSetup.LeftMargin = 0;
                WsFirst.PageSetup.RightMargin = 0;
                WsFirst.PageSetup.TopMargin = 0;
                WsFirst.PageSetup.BottomMargin = 0;
                WsFirst.PageSetup.HeaderMargin = 0;
                WsFirst.PageSetup.FooterMargin = 0;
                WsFirst.PageSetup.FitToPagesWide = 1;
                WsFirst.PageSetup.ScaleWithDocHeaderFooter = true;
                //WsFirst.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperUser;
                //System.Drawing.Printing.PrintDocument pd = new System.Drawing.Printing.PrintDocument();
                //pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(pd_PrintPage);
                // Specify the printer to use.
                //pd.PrinterSettings.PrinterName = toPrinterName;
                //pd.Print();
                xl.PrintCommunication = true;
                WsFirst.PrintOutEx(1, 1, 1, Type.Missing, toPrinterName);
                boolPrintLabel = true;
                //эта область не всегда совпадает с областью реальных данных
                //Excel.Range lRange = lWs.UsedRange; //lWs.Range["A1:E14"];
                //lRange.PrintOutEx(Type.Missing, Type.Missing, 1, Type.Missing, toPrinterName, Type.Missing);
            }

            /*xl.Quit();
            xl = null;
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;*/
            catch (Exception ex)
            {
                vpEventLog.WriteEntry("Print еrror: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                if (xl.Workbooks.Count > 0)
                {
                    xl.ActiveWorkbook.Close(false);
                }
                WsFirst = null;
                WsParams = null;
                FindParamValue = null;
            }

            return boolPrintLabel;
        }

        /*
        private void pd_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ev)
        {
            Bitmap image = new Bitmap(@"D:\1.bmp");

            ev.Graphics.DrawImage(image, 0, 0);
            ev.HasMorePages = false;
        }*/

        #endregion
    }
}
