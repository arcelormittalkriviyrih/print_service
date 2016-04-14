using System;
using System.Data;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace PrintWindowsService
{
    public static class printLabel
    {
        public static int pingTimeoutInSeconds;
        public static EventLog vpEventLog;
        public static ExcelApplication xl;
        public static string templateFile;

        //печать области на заданный принтер
        public static bool printTemplate(jobProps aJobProps)
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

            System.Threading.Thread.CurrentThread.CurrentCulture = xl.currentCI;
            try
            {
                xl.excelApp.Workbooks.Add(templateFile);//@"D:\template.xls");
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not open file. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            Excel.Worksheet WsFirst = (Excel.Worksheet)xl.excelApp.ActiveWorkbook.ActiveSheet; // get_Item(1); //(Excel.Worksheet)lWb.ActiveSheet; //

            Excel.Range FindParamValue;
            Excel.Worksheet WsParams;
            Boolean boolPrintLabel = false;
            try
            {
                //количество всегда на второй закладке в ячейке A2
                WsParams = (Excel.Worksheet)xl.excelApp.Sheets.get_Item(2);
                FindParamValue = (Excel.Range)WsParams.Cells[1, 3];
                FindParamValue.Value = aJobProps.PrintQuantity;

                int iRow = 2;
                while (((Excel.Range)WsParams.Cells[iRow, 1]).Value != null)
                {
                    ((Excel.Range)WsParams.Cells[iRow, 3]).Value = aJobProps.getLabelParamater(((Excel.Range)WsParams.Cells[iRow, 1]).Value.ToString(), ((Excel.Range)WsParams.Cells[iRow, 2]).Value.ToString());
                    iRow++;
                }
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Parameters sheet is not found. Error: " + ex.ToString(), EventLogEntryType.Warning);
            }

            try
            {
                //myPrinters.SetDefaultPrinter(toPrinterName);
                xl.excelApp.PrintCommunication = false;
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
                xl.excelApp.PrintCommunication = true;
                WsFirst.PrintOutEx(1, 1, 1, Type.Missing, aJobProps.PrinterName);
                boolPrintLabel = true;
            }

            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print еrror: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                if (xl.excelApp.Workbooks.Count > 0)
                {
                    xl.excelApp.ActiveWorkbook.Close(false);
                }
                WsFirst = null;
                WsParams = null;
                FindParamValue = null;
            }

            return boolPrintLabel;
        }
    }
}
