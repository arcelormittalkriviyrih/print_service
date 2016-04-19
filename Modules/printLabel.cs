using System;
using System.Data;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for initialising of parameters label and printing of the set label
    /// </summary>
    public static class printLabel
    {
        public static int pingTimeoutInSeconds;
        public static EventLog vpEventLog;
        public static ExcelApplication xl;
        public static string templateFile;

        /// <summary>
        /// Printing of the prepared label
        /// </summary>
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
                xl.OpenTemplate(templateFile);//@"D:\template.xls");
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Can not open file. Error: " + ex.ToString(), EventLogEntryType.Error);
                return false;
            }

            //Excel.Worksheet WsFirst = (Excel.Worksheet)xl.excelApp.ActiveWorkbook.ActiveSheet; // get_Item(1); //(Excel.Worksheet)lWb.ActiveSheet; //

            Boolean boolPrintLabel = false;
            try
            {
                //количество всегда на второй закладке в ячейке C1
                Excel.Worksheet WsParams = xl.GetParamsSheet();
                Excel.Range FindParamValue = (Excel.Range)WsParams.Cells[1, 3];
                FindParamValue.Value = aJobProps.PrintQuantity;

                int iRow = 2;
                while (((Excel.Range)WsParams.Cells[iRow, 1]).Value != null)
                {
                    ((Excel.Range)WsParams.Cells[iRow, 3]).Value = aJobProps.getLabelParamater(((Excel.Range)WsParams.Cells[iRow, 1]).Value.ToString(), ((Excel.Range)WsParams.Cells[iRow, 2]).Value.ToString());
                    iRow++;
                }

                WsParams = null;
                FindParamValue = null;
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Parameters sheet is not found. Error: " + ex.ToString(), EventLogEntryType.Warning);
            }

            try
            {
                //myPrinters.SetDefaultPrinter(toPrinterName);
                xl.PrintLabelSheet(aJobProps.PrinterName);
                boolPrintLabel = true;
            }

            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Print еrror: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                xl.CloseTemplate();
            }

            return boolPrintLabel;
        }
    }
}
