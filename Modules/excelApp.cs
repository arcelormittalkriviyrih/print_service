using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for work with Excel and data of label
    /// </summary>
    public class ExcelApplication
    {
        private Excel.Application excelApp;
        /// <summary>
        /// The old value of the CultureInfo
        /// </summary>
        private System.Globalization.CultureInfo saveCI;
        /// <summary>
        /// The new value of the CultureInfo
        /// </summary>
        private System.Globalization.CultureInfo newCI;
        /// <summary>
        /// First sheet with a label info
        /// </summary>
        private Excel.Worksheet WsFirst;

        /// <summary>
        /// The current value of the CultureInfo
        /// </summary>
        public System.Globalization.CultureInfo currentCI
        {
            get { return newCI; }
        }

        public ExcelApplication()
        {
            excelApp = new Excel.Application();
            //обход ошибки: System.Runtime.InteropServices.COMException (0x80028018): Использован старый формат, либо библиотека имеет неверный тип
            saveCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            newCI = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = newCI;

            //xl.UserControl = false;
            excelApp.DisplayAlerts = false;
            //xl.Interactive = false;
            //xl.Visible = true;
        }

        ~ ExcelApplication()
        {
            if (excelApp.Workbooks.Count > 0)
            {
                excelApp.ActiveWorkbook.Close(false);
            }
            excelApp.Quit();
            System.Threading.Thread.CurrentThread.CurrentCulture = newCI;
            excelApp.DisplayAlerts = true;
            System.Threading.Thread.CurrentThread.CurrentCulture = saveCI;
        }

        /// <summary>
        /// Open template of label
        /// </summary>
        public void OpenTemplate(string aFileName)
        {
            excelApp.Workbooks.Add(aFileName);
            WsFirst = (Excel.Worksheet)excelApp.ActiveWorkbook.ActiveSheet;
        }

        /// <summary>
        /// Setup and print label sheet
        /// </summary>
        public void PrintLabelSheet(string aPrinterName)
        {
            excelApp.PrintCommunication = false;
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
            excelApp.PrintCommunication = true;
            WsFirst.PrintOutEx(1, 1, 1, Type.Missing, aPrinterName);
        }

        /// <summary>
        /// Return parameters sheet
        /// </summary>
        public Excel.Worksheet GetParamsSheet()
        {
            return (Excel.Worksheet)excelApp.Sheets.get_Item(2);
        }

        /// <summary>
        /// Close template of label
        /// </summary>
        public void CloseTemplate()
        {
            if (excelApp.Workbooks.Count > 0)
            {
                excelApp.ActiveWorkbook.Close(false);
            }
            WsFirst = null;
        }
    }
}
