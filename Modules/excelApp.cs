using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace PrintWindowsService
{
    public class ExcelApplication
    {
        public Excel.Application excelApp;
        /// <summary>
        /// The old value of the CultureInfo
        /// </summary>
        private System.Globalization.CultureInfo saveCI;
        /// <summary>
        /// The new value of the CultureInfo
        /// </summary>
        private System.Globalization.CultureInfo newCI;

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
    }
}
