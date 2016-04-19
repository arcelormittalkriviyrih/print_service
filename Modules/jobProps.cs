using System;
using System.Data;
using System.IO;

namespace PrintWindowsService
{
    /// <summary>
    /// Class of label for print
    /// </summary>
    public class jobProps
    {
        private int productionResponseID;
        private string printerName;
        private string ipAddress;
        private string printQuantity;
        private byte[] xlFile;
        private DataTable tableLabelProperty;

        /// <summary>
        /// Production response ID
        /// </summary>
        public int ProductionResponseID
        {
            get { return productionResponseID; }
        }
        /// <summary>
        /// Printer for print label
        /// </summary>
        public string PrinterName
        {
            get { return printerName; }
        }
        /// <summary>
        /// IP of printer
        /// </summary>
        public string IpAddress
        {
            get { return ipAddress; }
        }
        /// <summary>
        /// Quantity parameter of label
        /// </summary>
        public string PrintQuantity
        {
            get { return printQuantity; }
        }
        /// <summary>
        /// Is exists template of label
        /// </summary>
        public bool isExistsTemplate
        {
            get { return xlFile.Length > 0; }
        }

        public jobProps(int cProductionResponseID, byte[] cXlFile, DataTable cTableLabelProperty)
        {
            productionResponseID = cProductionResponseID;
            xlFile = cXlFile;
            tableLabelProperty = cTableLabelProperty;

            printerName = getLabelParamater("EquipmentProperty", "2");
            ipAddress = getLabelParamater("EquipmentProperty", "3");
            printQuantity = getLabelParamater("Weight", "0");
        }
        /// <summary>
        /// Prepare template for print
        /// </summary>
        public void prepareTemplate()
        {
            if (xlFile.Length > 0)
            {
                using (FileStream fs = new FileStream(printLabel.templateFile, FileMode.Create))
                {
                    fs.Write(xlFile, 0, xlFile.Length);
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// Return parameter value by TypeProperty and ClassPropertyID
        /// </summary>
        public string getLabelParamater(string aTypeProperty, string aClassPropertyID)
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
    }
}