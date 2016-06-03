using System;
using System.Data;
using System.IO;
using System.Collections.Generic;

namespace PrintWindowsService
{
    /// <summary>
    /// Class of label for print
    /// </summary>
    public class jobPropsWS
    {
        private int productionResponseID;
        private string printerName;
        private string ipAddress;
        private string printQuantity;
        private byte[] xlFile;
        private List<PrintPropertiesValue> tableLabelProperty;

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

        public jobPropsWS(int cProductionResponseID, byte[] cXlFile, List<PrintPropertiesValue> cTableLabelProperty)
        {
            productionResponseID = cProductionResponseID;
            xlFile = cXlFile;
            tableLabelProperty = cTableLabelProperty;

            printerName = getLabelParamater("EquipmentProperty", 2);
            ipAddress = getLabelParamater("EquipmentProperty", 3);
            printQuantity = getLabelParamater("Weight", 0);
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
        public string getLabelParamater(string aTypeProperty, int aClassPropertyID)
        {
            string ParamValue = "";

            PrintPropertiesValue propertyFind = tableLabelProperty.Find(x => (x.TypeProperty == aTypeProperty) & (x.ClassPropertyID == aClassPropertyID));
            if (propertyFind != null)
            {
                ParamValue = propertyFind.ValueProperty;
            }

            return ParamValue;
        }
    }
}