using System;
using System.IO;
using System.Collections.Generic;

namespace PrintWindowsService
{
    /// <summary>
    /// Class of label for print
    /// </summary>
    public class PrintJobProps
    {
        private int jobOrderID;
        private string command;
        private string commandRule;
        private byte[] xlFile;
        private List<EquipmentPropertyValue> tableEquipmentProperty;
        private List<PrintPropertiesValue> tableLabelProperty;

        /// <summary>
        /// Job order ID
        /// </summary>
        public int JobOrderID
        {
            get { return jobOrderID; }
        }
        /// <summary>
        /// Job order command
        /// </summary>
        public string Command
        {
            get { return command; }
        }
        /// <summary>
        /// Job order command
        /// </summary>
        public string CommandRule
        {
            get { return commandRule; }
        }
        /// <summary>
        /// Printer name for print label
        /// </summary>
        public string PrinterName
        {
            get { return getEquipmentProperty("PRINTER_NAME"); }
        }
        /// <summary>
        /// IP of printer
        /// </summary>
        public string IpAddress
        {
            get { return getEquipmentProperty("PRINTER_IP"); }
        }
        /// <summary>
        /// Is exists template of label
        /// </summary>
        public bool isExistsTemplate
        {
            get { return xlFile.Length > 0; }
        }

        public PrintJobProps(int cJobOrderID,
                             string cCommand,
                             string cCommandRule,
                             byte[] cXlFile,
                             List<EquipmentPropertyValue> cTableEquipmentProperty,
                             List<PrintPropertiesValue> cTableLabelProperty)
        {
            jobOrderID = cJobOrderID;
            command = cCommand;
            commandRule = cCommandRule;
            xlFile = cXlFile;
            tableEquipmentProperty = cTableEquipmentProperty;
            tableLabelProperty = cTableLabelProperty;
        }
        /// <summary>
        /// Prepare template for print
        /// </summary>
        public void prepareTemplate(string ExcelTemplateFile)
        {
            if (xlFile.Length > 0)
            {
                using (FileStream fs = new FileStream(ExcelTemplateFile, FileMode.Create))
                {
                    fs.Write(xlFile, 0, xlFile.Length);
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// Return label parameter value by TypeProperty and PropertyCode
        /// </summary>
        public string getLabelParameter(string aTypeProperty, string aPropertyCode)
        {
            string ParamValue = "";

            PrintPropertiesValue propertyFind = tableLabelProperty.Find(x => (x.TypeProperty == aTypeProperty) & (x.PropertyCode == aPropertyCode));
            if (propertyFind != null)
            {
                ParamValue = propertyFind.Value;
            }

            return ParamValue;
        }
        /// <summary>
        /// Return equipment property value by Property
        /// </summary>
        public string getEquipmentProperty(string aProperty)
        {
            string ParamValue = "";

            EquipmentPropertyValue propertyFind = tableEquipmentProperty.Find(x => (x.Property == aProperty));
            if (propertyFind != null)
            {
                ParamValue = propertyFind.Value == null ? "" : propertyFind.Value.ToString();
            }

            return ParamValue;
        }
    }
}

namespace KEPServerSenderService
{
    /// <summary>
    /// Class of command properties
    /// </summary>
    public class SenderJobProps
    {
        private int productionResponseID;
        /// <summary>
        /// Production response ID
        /// </summary>
        public int ProductionResponseID
        {
            get { return productionResponseID; }
        }
        public SenderJobProps(int cProductionResponseID)
        {
            productionResponseID = cProductionResponseID;
        }
    }
}
