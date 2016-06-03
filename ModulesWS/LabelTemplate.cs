using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for work with Excel template and data of label
    /// </summary>
    public class LabelTemplate
    {
        private SpreadsheetDocument spreadSheet;
        private WorksheetPart worksheetPartParams; 
        private WorkbookPart workbookpart;
        public static EventLog vpEventLog;
        //private Sheet worksheetParams;


        public LabelTemplate(string aTemplateName)
        {
            spreadSheet = SpreadsheetDocument.Open(aTemplateName, true);
            workbookpart = spreadSheet.WorkbookPart;
            //worksheetParams = (Sheet)workbookpart.Workbook.Sheets.FirstOrDefault();
            string id = workbookpart.Workbook.Descendants<Sheet>().First(s => (s.SheetId == "2")).Id;
            //  worksheetPartParams = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(worksheetParams.Id);
            worksheetPartParams = (WorksheetPart)(workbookpart.GetPartById(id));
        }

        ~LabelTemplate()
        {
            try
            {
                spreadSheet.Dispose();
            }
            catch (Exception ex)
            {
                senderMonitorEvent.sendMonitorEvent(vpEventLog, "Spread Sheet dispose failed: " + ex.ToString(), EventLogEntryType.Error);                
            }            
        }

        private Cell CheckEmptyCell(Row currentRow, Cell afterCell, string columnCell)
        {
            Cell returnCell = currentRow.Elements<Cell>().Where(c => c.CellReference.Value == columnCell + currentRow.RowIndex).FirstOrDefault();
            if (returnCell == null)
            {
                returnCell = new Cell() { CellReference = columnCell + currentRow.RowIndex };
                currentRow.InsertAfter(returnCell, afterCell);
            }
            return returnCell;
        }

        private string GetCellValue(Cell aCell)
        {
            string resultValue = "";

            if (aCell != null)
            {
                resultValue = aCell.InnerText;

                if (aCell.DataType != null)
                {
                    switch (aCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable = workbookpart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                            if (stringTable != null)
                            {
                                resultValue = stringTable.SharedStringTable.ElementAt(int.Parse(resultValue)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (resultValue)
                            {
                                case "0":
                                    resultValue = "FALSE";
                                    break;
                                default:
                                    resultValue = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return resultValue;
        }

        /// <summary>
        /// Fill data sheet of parameters
        /// </summary>
        public void FillParamValues(jobPropsWS aJobProps)
        {
            foreach (Row rowParam in worksheetPartParams.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
            {
                Cell refCellA = rowParam.Elements<Cell>().Where(c => c.CellReference.Value == "A" + rowParam.RowIndex).FirstOrDefault();
                if (refCellA == null)
                {
                    break;
                }
                else
                {
                    Cell refCellB = CheckEmptyCell(rowParam, refCellA, "B");
                    Cell refCellC = CheckEmptyCell(rowParam, refCellB, "C");

                    if (rowParam.RowIndex == 1)
                    {
                        //first row for quantity
                        refCellC.CellValue = new CellValue(aJobProps.PrintQuantity);
                    }
                    else
                    {
                        //these rows for other params
                        refCellC.CellValue = new CellValue(aJobProps.getLabelParamater(GetCellValue(refCellA), int.Parse(GetCellValue(refCellB))));
                    }
                    refCellC.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            workbookpart.Workbook.Save();
            spreadSheet.Close();
        }
    }
}
