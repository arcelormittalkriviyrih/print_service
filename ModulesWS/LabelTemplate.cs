using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PrintWindowsService
{
    /// <summary>
    /// Class for work with Excel template and data of label
    /// </summary>
    public class LabelTemplate
    {
        private SpreadsheetDocument spreadSheet = null;
        private WorkbookPart workbookpart = null;
        private Sheet worksheetParams = null;
        private WorksheetPart worksheetPartParams = null; 

        /// <summary>	Constructor. </summary>
        ///
        /// <param name="templateName">	Name of the template. </param>
        public LabelTemplate(string templateName)
        {
            spreadSheet = SpreadsheetDocument.Open(templateName, true);
            workbookpart = spreadSheet.WorkbookPart;
            worksheetParams = workbookpart.Workbook.Descendants<Sheet>().First(s => (s.SheetId == "2"));
            worksheetPartParams = (WorksheetPart)(workbookpart.GetPartById(worksheetParams.Id));
        }

        /// <summary>	Inserts a shared string item. </summary>
        ///
        /// <param name="text">			  	The text. </param>
        /// <param name="shareStringPart">	The share string part. </param>
        ///
        /// <returns>	An int. </returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>	Check empty cell. </summary>
        ///
        /// <param name="currentRow">	The current row. </param>
        /// <param name="afterCell"> 	The after cell. </param>
        /// <param name="columnCell">	The column cell. </param>
        ///
        /// <returns>	A Cell. </returns>
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

        /// <summary>	Gets cell value. </summary>
        ///
        /// <param name="cell">	The cell. </param>
        ///
        /// <returns>	The cell value. </returns>
        private string GetCellValue(Cell cell)
        {
            string resultValue = string.Empty;

            if (cell != null)
            {
                resultValue = cell.InnerText;

                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
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

        /// <summary>	Calculates the reference cell values. </summary>
        private void RecalcRefCellValues()
        {
            WorksheetPart wsPartFirst = (WorksheetPart)(workbookpart.GetPartById(workbookpart.Workbook.Descendants<Sheet>().First(s => (s.SheetId == "1")).Id));
            foreach (Cell refCell in wsPartFirst.Worksheet.Descendants<Cell>())
            {
                if ((refCell.DataType == null) & (refCell.CellFormula != null))
                {
                    refCell.CellValue.Remove();
                }
            }
            workbookpart.Workbook.CalculationProperties.ForceFullCalculation = true;
            workbookpart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            wsPartFirst.Worksheet.Save();
        }

        /// <summary>
        /// Fill data sheet of parameters
        /// </summary>
        public void FillParamValues(PrintJobProps jobProps)
        {
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

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
                    worksheetPartParams.Worksheet.Save();
                    Cell refCellC = CheckEmptyCell(rowParam, refCellB, "C");
                    worksheetPartParams.Worksheet.Save();
                    Cell refCellD = CheckEmptyCell(rowParam, refCellC, "D");
                    worksheetPartParams.Worksheet.Save();

                    /*if (rowParam.RowIndex == 1)
                    {
                        //first row for quantity
                        refCellC.CellValue = new CellValue(aJobProps.PrintQuantity);
                        refCellC.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }
                    else
                    {*/
                    //these rows for other params
                    int index = InsertSharedStringItem(jobProps.getLabelParameter(GetCellValue(refCellA), GetCellValue(refCellB)), shareStringPart);
                    refCellD.CellValue = new CellValue(index.ToString());
                    refCellD.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    //}
                }
            }
            RecalcRefCellValues();
            workbookpart.Workbook.Save();
            spreadSheet.Close();
        }
    }
}
