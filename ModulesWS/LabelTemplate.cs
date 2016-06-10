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

        public LabelTemplate(string aTemplateName)
        {
            spreadSheet = SpreadsheetDocument.Open(aTemplateName, true);
            workbookpart = spreadSheet.WorkbookPart;
            worksheetParams = workbookpart.Workbook.Descendants<Sheet>().First(s => (s.SheetId == "2"));
            worksheetPartParams = (WorksheetPart)(workbookpart.GetPartById(worksheetParams.Id));
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

        //delete all cache
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
        public void FillParamValues(PrintJobProps aJobProps)
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
                    Cell refCellD = CheckEmptyCell(rowParam, refCellC, "D");

                    /*if (rowParam.RowIndex == 1)
                    {
                        //first row for quantity
                        refCellC.CellValue = new CellValue(aJobProps.PrintQuantity);
                        refCellC.DataType = new EnumValue<CellValues>(CellValues.Number);
                    }
                    else
                    {*/
                        //these rows for other params
                        refCellD.CellValue = new CellValue(aJobProps.getLabelParameter(GetCellValue(refCellA), GetCellValue(refCellB)));
                        refCellD.DataType = new EnumValue<CellValues>(CellValues.String);
                    //}
                }
            }
            worksheetPartParams.Worksheet.Save();
            RecalcRefCellValues();
            workbookpart.Workbook.Save();
            spreadSheet.Close();
        }
    }
}
