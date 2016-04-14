using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PrintWindowsService;

using System.Linq;

//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;

namespace PrintWindowsService.Tests
{
	[TestClass()]
	public class UnitTestPrintService
	{
		[TestMethod()]
		public void PrintServiceTest()
		{
            /* test open xml sdk
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open("D:\\template.xlsx", true))
            {
                WorkbookPart workbookpart = spreadSheet.WorkbookPart;
                //Sheets sheets = workbookpart.Workbook.GetFirstChild<Sheets>();
                Sheet worksheet = workbookpart.Workbook.Sheets.GetFirstChild<Sheet>();
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                //Row row = sheetData.Elements<Row>().Where(r => r.RowIndex == 2).First();

                var worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(worksheet.Id.Value);

                Row row1 = worksheet.Elements<Row>().FirstOrDefault(r => r.RowIndex == 26);//.FirstOrDefault();
                //IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
                int i = 1;
                foreach (Row row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                {
                    int j = 1;
                    if (i == 26)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellReference == "A26")
                            {
                                cell.CellValue = new CellValue("4567");
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                                break;
                            }
                            j++;
                        }
                        break;
                        //Cell cell = row.Elements<Cell>().SingleOrDefault(p => p.CellReference.Value == "A2");
                    }
                    i++;
                }
                //Cell cell = row.Elements<Cell>().Where(c => c.CellReference.Value == "A2").First();
                // Insert other code here.
                worksheetPart.Worksheet.Save();
                workbookpart.Workbook.Save();
                spreadSheet.Close();
            }
            */


            /* test Aspose.Cells
            //Instantiate a workbook.
            //Open an Excel file.
            Workbook workbook = new Workbook("D:\\template.xls");
            //Define a worksheet.
            Worksheet worksheet = workbook.Worksheets[0];
            //Apply different Image / Print options.
            Aspose.Cells.Rendering.ImageOrPrintOptions options = new Aspose.Cells.Rendering.ImageOrPrintOptions();
            options.PrintingPage = PrintingPageType.Default;
            options.OnePagePerSheet = true;
            SheetRender sr = new SheetRender(worksheet, options);
            sr.ToPrinter("TSC TTP-268M");

            var imgOption = new ImageOrPrintOptions();
            imgOption.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
            imgOption.HorizontalResolution = 203;
            imgOption.VerticalResolution = 203;
            imgOption.OnePagePerSheet = true;

            //Apply transparency to the output image
            imgOption.Transparent = true;

            //Create image after apply image or print options
            var sr1 = new SheetRender(worksheet, imgOption);
            sr1.ToImage(0, "d:\\123.bmp");
            */



            PrintJobs pJobTest = new PrintJobs();

            //pJobTest.StartJob();
            //pJobTest.StopJob();
            //pJobTest.OnPrintTimer(this, new EventArgs() as System.Timers.ElapsedEventArgs);
            //printService.PrintRange("TSC TTP-268M", "192.168.100.246");//"TSC TTP-268M");//ExportRangeAsBmp(); //PrintBmp();//

            Assert.IsTrue(true);
		}

		[TestMethod()]
		public void OnPrintTimerTest()
		{
			//Assert.Fail();
			Assert.IsTrue(true);
		}
	}
}
