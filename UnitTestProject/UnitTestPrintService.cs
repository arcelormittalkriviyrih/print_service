using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PrintWindowsService;
using System.Linq;
using System.Collections.Generic;

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
                                cell.DataType = new EnumValue<CellValues>(CellValues.string);

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

            /*System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo(@"D:\test.bmp");
            info.Arguments = "\"Bullzip PDF Printer\"";
            info.CreateNoWindow = true;
            info.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            info.UseShellExecute = true;
            info.Verb = "PrintTo";
            System.Diagnostics.Process.Start(info);*/

            PrintJobs pJobTest = new PrintJobs();
            pJobTest.OnPrintTimer(this, new EventArgs() as System.Timers.ElapsedEventArgs);

            /*List<jobPropsWS> JobData = new List<jobPropsWS>();
            ServicedbData lDbData = new ServicedbData("http://mssql2014srv/odata_unified_svc/api/Dynamic/");
            lDbData.fillJobData(ref JobData);*/

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

/*
в сервисе не работает System.Drawing.Printing 
System.Drawing.Printing.PrintDocument pd = new System.Drawing.Printing.PrintDocument();

//string pkInstalledPrinters = "";
//System.Drawing.Printing.PrinterSettings.StringCollection sc = System.Drawing.Printing.PrinterSettings.InstalledPrinters;
//for (int i = 0; i < sc.Count; i++)
//{
//    pkInstalledPrinters += "\n" + sc[i];
//}

pd.PrinterSettings.PrinterName = toPrinterName;

if (!pd.PrinterSettings.IsValid)
{
    vpEventLog.WriteEntry(string.Format("Printer {0} is not valid", toPrinterName));
    return false;
}
*/

/*
private void pd_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ev)
{
Bitmap image = new Bitmap(@"D:\1.bmp");

ev.Graphics.DrawImage(image, 0, 0);
ev.HasMorePages = false;
}*/

//WsFirst.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperUser;
//System.Drawing.Printing.PrintDocument pd = new System.Drawing.Printing.PrintDocument();
//pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(pd_PrintPage);
// Specify the printer to use.
//pd.PrinterSettings.PrinterName = toPrinterName;
//pd.Print();

//эта область не всегда совпадает с областью реальных данных
//Excel.Range lRange = lWs.UsedRange; //lWs.Range["A1:E14"];
//lRange.PrintOutEx(Type.Missing, Type.Missing, 1, Type.Missing, toPrinterName, Type.Missing);

// пример на Aspose.Cell
//            string designerFile = @"D:\template.xls";
//            Workbook workbook = new Workbook(designerFile);
//            Worksheet sheet = workbook.Worksheets[0];
//            sheet.SelectRange(1, 1, 5, 14, false);
//            workbook.Save(@"D:\1.tiff", SaveFormat.TIFF);
//второй вариант
//ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
//Specify the image format
//imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
//Only one page for the whole sheet would be rendered
//imgOptions.OnePagePerSheet = true;

//Render the sheet with respect to specified image/print options
//SheetRender sr = new SheetRender(sheet, imgOptions);
//Render the image for the sheet
//Bitmap bitmap = sr.ToImage(0);

//Save the image file specifying its image format.
//bitmap.Save(@"d:\1.bmp");


//Process.Start(@"D:\TMP\printExcel.exe");

/*
public static class myPrinters
{
    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool SetDefaultPrinter(string Name);

    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int pcchBuffer);

    private const int ERROR_FILE_NOT_FOUND = 2;

    private const int ERROR_INSUFFICIENT_BUFFER = 122;

    public static string getDefaultPrinter()
    {

        int pcchBuffer = 0;
        if (GetDefaultPrinter(null, ref pcchBuffer))
        {
            return null;
        }
        int lastWin32Error = Marshal.GetLastWin32Error();
        if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
        {
            StringBuilder pszBuffer = new StringBuilder(pcchBuffer);
            if (GetDefaultPrinter(pszBuffer, ref pcchBuffer))
            {
                return pszBuffer.ToString();
            }
            lastWin32Error = Marshal.GetLastWin32Error();
        }
        if (lastWin32Error == ERROR_FILE_NOT_FOUND)
        {
            return null;
        }
        throw new Exception("Marshal.GetLastWin32Error()");

    }
}
*/

/*
            SqlConnection dbConnection = new SqlConnection(dbConnectionString);

            try
            {
                dbConnection.Open();

                SqlCommand selectCommandProdResponse = new SqlCommand("SELECT ID, ResponseState, ProductionRequestID, EquipmentID, EquipmentClassID, ProductSegmentID, ProcessSegmentID\n" +
                      "FROM v_ProductionResponse\n" +
                      "WHERE (ResponseState = @State)\n" +
                      "  AND (EquipmentClassID = @EquipmentClassID)", dbConnection);
                selectCommandProdResponse.Parameters.AddWithValue("@State", "ToPrint");
                selectCommandProdResponse.Parameters.AddWithValue("@EquipmentClassID", "/2/");

                SqlCommand selectLabelProperty = new SqlCommand("SELECT TypeProperty, ClassPropertyID, ValueProperty\n" +
                      "FROM v_PrintProperties\n" +
                      "WHERE (ProductionResponse = @ProductionResponse)", dbConnection);
                selectLabelProperty.Parameters.AddWithValue("@ProductionResponse", null);

                SqlCommand CommandUpdateStatus = new SqlCommand("BEGIN TRANSACTION T1; UPDATE ProductionResponse SET ResponseState = @State WHERE ID = @ProductionResponseID; COMMIT TRANSACTION T1;", dbConnection);
                CommandUpdateStatus.Parameters.AddWithValue("@State", null);
                CommandUpdateStatus.Parameters.AddWithValue("@ProductionResponseID", null);

                SqlCommand selectCommandFiles = new SqlCommand("SELECT pf.Data\n" +
                      "FROM v_ProductionParameter_Files pf\n" +
                      "WHERE pf.ProductSegmentID = @ProductSegmentID\n" +
                      "  AND pf.ProcessSegmentID = @ProcessSegmentID\n" +
                      "  AND pf.PropertyType = @PropertyType"
                      //"  AND pf.FileType = @FileType" ???
                      , dbConnection);
                selectCommandFiles.Parameters.AddWithValue("@ProductSegmentID", null);
                selectCommandFiles.Parameters.AddWithValue("@ProcessSegmentID", null);
                selectCommandFiles.Parameters.AddWithValue("@PropertyType", 1);
                //selectCommandFiles.Parameters.AddWithValue("@FileType", null);
                tableLabelProperty = new DataTable();

                string QuantityParam = "", ToPrinterName = "", IpAddress = "", printState = "";
                
                using (SqlDataReader dbReaderProdResponse = selectCommandProdResponse.ExecuteReader())
                {
                    while (dbReaderProdResponse.Read())
                    {
                        //чтение параметров для шаблона и печати
                        selectLabelProperty.Parameters["@ProductionResponse"].Value = dbReaderProdResponse["ID"];
                        tableLabelProperty.Clear();
                        using (SqlDataAdapter adapterLabelProp = new SqlDataAdapter(selectLabelProperty))
                        {
                            adapterLabelProp.Fill(tableLabelProperty);
                        }
                        printLabel.tableLabelProperty = tableLabelProperty;

                        QuantityParam = GetParamaterFromDb("Weight", "0");
                        ToPrinterName = GetParamaterFromDb("EquipmentProperty", "2");
                        IpAddress = GetParamaterFromDb("EquipmentProperty", "3");

                        //чтение шаблона для печати этикетки
                        selectCommandFiles.Parameters["@ProductSegmentID"].Value = dbReaderProdResponse["ProductSegmentID"];
                        selectCommandFiles.Parameters["@ProcessSegmentID"].Value = dbReaderProdResponse["ProcessSegmentID"];
                        byte[] XlFile;
                        using (SqlDataReader dbReaderFiles = selectCommandFiles.ExecuteReader())
                        {
                            dbReaderFiles.Read();
                            XlFile = (byte[])dbReaderFiles["Data"];
                            dbReaderFiles.Close();
                        }

                        if (XlFile.Length > 0)
                        {
                            using (FileStream fs = new FileStream(printLabel.templateFile, FileMode.Create))
                            {
                                fs.Write(XlFile, 0, XlFile.Length);
                                fs.Close();
                            }

                            if (printLabel.printTemplate(ToPrinterName, IpAddress, QuantityParam))
                            {
                                printState = "Printed";
                            }
                            else
                            {
                                printState = "Failed";
                            }
                            SenderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("ProductionResponseID: {0}. Print to: {1}. Status: {2}", dbReaderProdResponse["ID"], ToPrinterName, printState), printState == "Failed"? EventLogEntryType.FailureAudit : EventLogEntryType.SuccessAudit);
                        }
                        else
                        {
                            printState = "Failed";
                            SenderMonitorEvent.sendMonitorEvent(vpEventLog, "Excel template is empty", EventLogEntryType.Error);
                        }

                        CommandUpdateStatus.Parameters["@ProductionResponseID"].Value = dbReaderProdResponse["ID"];
                        CommandUpdateStatus.Parameters["@State"].Value = printState;
                        CommandUpdateStatus.ExecuteNonQuery();
                        RequestCount++;
                    }
                    dbReaderProdResponse.Close();
                }
            }
            catch (Exception ex)
            {
                SenderMonitorEvent.sendMonitorEvent(vpEventLog, "Get data from DB. Error: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                dbConnection.Close();
            }
            SenderMonitorEvent.sendMonitorEvent(vpEventLog, string.Format("Print is done. {0} tasks", RequestCount), EventLogEntryType.Information);
*/

/*
        //чтение параметров из бд
        public string GetParamaterFromDb(string aTypeProperty, string aClassPropertyID)
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
*/

/*
//конвертация в монохром
public static Bitmap BitmapTo1Bpp(Bitmap img)
{
    int w = img.Width;
    int h = img.Height;
    Bitmap bmp = new Bitmap(w, h, PixelFormat.Format1bppIndexed);
    bmp.SetResolution(300, 300);
    BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format1bppIndexed);
    bmp.UnlockBits(data);
    for (int y = 0; y < h; y++)
    {
        byte[] scan = new byte[(w + 7) / 8];
        for (int x = 0; x < w; x++)
        {
            Color c = img.GetPixel(x, y);
            if (c.GetBrightness() >= 0.8) scan[x / 8] |= (byte)(0x80 >> (x % 8));
        }
        Marshal.Copy(scan, 0, (IntPtr)((int)data.Scan0 + data.Stride * y), scan.Length);
    }
    return bmp;
}

//печать изображения на ZPL
public void PrintBmp ()
{
    string bitmapFilePath = @"D:\image2.bmp";//@"c:\birka.bmp";  // file is attached to this support article
    string bitmapFilePathRotate = @"c:\birka1.bmp";  // file is attached to this support article
    Bitmap bmpSrc = new Bitmap(Image.FromFile(bitmapFilePath));
    bmpSrc.RotateFlip(RotateFlipType.Rotate180FlipX);
    bmpSrc = BitmapTo1Bpp(bmpSrc);
    bmpSrc.Save(bitmapFilePathRotate, ImageFormat.Bmp);
    byte[] bitmapFileData = System.IO.File.ReadAllBytes(bitmapFilePathRotate);
    int fileSize = bitmapFileData.Length;

    // The following is known about test.bmp.  It is up to the developer
    // to determine this information for bitmaps besides the given test.bmp.
    int bitmapDataOffset = 62;
    int width = bmpSrc.Width;
    int height = bmpSrc.Height;
    // Monochrome image required!
    double widthInBytes = Math.Ceiling(width / 8.0);
    int bitmapDataLength = fileSize - bitmapDataOffset;//(int)widthInBytes * height;

    // Copy over the actual bitmap data from the bitmap file.
    // This represents the bitmap data without the header information.
    byte[] bitmap = new byte[bitmapDataLength];
    Buffer.BlockCopy(bitmapFileData, bitmapDataOffset, bitmap, 0, bitmapDataLength);

    // Invert bitmap colors
    for (int i = 0; i < bitmapDataLength; i++)
    {
        bitmap[i] ^= 0xFF;
    }

    // Create ASCII ZPL string of hexadecimal bitmap data
    string ZPLImageDataString = BitConverter.ToString(bitmap);
    ZPLImageDataString = ZPLImageDataString.Replace("-", string.Empty);

    // Create ZPL command to print image
    string[] ZPLCommand = new string[4];

    ZPLCommand[0] = "^XA";
    ZPLCommand[1] = "^FO0,0";
    ZPLCommand[2] =
        "^GFA, " +
        bitmapDataLength.ToString() + "," +
        bitmapDataLength.ToString() + "," +
        (widthInBytes+1).ToString() + "," +
        ZPLImageDataString;

    ZPLCommand[3] = "^XZ";

    // Connect to printer
    string ipAddress = "192.168.100.160";
    int port = 9100;
    System.Net.Sockets.TcpClient client =
        new System.Net.Sockets.TcpClient();
    client.Connect(ipAddress, port);
    System.Net.Sockets.NetworkStream stream = client.GetStream();
    System.IO.StreamWriter mystreamwriter = new System.IO.StreamWriter(stream);

    // Send command strings to printer
    foreach (string commandLine in ZPLCommand)
    {
        mystreamwriter.WriteLine(commandLine);
        mystreamwriter.Flush();
    }
    client.Close();
}

//подгонка изображения к нужной ширине
private void scaleBitmap(Bitmap dest, Bitmap src)
{
    Rectangle srcRect = new Rectangle();
    Rectangle destRect = new Rectangle();

    destRect.Width = dest.Width;
    destRect.Height = dest.Height;
    using (Graphics g = Graphics.FromImage(dest))
    {
        Color backgroundColor = Color.White;
        Brush b = new SolidBrush(backgroundColor);
        g.FillRectangle(b, destRect);
        srcRect.Width = src.Width;
        srcRect.Height = src.Height;
        float sourceAspect = (float)src.Width / (float)src.Height;
        float destAspect = (float)dest.Width / (float)dest.Height;
        if (sourceAspect > destAspect)
        {
            // wider than high heep the width and scale the height
            destRect.Width = dest.Width;
            destRect.Height = (int)((float)dest.Width / sourceAspect);
            destRect.X = 0;
            destRect.Y = (dest.Height - destRect.Height) / 2;
        }
        else
        {
            // higher than wide – keep the height and scale the width
            destRect.Height = dest.Height;
            destRect.Width = (int)((float)dest.Height * sourceAspect);
            destRect.X = (dest.Width - destRect.Width) / 2;
            destRect.Y = 0;
        }
        g.DrawImage(src, destRect, srcRect, System.Drawing.GraphicsUnit.Pixel);
    }
}

//сохранение области в изображение
public void ExportRangeAsBmp()
{
    Excel.Application xl = new Excel.Application();

    xl.UserControl = true;

    //обход ошибки неверной версии длл и совместимости типов
    System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
    xl.Workbooks.Add(@"D:\template.xls");//(@"D:\text.xlsx");
    Excel.Worksheet lWs = (Excel.Worksheet)xl.ActiveWorkbook.ActiveSheet; // get_Item(1); //(Excel.Worksheet)lWb.ActiveSheet; //
    lWs.Protect(Contents: false);
    Excel.Range lRange = lWs.UsedRange; //lWs.Range["A1:E14"];
    lRange.CopyPicture(Excel.XlPictureAppearance.xlPrinter);

    // пример на Aspose.Cell
    //            string designerFile = @"D:\text.xlsx";
    //            Workbook workbook = new Workbook(designerFile);
    //            Worksheet sheet = workbook.Worksheets[0];
    //            sheet.SelectRange(1, 1, 5, 14, false);
    //            workbook.Save(@"D:\text.tiff", SaveFormat.TIFF);

    Bitmap image = new Bitmap(Clipboard.GetImage());
    image.SetResolution(300, 300);

    string bitmapFilePathCorrect = @"D:\image2.bmp";
    image = BitmapTo1Bpp(image);

    Clipboard.Clear();
    xl.ActiveWorkbook.Close(false);
    xl.Quit();
    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
    //xl.DisplayAlerts = true;

    double WidthInByte = image.Width / 8.0;
    double WidthInByteRound = Math.Ceiling(image.Width / 8.0);
    int bmpWidthCorrect = WidthInByte == WidthInByteRound ? image.Width : (int)(Math.Ceiling(image.Width / 8.0) + 1) * 8;

    image.Save(bitmapFilePathCorrect, ImageFormat.Bmp);
    if (bmpWidthCorrect != image.Width)
    {
        Bitmap imageCorrect = new Bitmap(bmpWidthCorrect, image.Height);
        imageCorrect.SetResolution(300, 300);
        scaleBitmap(imageCorrect, image);
        imageCorrect = BitmapTo1Bpp(imageCorrect);
        imageCorrect.Save(bitmapFilePathCorrect, ImageFormat.Bmp);
    }
}
*/
