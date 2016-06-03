using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
//using ios = System.Runtime.InteropServices;
//using Aspose.Cells;
//using Aspose.Cells.Rendering;

namespace PrintWindowsService
{
	public partial class PrintService : ServiceBase
	{
		private PrintJobs pJobs;

        #region Constructor

        public PrintService()
		{
			InitializeComponent();

            // Set up a timer to trigger every print task frequency.
            pJobs = new PrintJobs();
        }

        #endregion

        #region Methods

        protected override void OnStart(string[] args)
		{
            pJobs.StartJob();
        }

		protected override void OnStop()
		{
            pJobs.StopJob();
        }
        #endregion
    }
}

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
