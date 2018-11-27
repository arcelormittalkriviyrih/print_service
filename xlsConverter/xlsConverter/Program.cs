using System;
using Spire.Xls;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace xlsConverter
{
	[Serializable]
    public class Program
    {
        static void Main(string[] args)
        {
			//string input_file = @"d:\test2.xlsx";
			//string output_file = @"d:\test.bmp";
			//Convert(input_file, output_file);
			if (args.Length == 2)
			{
				string input_file = args[0];
				string output_file = args[1];
				ConvertNoRotate(input_file, output_file);
			}
			else
			{
				Console.Error.WriteLine("Wrong amount of input parameters!");
				Console.Error.WriteLine("");
				Console.Error.WriteLine("Usage: xlsConverter.exe input_file output_file");
				Console.Error.WriteLine("Supported output formats: PDF, BMP, PNG, GIF, JPG, JPEG, TIFF");
			}
        }

        public static void ConvertNoRotate(string input_file, string output_file)
        {
            Convert(input_file, output_file, 300f, 300f, false);
        }

        public static void ConvertWithRotate(string input_file, string output_file)
        {
            Convert(input_file, output_file, 300f, 300f, true);
        }

        public static void Convert(string input_file, string output_file, float dpiX, float dpiY, bool rotate)
        {
            string output_temp_file = string.Empty;

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(input_file);
            Worksheet sheet = workbook.Worksheets[0];

            if (output_file.ToLower().EndsWith(".pdf"))
            {
                sheet.SaveToPdf(output_file);
            }
            else if (output_file.ToLower().EndsWith(".emf"))
            {
                int lastRow = 0;
                int lastColumn = 0;
                for (int i = sheet.LastRow - 1; i > 0; i--)
                {
                    if (!sheet.Rows[i].IsBlank)
                    {
                        lastRow = i;
                        break;
                    }
                }
                for (int j = sheet.LastColumn - 1; j > 0; j--)
                {
                    if (!sheet.Columns[j].IsBlank)
                    {
                        lastColumn = j;
                        break;
                    }
                }
                sheet.SaveToEMFImage(output_file, 1, 1, lastRow + 2, lastColumn + 2, System.Drawing.Imaging.EmfType.EmfPlusDual);
            }
            else
            {
                using (Image image = sheet.ParentWorkbook.SaveAsImages(0, dpiX, dpiY))
                {
                    using (Bitmap originalImage = new Bitmap(image))
                    {
                        using (Image croppedImage = AutoCrop(originalImage))
                        {
                            if (rotate)
                            {
                                croppedImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                            }
                            croppedImage.Save(output_file);
                        }
                    }
                }
            }
        }

		private static Image AutoCrop(Bitmap bmp)
		{
			if (Image.GetPixelFormatSize(bmp.PixelFormat) != 32)
				throw new InvalidOperationException("Autocrop currently only supports 32 bits per pixel images.");

			// Initialize variables
			var cropColor = Color.White;

			var bottom = 0;
			var left = bmp.Width; // Set the left crop point to the width so that the logic below will set the left value to the first non crop color pixel it comes across.
			var right = 0;
			var top = bmp.Height; // Set the top crop point to the height so that the logic below will set the top value to the first non crop color pixel it comes across.

			var bmpData = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);

			unsafe
			{
				var dataPtr = (byte*)bmpData.Scan0;

				for (var y = 0; y < bmp.Height; y++)
				{
					for (var x = 0; x < bmp.Width; x++)
					{
						var rgbPtr = dataPtr + (x * 4);

						var b = rgbPtr[0];
						var g = rgbPtr[1];
						var r = rgbPtr[2];
						var a = rgbPtr[3];

						// If any of the pixel RGBA values don't match and the crop color is not transparent, or if the crop color is transparent and the pixel A value is not transparent
						if ((cropColor.A > 0 && (b != cropColor.B || g != cropColor.G || r != cropColor.R || a != cropColor.A)) || (cropColor.A == 0 && a != 0))
						{
							if (x < left)
								left = x;

							if (x >= right)
								right = x + 1;

							if (y < top)
								top = y;

							if (y >= bottom)
								bottom = y + 1;
						}
					}

					dataPtr += bmpData.Stride;
				}
			}

			bmp.UnlockBits(bmpData);

			if (left < right && top < bottom)
				//return bmp.Clone(new Rectangle(left, top, right - left, bottom - top), bmp.PixelFormat);
				return bmp.Clone(new Rectangle(/*left*/0, /*top*/0, right, bottom), bmp.PixelFormat);

			return null; // Entire image should be cropped, so just return null
		}
    }

    
}
