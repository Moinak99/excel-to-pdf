using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using System;
using System.Drawing;
using System.IO;

namespace imageMani
{
    class TreatExcel
    {


        public TreatExcel(string inputFilePath)
        {


            string outputPath = Path.GetDirectoryName(inputFilePath) + @"\" + Path.GetFileNameWithoutExtension(inputFilePath) + "_converted";
            ExtractImageFromExcel(inputFilePath, outputPath);
        }


        public void ExtractImageFromExcel(string inputFilePath, string outputPath)
        {
            try
            {

                string imagePath = outputPath;

                // Create workbook from source file.
                Workbook workbook = new Workbook(inputFilePath);

                // Accessing the first worksheet
                Worksheet worksheet = workbook.Worksheets["1"];

                // Setting the print area with  desired range
                worksheet.PageSetup.PrintArea = "A1:M60";


                // Setting all margins as 0
                worksheet.PageSetup.LeftMargin = 0;
                worksheet.PageSetup.RightMargin = 0;
                worksheet.PageSetup.TopMargin = 0;
                worksheet.PageSetup.BottomMargin = 0;

                // Setting OnePagePerSheet option as true
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true,
                    ImageType = ImageType.Png,
                    HorizontalResolution = 200,
                    VerticalResolution = 200
                };


                // Taking the image of worksheet
                SheetRender sr = new SheetRender(worksheet, options);
                sr.ToImage(0, outputPath + ".png");

                ConvertImageToPdf(imagePath);

            }
            catch (Exception e)
            {
                Console.WriteLine("Try Again" + e.Message);
            }

        }

        public void ConvertImageToPdf(string outputPath)
        {
            try
            {
                //Creating Pdf obj
                var pdf = new Aspose.Pdf.Document();
                //Adding page to that pdf obj
                var pdfImageSection = pdf.Pages.Add();
                // getting image from location
                FileStream stream = new FileStream(outputPath + ".png", FileMode.Open);
                //operations
                Image img = new Bitmap(stream);
                MemoryStream ms = new MemoryStream();
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Seek(0, SeekOrigin.Begin);
                var image = new Aspose.Pdf.Image { ImageStream = ms };
                pdfImageSection.Paragraphs.Add(image);
                //finally saving the pdf
                pdf.Save(outputPath + ".pdf");
                Console.WriteLine("success!");
            }
            catch (Exception e)
            {
                Console.WriteLine("Try Again" + e.Message);
            }
            finally
            {
                Console.WriteLine("Press Any Key to terminate");
                Console.ReadKey();
            }
        }


    }
}
