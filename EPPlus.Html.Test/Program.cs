using EPPlus.Html.Converters;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;

namespace EPPlus.Html.Test
{
    class Program
    {
        private static string CurrentLocation = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        

        static void Main(string[] args)
        {
            //FileInfo testFile = new FileInfo(CurrentLocation + "/Resources/Test003.xlsx");            
            //ExcelWorksheet worksheet = GetEmptytestTemplateAndGenerateSomeContent();

            ExcelWorksheet worksheet = GetExampleSheet();
            //var base64String = ExtractImage(worksheet);
            string html = worksheet.ToHtml(ConversionFlags.AutoWidths | ConversionFlags.AggregateStyleSheet);
            Show(html);
        }

        private static ExcelWorksheet GetExampleSheet()
        {

            ExcelWorkbook workbook = GetExcelResource("Avregning kundekonto.xlsx").Workbook;
            return workbook.Worksheets[1];
        }

        private static ExcelWorksheet GetEmptytestTemplateAndGenerateSomeContent()
        {
            ExcelPackage package = GetExcelResource("debank.xltx");
            var worksheet = package.Workbook.Worksheets[1];

            worksheet.SetValue("A1", "Headeren");

            for (var i = 1; i < 6; i++)
            {
                worksheet.SetValue(3, i, $"Kolonne {i}");
            }
            var rnd = new Random(DateTime.Now.Millisecond);
            for (var r = 4; r < 14; r++)
            {
                for (var i = 1; i < 6; i++)
                {
                    worksheet.SetValue(r, i, rnd.NextDouble());
                }
            }

            return worksheet;
        }

        private static string ExtractImage(ExcelWorksheet worksheet)
        {
            string base64String;
            using (var stream = new MemoryStream())
            {
                using (var drawing = (ExcelPicture)worksheet.Drawings[0])
                {
                    //ImageCodecInfo pngEncoder = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == ImageFormat.Png.Guid);
                    //drawing.Image.Save(stream, pngEncoder, null);

                    drawing.Image.Save(stream, ImageFormat.Png);
                    byte[] imageBytes = stream.ToArray();
                    base64String = Convert.ToBase64String(imageBytes);
                }
            }
            return base64String;
        }

        private static ExcelPackage GetExcelResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            ExcelPackage package = null;
            using (Stream stream = assembly.GetManifestResourceStream($"EPPlus.Html.Test.Resources.{resourceName}"))
            {
                package = new ExcelPackage(stream);
            }
            return package;
        }

        static void Show(string html)
        {
            string tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);
            Process.Start(tmpFile);
        }
    }
}
