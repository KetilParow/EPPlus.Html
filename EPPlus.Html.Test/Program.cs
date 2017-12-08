using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Html.Test
{
    class Program
    {
        private static string CurrentLocation = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        

        static void Main(string[] args)
        {
            FileInfo testFile = new FileInfo(CurrentLocation + "/Resources/Test002.xlsx");
            var package = new ExcelPackage(testFile);
            var worksheet = package.Workbook.Worksheets[1];

            string html = worksheet.ToHtml(ConversionFlags.AutoWidths);

            Show(html);
        }

        static void Show(string html)
        {
            string tmpFile = Path.GetTempFileName() + ".html";
            File.WriteAllText(tmpFile, html);
            Process.Start(tmpFile);
        }
    }
}
