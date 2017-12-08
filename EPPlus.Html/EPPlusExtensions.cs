using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using System;

namespace EPPlus.Html
{
    public static class EPPlusExtensions
    {
        public static string ToHtml(this ExcelWorksheet sheet, ConversionFlags flags = ConversionFlags.None)
        {
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;

            HtmlElement htmlTable = new HtmlElement("table");
            htmlTable.Attributes["cellspacing"] = 0;
            htmlTable.Styles["white-space"] = "nowrap";

            //render rows
            for (int row = 1; row <= lastRow; row++)
            {
                ExcelRow excelRow = sheet.Row(row);

                var test = excelRow.Style;
                HtmlElement htmlRow = htmlTable.AddChild("tr");
                htmlRow.Styles.Update(excelRow.ToCss(flags));
                int col = 1;
                while (col <= lastCol)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    HtmlElement mergedCell = null;
                    if (!flags.HasFlag(ConversionFlags.IgnoreColSpans) && excelCell.Merge)
                    {
                        mergedCell = htmlCell;
                        int spansTo = col;
                        while(sheet.Cells[row, 1, row, spansTo +1].Merge)
                        {
                            spansTo++;
                        }
                        col = spansTo;
                        htmlCell.Attributes.Add("colspan", spansTo);
                    }
                    
                    htmlCell.Content = excelCell.Text;
                    htmlCell.Styles.Update(excelCell.ToCss(flags));
                    col++;
                }
            }

            return htmlTable.ToString();
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }
    }

    [Flags]
    public enum ConversionFlags
    {
        None = 0,
        AutoWidths = 1,
        RightAlignNumerals = 2,
        IgnoreColSpans = 4,
        ExactPixelHeights = 8
    }

}

