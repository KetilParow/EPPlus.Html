using EPPlus.Html.Converters;
using EPPlus.Html.Html;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Html
{
    public static class EPPlusExtensions
    {
        public static string ToHtml(this ExcelWorksheet sheet, ConversionFlags flags = ConversionFlags.Default)
        {
            StringBuilder inlineStylesBuilder = flags.HasFlag(ConversionFlags.AggregateStyleSheet) ? new StringBuilder() : null;
            HtmlElement styleelement = null;

            var addedStyleForElements = new Dictionary<string, List<HtmlElement>>();
            if (flags.HasFlag(ConversionFlags.NoCss))
            {
                flags = flags | ConversionFlags.AutoWidths | ConversionFlags.AutoHeights;
            }
            int lastRow = sheet.Dimension.Rows;
            int lastCol = sheet.Dimension.Columns;

            HtmlElement scopedDiv = new HtmlElement("div");
            scopedDiv.Attributes.Add("id", "excelScope");

            if (inlineStylesBuilder != null)
            {
                styleelement = scopedDiv.AddChild("style");
                styleelement.Attributes.Add("scoped", null);
            }

            HtmlElement htmlTable = scopedDiv.AddChild("table");
            htmlTable.Attributes["cellspacing"] = 0;
            htmlTable.Styles["white-space"] = "nowrap";
            htmlTable.AddChild("thead");
            HtmlElement tBody = htmlTable.AddChild("tbody");

            //render rows
            for (int row = 1; row <= lastRow; row++)
            {
                ExcelRow excelRow = sheet.Row(row);
                HtmlElement htmlRow = tBody.AddChild("tr");
                htmlRow.Attributes.Add("er", row);
                htmlRow.Styles.Update(excelRow.ToCss(flags));

                HashStyles(inlineStylesBuilder, addedStyleForElements, htmlRow);
                int col = 1;
                while (col <= lastCol)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];
                    HtmlElement htmlCell = htmlRow.AddChild("td");
                    int ec = col;
                    htmlCell.Attributes.Add("ec", ec);

                    col = SpanIfMergedCells(sheet, flags, row, col, excelCell, htmlCell);

                    htmlCell.Content = excelCell.Text;

                    CssDeclaration cssDeclarations = excelCell.ToCss(flags);
                    CheckForAndAddBackgroundImages(row, ec, sheet, cssDeclarations);
                    htmlCell.Styles.Update(cssDeclarations);

                    HashStyles(inlineStylesBuilder, addedStyleForElements, htmlCell);

                    col++;
                }
            }

            SubstituteInlineStyles(inlineStylesBuilder, styleelement, addedStyleForElements);

            return scopedDiv.ToString();
        }

        private static int SpanIfMergedCells(ExcelWorksheet sheet, ConversionFlags flags, int row, int col, ExcelRange excelCell, HtmlElement htmlCell)
        {
            if (flags.HasFlag(ConversionFlags.IgnoreColSpans) || !excelCell.Merge)
            {
                return col;
            }
            int spansTo = col;
            while (sheet.Cells[row, col, row, spansTo].Merge)
            {
                spansTo++;
            }
            
            htmlCell.Attributes.Add("colspan", spansTo - col);
            
            //Should now be "next" column.
            return spansTo - 1;
        }

        private static void SubstituteInlineStyles(StringBuilder inlineStylesBuilder, HtmlElement styleelement, Dictionary<string, List<HtmlElement>> addedStyleForElements)
        {
            if (inlineStylesBuilder == null)
                return;

            foreach (var elementStyle in addedStyleForElements.Keys)
            {
                var hash = $"f-{Math.Abs(elementStyle.GetHashCode())}";
                inlineStylesBuilder.Append($"#excelScope .{hash} {{{elementStyle}}}\n");
                foreach (var element in addedStyleForElements[elementStyle])
                {
                    element.Styles.Clear();
                    element.Attributes.Add("class", $"{hash}");
                }
            }
            styleelement.Content = inlineStylesBuilder.ToString();
        }

        private static void HashStyles(StringBuilder inlineStylesBuilder, Dictionary<string, List<HtmlElement>> addedStyleForCells, HtmlElement htmlCell)
        {
            if (inlineStylesBuilder == null)
            {
                return;
            }
            if (!addedStyleForCells.ContainsKey(htmlCell.Styles.ToString()))
            {
                addedStyleForCells[htmlCell.Styles.ToString()] = new List<HtmlElement>();
            }
            addedStyleForCells[htmlCell.Styles.ToString()].Add(htmlCell);
        }

        private static void CheckForAndAddBackgroundImages(int row, int col, ExcelWorksheet sheet, CssDeclaration cssDeclarations)
        {
            foreach(var picture in sheet.AllDrawings())
            {
                if (Math.Max(1, picture.From.Row) != row) return;
                if (Math.Max(1, picture.From.Column) != col) return;
                cssDeclarations["background-image"] = "url('data:image/png;base64, " + ExtractImage(picture) + "')";
                cssDeclarations["background-repeat"] = "no-repeat";
                cssDeclarations["background-size"] = "contain";
                cssDeclarations["background-origin"] = "content-box";
                cssDeclarations["background-position-y"] = "center";
            }
        }

        private static string ExtractImage(ExcelPicture drawing)
        {
            string base64String;
            using (var stream = new MemoryStream())
            {
                drawing.Image.Save(stream, ImageFormat.Png);
                byte[] imageBytes = stream.ToArray();
                base64String = Convert.ToBase64String(imageBytes);
            }
            return base64String;
        }

        public static CssInlineStyles ToCSS(this ExcelStyles styles)
        {
            throw new NotImplementedException();
        }

        public static bool HasDrawings(this ExcelWorksheet sheet)
        {
            return sheet.Drawings.Any(d => d is ExcelPicture && ((ExcelPicture) d).Image != null);
        }
        public static IEnumerable<ExcelPicture> AllDrawings(this ExcelWorksheet sheet)
        {
            return sheet.Drawings.Where(d => d is ExcelPicture && ((ExcelPicture)d).Image != null).Select(d => d as ExcelPicture);
        }
    }

}

