using System.Globalization;
using EPPlus.Html.Html;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;

namespace EPPlus.Html.Converters
{
    internal static class ExcelToCss
    {
        internal static CssDeclaration ToCss(this ExcelRow excelRow, ConversionFlags flags = ConversionFlags.Default)
        {
            var css = new CssDeclaration();
            if (flags.HasFlag(ConversionFlags.NoCss))
            {
                return css;
            }
            //Make sure we don't return comma as decimal separator. Not supported in HTML.
            if (!flags.HasFlag(ConversionFlags.AutoHeights))
            {
                css["height"] = (!flags.HasFlag(ConversionFlags.ExactPixelHeights) ? ((int)Math.Round(excelRow.Height, 0)).ToString() : excelRow.Height.ToString(CultureInfo.GetCultureInfo("en-US"))) + "px";
            }
            
            css.Update(excelRow.Style.ToCss());
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelRange excelRange, ConversionFlags flags = ConversionFlags.Default)
        {
            var css = new CssDeclaration();
            if (flags.HasFlag(ConversionFlags.NoCss))
            {
                return css;
            }
            if (!flags.HasFlag(ConversionFlags.AutoWidths) && excelRange.Columns == 1 && excelRange.Rows == 1)
            {
                var excelColumn = excelRange.Worksheet.Column(excelRange.Start.Column);
                //Make sure we don't return comma as decimal separator. Not supported in HTML.
                css["max-width"] = excelColumn.Width.ToString(CultureInfo.GetCultureInfo("en-US")) + "em";
                css["width"] = excelColumn.Width.ToString(CultureInfo.GetCultureInfo("en-US")) + "em";
            }
            css.Update(excelRange.Style.ToCss());
            return css;
        }      

        internal static CssDeclaration ToCss(this ExcelStyle excelStyle)
        {
            var css = new CssDeclaration();
            css["text-align"] = excelStyle.HorizontalAlignment.ToCssProperty();
            css["background-color"] = excelStyle.Fill.BackgroundColor.ToHexCode();
            css.Update(excelStyle.Font.ToCss());
            css.Update(excelStyle.Border.ToCss());
            return css;
        }

        internal static CssDeclaration ToCss(this ExcelFont excelFont)
        {
            var css = new CssDeclaration();

            if (excelFont.Bold)
            {
                css["font-weight"] = "bold";
            }
            if (excelFont.Italic)
            {
                css["font-style"] = "italic";
            }
            else if (excelFont.UnderLine)
            {
                css["font-style"] = "underline";
            }
            else if (excelFont.UnderLine)
            {
                css["text-decoration"] = "underline";
            }
            
            css["font-family"] = excelFont.Name;
            css["font-size"] = excelFont.Size + "pt";

            css["color"] = excelFont.Color.ToHexCode();

            return css;
        }

        internal static CssDeclaration ToCss(this Border border)
        {
            var css = new CssDeclaration();

            css["border-top"] = border.Top.ToCssProperty();
            css["border-bottom"] = border.Bottom.ToCssProperty();
            css["border-right"] = border.Right.ToCssProperty();
            css["border-left"] = border.Left.ToCssProperty();

            return css;
        }

        internal static string ToCssProperty(this ExcelBorderItem excelBorderItem)
        {
            if (excelBorderItem != null)
            {
                string color = (excelBorderItem.Color.Rgb != null)
                    ? excelBorderItem.Color.ToHexCode()
                    : "black";

                return excelBorderItem.Style.ToCssProperty() + " " + color;
            }
            else
            {
                return null;
            }
        }

        internal static string ToCssProperty(this ExcelBorderStyle excelBorderStyle)
        {
            switch (excelBorderStyle)
            {
                case ExcelBorderStyle.Thin:
                    return "1px solid";
                case ExcelBorderStyle.Thick:
                    return "2px solid";
                default:
                    return "none";
            }
        }

        internal static string ToCssProperty(this ExcelHorizontalAlignment excelHorizontalAlignment)
        {
            switch (excelHorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Left:
                    return "left";
                case ExcelHorizontalAlignment.Center:
                    return "center";
                default:
                    return null;
            }
        }

        internal static string ToHexCode(this ExcelColor excelColor)
        {
            if (excelColor != null && excelColor.Rgb != null && excelColor.Rgb.Length > 3)
            {
                return "#" + excelColor.Rgb.Substring(2);
            }
            else
            {
                return null;
            }
        }
    }
}
