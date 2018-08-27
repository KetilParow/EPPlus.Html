using System;

namespace EPPlus.Html.Converters
{
    [Flags]
    public enum ConversionFlags
    {
        Default = 0,
        AutoWidths = 1,
        AutoHeights = 2,
        RightAlignNumerals = 4,
        IgnoreColSpans = 8,
        ExactPixelHeights = 16,
        NoCss = 32,
        AggregateStyleSheet = 64
    }

}

