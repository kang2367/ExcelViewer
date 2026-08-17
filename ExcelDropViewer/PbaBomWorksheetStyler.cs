using System.Drawing;
using unvell.ReoGrid;

namespace ExcelDropViewer
{
    internal static class PbaBomWorksheetStyler
    {
        private const float DefaultBodyFontSize = 10f;

        public static float TryGetBaseFontSize(Worksheet sheet, int sampleRow, int sampleColumn)
        {
            if (sampleRow < 0 || sampleColumn < 0)
            {
                return DefaultBodyFontSize;
            }

            var style = sheet.GetCellStyles(sampleRow, sampleColumn);
            if (style != null && style.FontSize > 0)
            {
                return style.FontSize;
            }

            return DefaultBodyFontSize;
        }

        public static void Apply(Worksheet sheet, PbaBomOutputLayout layout, float baseFontSize)
        {
            if (layout.DataEndRow >= layout.DataStartRow)
            {
                if (layout.ReferenceColumn >= 0)
                {
                    ApplyReferenceColumnStyles(sheet, layout);
                }

                if (layout.ItemDescriptionColumn >= 0)
                {
                    ApplyItemDescriptionTopAlign(sheet, layout);
                }
            }

            if (layout.NcListTitleRow >= 0)
            {
                ApplyNcListTitleStyle(sheet, layout, baseFontSize);
            }
        }

        private static void ApplyReferenceColumnStyles(Worksheet sheet, PbaBomOutputLayout layout)
        {
            var rowCount = layout.DataEndRow - layout.DataStartRow + 1;
            sheet.SetRangeStyles(
                new RangePosition(layout.DataStartRow, layout.ReferenceColumn, rowCount, 1),
                new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.TextWrap | PlainStyleFlag.VerticalAlign,
                    TextWrapMode = TextWrapMode.WordBreak,
                    VAlign = ReoGridVerAlign.Top
                });
        }

        private static void ApplyItemDescriptionTopAlign(Worksheet sheet, PbaBomOutputLayout layout)
        {
            var rowCount = layout.DataEndRow - layout.DataStartRow + 1;
            sheet.SetRangeStyles(
                new RangePosition(layout.DataStartRow, layout.ItemDescriptionColumn, rowCount, 1),
                new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.VerticalAlign,
                    VAlign = ReoGridVerAlign.Top
                });
        }

        private static void ApplyNcListTitleStyle(
            Worksheet sheet,
            PbaBomOutputLayout layout,
            float baseFontSize)
        {
            sheet.SetRangeStyles(
                new RangePosition(layout.NcListTitleRow, layout.NcListTitleColumn, 1, 1),
                new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontStyleBold | PlainStyleFlag.FontSize,
                    Bold = true,
                    FontSize = baseFontSize + 2f,
                    TextColor = Color.Black
                });
        }
    }
}
