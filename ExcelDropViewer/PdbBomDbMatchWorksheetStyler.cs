using System.Drawing;
using unvell.ReoGrid;

namespace ExcelDropViewer
{
    internal static class PdbBomDbMatchWorksheetStyler
    {
        private static readonly Color MismatchHighlightColor = Color.FromArgb(255, 220, 200);

        public static void Apply(Worksheet sheet, PdbBomDbMatchLayout layout)
        {
            if (layout.DataEndRow < layout.DataStartRow)
            {
                return;
            }

            ApplyTopAlignment(sheet, layout);
            ApplyMismatchHighlights(sheet, layout);
        }

        private static void ApplyTopAlignment(Worksheet sheet, PdbBomDbMatchLayout layout)
        {
            var rowCount = layout.DataEndRow - layout.DataStartRow + 1;
            foreach (var columnIndex in GetMatchColumns(layout))
            {
                if (columnIndex < 0)
                {
                    continue;
                }

                sheet.SetRangeStyles(
                    new RangePosition(layout.DataStartRow, columnIndex, rowCount, 1),
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.VerticalAlign,
                        VAlign = ReoGridVerAlign.Top
                    });
            }
        }

        private static void ApplyMismatchHighlights(Worksheet sheet, PdbBomDbMatchLayout layout)
        {
            if (layout.DbDescriptionColumn < 0)
            {
                return;
            }

            foreach (var rowIndex in layout.UnmatchedRows)
            {
                sheet.SetRangeStyles(
                    new RangePosition(rowIndex, layout.DbDescriptionColumn, 1, 1),
                    new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor,
                        BackColor = MismatchHighlightColor
                    });

                if (layout.PartNumberColumn >= 0)
                {
                    sheet.SetRangeStyles(
                        new RangePosition(rowIndex, layout.PartNumberColumn, 1, 1),
                        new WorksheetRangeStyle
                        {
                            Flag = PlainStyleFlag.BackColor,
                            BackColor = MismatchHighlightColor
                        });
                }

                if (layout.ManufacturerColumn >= 0)
                {
                    sheet.SetRangeStyles(
                        new RangePosition(rowIndex, layout.ManufacturerColumn, 1, 1),
                        new WorksheetRangeStyle
                        {
                            Flag = PlainStyleFlag.BackColor,
                            BackColor = MismatchHighlightColor
                        });
                }
            }
        }

        private static IEnumerable<int> GetMatchColumns(PdbBomDbMatchLayout layout)
        {
            yield return layout.PartNameColumn;
            yield return layout.PartNumberColumn;
            yield return layout.ManufacturerColumn;
            yield return layout.DbDescriptionColumn;
        }
    }
}
