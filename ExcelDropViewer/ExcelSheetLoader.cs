using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using ExcelDataReader;

namespace ExcelDropViewer
{
    internal static class ExcelSheetLoader
    {
        private const string MaxTextLengthKey = "MaxTextLength";

        public static DataTable LoadFirstSheet(string filePath)
        {
            using var stream = File.Open(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            using var reader = ExcelReaderFactory.CreateReader(stream);

            var rowBuffer = new List<string[]>();
            var columnStats = new List<ColumnStats>();
            var columnCount = 0;

            while (reader.Read())
            {
                var fieldCount = reader.FieldCount;
                if (fieldCount > columnCount)
                {
                    columnCount = fieldCount;
                }

                while (columnStats.Count < fieldCount)
                {
                    columnStats.Add(new ColumnStats());
                }

                var row = new string[fieldCount];
                for (var columnIndex = 0; columnIndex < fieldCount; columnIndex++)
                {
                    var text = FormatCellValue(reader.GetValue(columnIndex));
                    row[columnIndex] = text;

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    var stats = columnStats[columnIndex];
                    stats.HasData = true;
                    if (text.Length > stats.MaxTextLength)
                    {
                        stats.MaxTextLength = text.Length;
                    }
                }

                rowBuffer.Add(row);
            }

            var includedSourceIndexes = new List<int>(columnCount);
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                if (columnIndex < columnStats.Count && columnStats[columnIndex].HasData)
                {
                    includedSourceIndexes.Add(columnIndex);
                }
            }

            var table = new DataTable();
            for (var j = 0; j < includedSourceIndexes.Count; j++)
            {
                var sourceIndex = includedSourceIndexes[j];
                var column = table.Columns.Add($"F{j}", typeof(string));
                column.ExtendedProperties[MaxTextLengthKey] = columnStats[sourceIndex].MaxTextLength;
            }

            table.BeginLoadData();
            foreach (var sourceRow in rowBuffer)
            {
                var newRow = table.NewRow();
                for (var j = 0; j < includedSourceIndexes.Count; j++)
                {
                    var sourceIndex = includedSourceIndexes[j];
                    newRow[j] = sourceIndex < sourceRow.Length ? sourceRow[sourceIndex] : string.Empty;
                }

                table.Rows.Add(newRow);
            }

            table.EndLoadData();
            return table;
        }

        private sealed class ColumnStats
        {
            public bool HasData;
            public int MaxTextLength;
        }

        private static string FormatCellValue(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            switch (value)
            {
                case string s:
                    return s;
                case DateTime dt:
                    return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
                case bool b:
                    return b ? "TRUE" : "FALSE";
                case double d:
                    return FormatNumber(d);
                case float f:
                    return FormatNumber(f);
                case decimal m:
                    return m.ToString(CultureInfo.CurrentCulture);
                default:
                    return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
            }
        }

        private static string FormatNumber(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return string.Empty;
            }

            if (Math.Abs(value) < double.Epsilon)
            {
                return "0";
            }

            if (value > long.MinValue && value < long.MaxValue
                && Math.Abs(value - Math.Round(value)) < 1e-9)
            {
                return ((long)Math.Round(value)).ToString(CultureInfo.CurrentCulture);
            }

            return value.ToString("G", CultureInfo.CurrentCulture);
        }
    }
}
