using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelDropViewer
{
    internal static class ReferenceCheckTransformer
    {
        private static readonly string[] ReferenceHeaderNames =
        {
            "Reference",
            "Reference Designator",
            "Ref. Designator",
            "Ref Designator",
            "Ref",
            "Ref Des",
            "Designator",
            "부품위치"
        };

        private static readonly Regex ReferenceTokenPattern = new(
            @"^([A-Za-z]+)(\d+)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private const string PrefixGroupSeparator = "----------";

        public static ReferenceCheckReport CheckContinuity(
            DataTable table,
            Action<int, int>? onProgress = null)
        {
            if (table == null || table.Rows.Count < 2)
            {
                throw new InvalidOperationException("헤더 행을 제외하고 확인할 데이터 행이 없습니다.");
            }

            var referenceColumn = FindReferenceColumn(table.Rows[0]);
            if (referenceColumn < 0)
            {
                throw new InvalidOperationException("첫 행에서 Reference 열을 찾을 수 없습니다.");
            }

            var referencesByPrefix = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            var totalRows = Math.Max(0, table.Rows.Count - 1);

            for (var rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
            {
                onProgress?.Invoke(rowIndex, totalRows);
                CollectReferences(GetCellText(table.Rows[rowIndex][referenceColumn]), referencesByPrefix);
            }

            if (referencesByPrefix.Count == 0)
            {
                throw new InvalidOperationException("Reference 열에서 확인할 레퍼런스 번호를 찾을 수 없습니다.");
            }

            var reportLines = new List<string>();
            var isFirstPrefix = true;
            foreach (var prefixEntry in referencesByPrefix.OrderBy(static entry => entry.Key, StringComparer.OrdinalIgnoreCase))
            {
                if (!isFirstPrefix)
                {
                    reportLines.Add(PrefixGroupSeparator);
                }

                isFirstPrefix = false;

                var prefix = prefixEntry.Key;
                var numbers = prefixEntry.Value;
                var min = numbers.Min();
                var max = numbers.Max();

                var missing = new List<string>();
                for (var number = min; number < max; number++)
                {
                    if (!numbers.Contains(number))
                    {
                        missing.Add($"{prefix}{number}");
                    }
                }

                var missingText = missing.Count == 0
                    ? "없음"
                    : string.Join(", ", missing);

                reportLines.Add($"미사용 번호 : {missingText}");
                reportLines.Add($"마지막 번호 : {prefix}{max}");
            }

            return new ReferenceCheckReport(reportLines);
        }

        private static void CollectReferences(string cellText, IDictionary<string, HashSet<int>> referencesByPrefix)
        {
            if (string.IsNullOrWhiteSpace(cellText))
            {
                return;
            }

            foreach (var token in SplitReferenceTokens(cellText))
            {
                var match = ReferenceTokenPattern.Match(token);
                if (!match.Success)
                {
                    continue;
                }

                var prefix = match.Groups[1].Value;
                if (!int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var number))
                {
                    continue;
                }

                if (!referencesByPrefix.TryGetValue(prefix, out var numbers))
                {
                    numbers = new HashSet<int>();
                    referencesByPrefix[prefix] = numbers;
                }

                numbers.Add(number);
            }
        }

        private static IEnumerable<string> SplitReferenceTokens(string cellText)
        {
            return cellText
                .Split(new[] { ',', ';', '\n', '\r', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(static token => token.Trim())
                .Where(static token => token.Length > 0);
        }

        private static int FindReferenceColumn(DataRow headerRow)
        {
            for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
            {
                if (MatchesReferenceHeader(GetCellText(headerRow[columnIndex])))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static bool MatchesReferenceHeader(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var normalized = text.Trim();
            foreach (var candidate in ReferenceHeaderNames.OrderByDescending(static name => name.Length))
            {
                if (normalized.Equals(candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return System.Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
        }
    }

    internal sealed class ReferenceCheckReport
    {
        public ReferenceCheckReport(IReadOnlyList<string> reportLines)
        {
            ReportLines = reportLines;
            MessageText = string.Join(Environment.NewLine, reportLines);
        }

        public IReadOnlyList<string> ReportLines { get; }

        public string MessageText { get; }
    }
}
