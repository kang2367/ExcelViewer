using System.Text.RegularExpressions;

namespace ExcelDropViewer
{
    internal enum BomDescriptionMatchKind
    {
        Exact,
        Attribute,
        Fuzzy
    }

    internal sealed class BomDescriptionMatchResult
    {
        public BomPartRecord Record { get; init; } = new();
        public double SimilarityPercent { get; init; }
        public BomDescriptionMatchKind MatchKind { get; init; }
    }

    internal static class BomDescriptionMatchEngine
    {
        private const double FuzzyMatchThresholdPercent = 70.0;

        private static readonly Regex CapacitanceTokenPattern = new(
            @"\d+(?:\.\d+)?\s*(?:uF|nF|pF|µF)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static BomDescriptionMatchResult? Match(
            string excelItemDescription,
            IReadOnlyList<BomPartRecord> allParts)
        {
            if (string.IsNullOrWhiteSpace(excelItemDescription) || allParts.Count == 0)
            {
                return null;
            }

            var normalizedExcel = excelItemDescription.Trim();

            var exactMatch = allParts.FirstOrDefault(part =>
                string.Equals(part.ItemDescription, normalizedExcel, StringComparison.Ordinal));
            if (exactMatch != null)
            {
                return CreateResult(exactMatch, 100, BomDescriptionMatchKind.Exact);
            }

            exactMatch = allParts.FirstOrDefault(part =>
                string.Equals(part.ItemDescription.Trim(), normalizedExcel, StringComparison.OrdinalIgnoreCase));
            if (exactMatch != null)
            {
                return CreateResult(exactMatch, 100, BomDescriptionMatchKind.Exact);
            }

            if (TryMatchResistorByAttributes(normalizedExcel, allParts, out var resistorMatch))
            {
                return CreateResult(resistorMatch, 100, BomDescriptionMatchKind.Attribute);
            }

            if (TryMatchCapacitorByAttributes(normalizedExcel, allParts, out var capacitorMatch))
            {
                return CreateResult(capacitorMatch, 100, BomDescriptionMatchKind.Attribute);
            }

            return TryMatchByFuzzySimilarity(normalizedExcel, allParts);
        }

        public static string FormatDbDescriptionForOutput(BomDescriptionMatchResult matchResult)
        {
            var description = matchResult.Record.ItemDescription;
            if (matchResult.MatchKind == BomDescriptionMatchKind.Fuzzy)
            {
                return $"{description} (매칭률: {matchResult.SimilarityPercent:0}%)";
            }

            return description;
        }

        private static bool TryMatchResistorByAttributes(
            string excelDescription,
            IReadOnlyList<BomPartRecord> allParts,
            out BomPartRecord matchedRecord)
        {
            matchedRecord = null!;
            var excelAttributes = ParseResistorAttributes(excelDescription);
            if (excelAttributes == null)
            {
                return false;
            }

            foreach (var part in allParts)
            {
                if (!part.ItemDescription.TrimStart().StartsWith("RES", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (MatchesResistorAttributes(excelAttributes, part.ItemDescription))
                {
                    matchedRecord = part;
                    return true;
                }
            }

            return false;
        }

        private static bool TryMatchCapacitorByAttributes(
            string excelDescription,
            IReadOnlyList<BomPartRecord> allParts,
            out BomPartRecord matchedRecord)
        {
            matchedRecord = null!;
            var excelAttributes = ParseCapacitorAttributes(excelDescription);
            if (excelAttributes == null)
            {
                return false;
            }

            foreach (var part in allParts)
            {
                if (!part.ItemDescription.TrimStart().StartsWith("CAP", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (MatchesCapacitorAttributes(excelAttributes, part.ItemDescription))
                {
                    matchedRecord = part;
                    return true;
                }
            }

            return false;
        }

        private static BomDescriptionMatchResult? TryMatchByFuzzySimilarity(
            string excelDescription,
            IReadOnlyList<BomPartRecord> allParts)
        {
            BomPartRecord? bestRecord = null;
            var bestSimilarity = 0.0;

            foreach (var part in allParts)
            {
                var similarity = CalculateSimilarityPercent(excelDescription, part.ItemDescription);
                if (similarity > bestSimilarity)
                {
                    bestSimilarity = similarity;
                    bestRecord = part;
                }
            }

            if (bestRecord == null || bestSimilarity < FuzzyMatchThresholdPercent)
            {
                return null;
            }

            return CreateResult(bestRecord, bestSimilarity, BomDescriptionMatchKind.Fuzzy);
        }

        private static ResistorAttributes? ParseResistorAttributes(string description)
        {
            if (!description.TrimStart().StartsWith("RES", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var tokens = SplitDescriptionTokens(description);
            if (tokens.Count < 4)
            {
                return null;
            }

            return new ResistorAttributes
            {
                Size = tokens[1],
                Resistance = tokens[2],
                Tolerance = tokens[3]
            };
        }

        private static CapacitorAttributes? ParseCapacitorAttributes(string description)
        {
            if (!description.TrimStart().StartsWith("CAP", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var tokens = SplitDescriptionTokens(description);
            if (tokens.Count < 4)
            {
                return null;
            }

            var sizeIndex = tokens.Count >= 6 ? 2 : 1;
            var valueIndex = sizeIndex + 1;
            var voltageIndex = valueIndex + 1;
            var toleranceIndex = voltageIndex + 1;

            if (tokens.Count <= toleranceIndex)
            {
                return null;
            }

            return new CapacitorAttributes
            {
                Size = tokens[sizeIndex],
                Capacitance = tokens[valueIndex],
                Voltage = tokens[voltageIndex],
                Tolerance = tokens[toleranceIndex]
            };
        }

        private static bool MatchesResistorAttributes(ResistorAttributes attributes, string dbDescription)
        {
            return ContainsToken(dbDescription, attributes.Size)
                && ContainsToken(dbDescription, attributes.Resistance)
                && ContainsToken(dbDescription, attributes.Tolerance);
        }

        private static bool MatchesCapacitorAttributes(CapacitorAttributes attributes, string dbDescription)
        {
            return ContainsToken(dbDescription, attributes.Size)
                && ContainsCapacitanceToken(dbDescription, attributes.Capacitance)
                && ContainsToken(dbDescription, attributes.Voltage)
                && ContainsToken(dbDescription, attributes.Tolerance);
        }

        private static List<string> SplitDescriptionTokens(string description)
        {
            return description
                .Split(',')
                .Select(token => token.Trim())
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToList();
        }

        private static bool ContainsToken(string source, string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            return NormalizeForComparison(source).Contains(
                NormalizeForComparison(token),
                StringComparison.Ordinal);
        }

        private static bool ContainsCapacitanceToken(string source, string capacitance)
        {
            if (string.IsNullOrWhiteSpace(capacitance))
            {
                return false;
            }

            if (ContainsToken(source, capacitance))
            {
                return true;
            }

            var normalizedCapacitance = NormalizeCapacitance(capacitance);
            return CapacitanceTokenPattern.Matches(source)
                .Select(match => NormalizeCapacitance(match.Value))
                .Any(match => match.Equals(normalizedCapacitance, StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeCapacitance(string value)
        {
            return value
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("µF", "uF", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeForComparison(string value)
        {
            return value
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("%", "%", StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        private static double CalculateSimilarityPercent(string left, string right)
        {
            if (string.IsNullOrEmpty(left) && string.IsNullOrEmpty(right))
            {
                return 100.0;
            }

            if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
            {
                return 0.0;
            }

            var distance = ComputeLevenshteinDistance(left, right);
            var maxLength = Math.Max(left.Length, right.Length);
            return (1.0 - (double)distance / maxLength) * 100.0;
        }

        private static int ComputeLevenshteinDistance(string left, string right)
        {
            var leftLength = left.Length;
            var rightLength = right.Length;
            var distances = new int[leftLength + 1, rightLength + 1];

            for (var leftIndex = 0; leftIndex <= leftLength; leftIndex++)
            {
                distances[leftIndex, 0] = leftIndex;
            }

            for (var rightIndex = 0; rightIndex <= rightLength; rightIndex++)
            {
                distances[0, rightIndex] = rightIndex;
            }

            for (var leftIndex = 1; leftIndex <= leftLength; leftIndex++)
            {
                for (var rightIndex = 1; rightIndex <= rightLength; rightIndex++)
                {
                    var cost = left[leftIndex - 1] == right[rightIndex - 1] ? 0 : 1;
                    distances[leftIndex, rightIndex] = Math.Min(
                        Math.Min(distances[leftIndex - 1, rightIndex] + 1, distances[leftIndex, rightIndex - 1] + 1),
                        distances[leftIndex - 1, rightIndex - 1] + cost);
                }
            }

            return distances[leftLength, rightLength];
        }

        private static BomDescriptionMatchResult CreateResult(
            BomPartRecord record,
            double similarityPercent,
            BomDescriptionMatchKind matchKind)
        {
            return new BomDescriptionMatchResult
            {
                Record = record,
                SimilarityPercent = similarityPercent,
                MatchKind = matchKind
            };
        }

        private sealed class ResistorAttributes
        {
            public string Size { get; init; } = string.Empty;
            public string Resistance { get; init; } = string.Empty;
            public string Tolerance { get; init; } = string.Empty;
        }

        private sealed class CapacitorAttributes
        {
            public string Size { get; init; } = string.Empty;
            public string Capacitance { get; init; } = string.Empty;
            public string Voltage { get; init; } = string.Empty;
            public string Tolerance { get; init; } = string.Empty;
        }
    }
}
