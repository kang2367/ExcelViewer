using System.Text.RegularExpressions;

namespace ExcelDropViewer
{
    internal static class DigiKeyManufacturerNameNormalizer
    {
        private static readonly Regex ParenthesesOrBracketPattern = new(
            @"\s*[\(\[][^)\]]*[\)\]]",
            RegexOptions.Compiled);

        private static readonly Regex MultipleSpacesPattern = new(
            @"\s{2,}",
            RegexOptions.Compiled);

        public static string Normalize(string? manufacturerName)
        {
            if (string.IsNullOrWhiteSpace(manufacturerName))
            {
                return string.Empty;
            }

            var withoutBrackets = ParenthesesOrBracketPattern.Replace(manufacturerName, string.Empty);
            return MultipleSpacesPattern.Replace(withoutBrackets, " ").Trim();
        }
    }
}
