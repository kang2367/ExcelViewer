using System.IO;
using System.Text;

namespace ExcelDropViewer
{
    internal static class ConfigIniFile
    {
        public static string FilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CONFIG.INI");

        public static string? ReadValue(string section, string key)
        {
            if (!File.Exists(FilePath))
            {
                return null;
            }

            var inSection = false;
            foreach (var rawLine in File.ReadAllLines(FilePath, Encoding.UTF8))
            {
                var line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith(';') || line.StartsWith('#'))
                {
                    continue;
                }

                if (line.StartsWith('[') && line.EndsWith(']'))
                {
                    inSection = string.Equals(line[1..^1].Trim(), section, StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (!inSection)
                {
                    continue;
                }

                var separatorIndex = line.IndexOf('=');
                if (separatorIndex < 0)
                {
                    continue;
                }

                var currentKey = line[..separatorIndex].Trim();
                if (!string.Equals(currentKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return line[(separatorIndex + 1)..].Trim();
            }

            return null;
        }

        public static void WriteSection(string section, IReadOnlyDictionary<string, string> values)
        {
            var lines = File.Exists(FilePath)
                ? File.ReadAllLines(FilePath, Encoding.UTF8).ToList()
                : new List<string>();

            var sectionHeader = $"[{section}]";
            var sectionStart = FindSectionStart(lines, section);
            if (sectionStart >= 0)
            {
                var sectionEnd = FindSectionEnd(lines, sectionStart);
                lines.RemoveRange(sectionStart, sectionEnd - sectionStart);
            }

            var sectionLines = new List<string> { sectionHeader };
            foreach (var pair in values)
            {
                sectionLines.Add($"{pair.Key}={pair.Value}");
            }

            if (lines.Count > 0 && !string.IsNullOrWhiteSpace(lines[^1]))
            {
                lines.Add(string.Empty);
            }

            lines.AddRange(sectionLines);
            File.WriteAllLines(FilePath, lines, Encoding.UTF8);
        }

        private static int FindSectionStart(List<string> lines, string section)
        {
            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (line.StartsWith('[') && line.EndsWith(']')
                    && string.Equals(line[1..^1].Trim(), section, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindSectionEnd(List<string> lines, int sectionStart)
        {
            for (var i = sectionStart + 1; i < lines.Count; i++)
            {
                var line = lines[i].Trim();
                if (line.StartsWith('[') && line.EndsWith(']'))
                {
                    return i;
                }
            }

            return lines.Count;
        }
    }
}
