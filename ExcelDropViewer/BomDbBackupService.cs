using System;
using System.Globalization;
using System.IO;

namespace ExcelDropViewer
{
    internal static class BomDbBackupService
    {
        public const string BackupFolderName = "Backup";

        public static string? CreateBackupIfExists(string databasePath)
        {
            if (!File.Exists(databasePath))
            {
                return null;
            }

            var backupDirectory = GetBackupDirectory();
            Directory.CreateDirectory(backupDirectory);

            var lastWriteTime = File.GetLastWriteTime(databasePath);
            var dateSuffix = lastWriteTime.ToString("yyMMdd", CultureInfo.InvariantCulture);
            var baseFileName = $"BOM_Master_{dateSuffix}.db";
            var backupFileName = ResolveUniqueBackupFileName(backupDirectory, baseFileName, dateSuffix);
            var backupPath = Path.Combine(backupDirectory, backupFileName);

            File.Copy(databasePath, backupPath, overwrite: false);
            return backupFileName;
        }

        public static string GetBackupDirectory()
        {
            return Path.Combine(BomDbPaths.GetDataDirectory(), BackupFolderName);
        }

        private static string ResolveUniqueBackupFileName(
            string backupDirectory,
            string baseFileName,
            string dateSuffix)
        {
            var dataDirectory = BomDbPaths.GetDataDirectory();
            var sequence = 0;

            while (true)
            {
                var candidateFileName = sequence == 0
                    ? baseFileName
                    : $"BOM_Master_{dateSuffix}_{sequence:000}.db";

                if (!BackupFileExists(backupDirectory, dataDirectory, candidateFileName))
                {
                    return candidateFileName;
                }

                sequence++;
            }
        }

        private static bool BackupFileExists(
            string backupDirectory,
            string dataDirectory,
            string fileName)
        {
            return File.Exists(Path.Combine(backupDirectory, fileName))
                || File.Exists(Path.Combine(dataDirectory, fileName));
        }
    }
}
