using System;
using System.IO;

namespace ExcelDropViewer
{
    internal static class BomDbPaths
    {
        public const string DataFolderName = "Data";
        public const string DatabaseFileName = "BOM_Master.db";

        public static string GetDatabasePath()
        {
            var dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DataFolderName);
            Directory.CreateDirectory(dataDirectory);
            return Path.Combine(dataDirectory, DatabaseFileName);
        }

        public static string GetDataDirectory()
        {
            var dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DataFolderName);
            Directory.CreateDirectory(dataDirectory);
            return dataDirectory;
        }
    }
}
