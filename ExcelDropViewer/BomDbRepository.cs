using System;
using System.Globalization;
using System.IO;
using Microsoft.Data.Sqlite;

namespace ExcelDropViewer
{
    internal sealed class BomDbRepository : IDisposable
    {
        private const string TableName = "PartMaster";

        private static readonly string CreateTableSql = $"""
            CREATE TABLE IF NOT EXISTS {TableName} (
                item_number TEXT PRIMARY KEY NOT NULL,
                part_name TEXT NOT NULL DEFAULT '',
                item_description TEXT NOT NULL DEFAULT '',
                manufacturer TEXT NOT NULL DEFAULT '',
                operating_temperature TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL
            );
            """;

        private static readonly string ExistsSql = $"""
            SELECT 1
            FROM {TableName}
            WHERE item_number = $item_number
            LIMIT 1;
            """;

        private static readonly string InsertSql = $"""
            INSERT INTO {TableName} (
                item_number,
                part_name,
                item_description,
                manufacturer,
                operating_temperature,
                updated_at)
            VALUES (
                $item_number,
                $part_name,
                $item_description,
                $manufacturer,
                $operating_temperature,
                $updated_at);
            """;

        private static readonly string UpdateSql = $"""
            UPDATE {TableName}
            SET
                part_name = $part_name,
                item_description = $item_description,
                manufacturer = $manufacturer,
                operating_temperature = $operating_temperature,
                updated_at = $updated_at
            WHERE item_number = $item_number;
            """;

        private static readonly string SelectByDescriptionExactSql = $"""
            SELECT
                item_number,
                part_name,
                item_description,
                manufacturer,
                operating_temperature
            FROM {TableName}
            WHERE item_description = $item_description
            LIMIT 1;
            """;

        private static readonly string SelectByDescriptionIgnoreCaseSql = $"""
            SELECT
                item_number,
                part_name,
                item_description,
                manufacturer,
                operating_temperature
            FROM {TableName}
            WHERE UPPER(TRIM(item_description)) = UPPER(TRIM($item_description))
            LIMIT 1;
            """;

        private static readonly string SelectAllSql = $"""
            SELECT
                item_number,
                part_name,
                item_description,
                manufacturer,
                operating_temperature
            FROM {TableName};
            """;

        private readonly SqliteConnection _connection;
        private readonly string _databasePath;
        private SqliteTransaction? _transaction;

        public BomDbRepository(string databasePath)
        {
            var directory = Path.GetDirectoryName(databasePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            _databasePath = Path.GetFullPath(databasePath);
            _connection = new SqliteConnection($"Data Source={_databasePath}");
            _connection.Open();
            EnsureSchema();
        }

        public string DatabasePath => _databasePath;

        public void BeginTransaction()
        {
            if (_transaction != null)
            {
                throw new InvalidOperationException("이미 진행 중인 트랜잭션이 있습니다.");
            }

            _transaction = _connection.BeginTransaction();
        }

        public void CommitTransaction()
        {
            if (_transaction == null)
            {
                return;
            }

            _transaction.Commit();
            _transaction.Dispose();
            _transaction = null;
        }

        public void RollbackTransaction()
        {
            if (_transaction == null)
            {
                return;
            }

            _transaction.Rollback();
            _transaction.Dispose();
            _transaction = null;
        }

        public bool ExistsByItemNumber(string itemNumber)
        {
            using var command = CreateCommand(ExistsSql);
            command.Parameters.AddWithValue("$item_number", itemNumber);
            using var reader = command.ExecuteReader();
            return reader.Read();
        }

        public int Insert(BomPartRecord record)
        {
            using var command = CreateCommand(InsertSql);
            BindRecordParameters(command, record);
            return command.ExecuteNonQuery();
        }

        public int Update(BomPartRecord record)
        {
            using var command = CreateCommand(UpdateSql);
            BindRecordParameters(command, record);
            return command.ExecuteNonQuery();
        }

        public BomPartRecord? TryFindByItemDescription(string itemDescription)
        {
            if (string.IsNullOrWhiteSpace(itemDescription))
            {
                return null;
            }

            var exactMatch = TryFindByItemDescriptionInternal(itemDescription, ignoreCase: false);
            if (exactMatch != null)
            {
                return exactMatch;
            }

            return TryFindByItemDescriptionInternal(itemDescription, ignoreCase: true);
        }

        public IReadOnlyList<BomPartRecord> GetAllParts()
        {
            var records = new List<BomPartRecord>();
            using var command = CreateCommand(SelectAllSql);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                records.Add(ReadPartRecord(reader));
            }

            return records;
        }

        private BomPartRecord? TryFindByItemDescriptionInternal(string itemDescription, bool ignoreCase)
        {
            using var command = CreateCommand(ignoreCase
                ? SelectByDescriptionIgnoreCaseSql
                : SelectByDescriptionExactSql);
            command.Parameters.AddWithValue("$item_description", itemDescription.Trim());
            using var reader = command.ExecuteReader();
            return reader.Read() ? ReadPartRecord(reader) : null;
        }

        private static BomPartRecord ReadPartRecord(SqliteDataReader reader)
        {
            return new BomPartRecord
            {
                ItemNumber = reader.GetString(0),
                PartName = reader.GetString(1),
                ItemDescription = reader.GetString(2),
                Manufacturer = reader.GetString(3),
                OperatingTemperature = reader.GetString(4)
            };
        }

        public void Dispose()
        {
            _transaction?.Dispose();
            _connection.Dispose();
        }

        private void EnsureSchema()
        {
            using var command = _connection.CreateCommand();
            command.CommandText = CreateTableSql;
            command.ExecuteNonQuery();
        }

        private SqliteCommand CreateCommand(string commandText)
        {
            var command = _connection.CreateCommand();
            command.CommandText = commandText;
            if (_transaction != null)
            {
                command.Transaction = _transaction;
            }

            return command;
        }

        private static void BindRecordParameters(SqliteCommand command, BomPartRecord record)
        {
            command.Parameters.AddWithValue("$item_number", record.ItemNumber);
            command.Parameters.AddWithValue("$part_name", record.PartName);
            command.Parameters.AddWithValue("$item_description", record.ItemDescription);
            command.Parameters.AddWithValue("$manufacturer", record.Manufacturer);
            command.Parameters.AddWithValue("$operating_temperature", record.OperatingTemperature);
            command.Parameters.AddWithValue(
                "$updated_at",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
        }

        public static string GetDefaultDatabasePath()
        {
            return BomDbPaths.GetDatabasePath();
        }
    }
}
