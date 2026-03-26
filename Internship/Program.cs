using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ExcelDataReader;
using Npgsql;

namespace Internship
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Starting Excel to PostgreSQL import process...");

                var config = new AppConfig
                {
                    FolderConfigs = new List<FolderConfig>
                    {
                       new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\bole_rosii",
                            TableName = "public.bole",
                            MandatoryColumns = new[] { "Created", "Notifctn", "Material" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\CAF",
                            TableName = "public.CAF",
                            MandatoryColumns = new[] { "Order", "QTY_SCRAP", "Cause" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\CB",
                            TableName = "public.cb",
                            MandatoryColumns = new[] { "Order", "QTY_SCRAP", "Cause" }
                        },
            new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\Cottura",
                            TableName = "public.Cottura1",
                            MandatoryColumns = new[] { "Comandă", "Material", "Cauză" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\DB",
                            TableName = "public.DB1",
                            MandatoryColumns = new[] { "Comandă", "Material", "Cauză" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\LAV",
                            TableName = "public.LAV",
                            MandatoryColumns = new[] { "Order", "Material", "Cause" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\NPL",
                            TableName = "public.NPL",
                            MandatoryColumns = new[] { "Order", "Material", "Cause" }
                        },
            new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\Presso",
                            TableName = "public.presso",
                            MandatoryColumns = new[] { "Comandă", "Material", "Cauză" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\RD+CB",
                            TableName = "public.rdcb",
                            MandatoryColumns = new[] { "Order", "Material", "Cause" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\TB&OV&PS",
                            TableName = "public.\"TB&OV&PS\"",
                            MandatoryColumns = new[] { "Comandă", "Material", "Cauză" }
                        },
                         new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\RD",
                            TableName = "public.rd1",
                            MandatoryColumns = new[] { "Comandă", "Material", "Cauză" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\RS",
                            TableName = "public.RS",
                            MandatoryColumns = new[] { "Ordine", "Materiale", "Causa" }
                        },
            new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Assiemaggio GTH",
                            TableName = "public.AssiemaggioGTH",
                            MandatoryColumns = new[] { "Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Automotive_RICA",
                            TableName = "public.Automotive",
                            MandatoryColumns = new[] { "Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Boiler Lamiera",
                            TableName = "public.BoilerLAM",
                            MandatoryColumns = new[] { "Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\CAM GTH",
                            TableName = "public.CAMGTH",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
            new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\cartucce",
                            TableName = "public.Cartucce",
                            MandatoryColumns = new[] { "Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Eltra plat",
                            TableName = "public.EltPlat",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Etched foil",
                            TableName = "public.EtcFoil",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Finitura GTH",
                            TableName = "public.FinGTH",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
            new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\Piegatura GTH",
                            TableName = "public.PieGTH",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\RAPOARTE REBUT\rica\RID",
                            TableName = "public.RID",
                            MandatoryColumns = new[] {"Ordine", "Materiale", "Causa" }
                        },
                        new FolderConfig
                        {
                            FolderPath = @"T:\Ianovici Catalin\Productie",
                            TableName = "public.productie",
                            MandatoryColumns = new[] {"Order", "Order quantity", "Confirmed scrap" }
                        }
                    },
                    DatabaseConfig = new DatabaseConfig
                    {
                        Host = "aws-0-eu-central-1.pooler.supabase.com",
                        Port = 5432,
                        Username = "postgres.mkzcckpvsvjrarfhzmnp",
                        Password = "ZOPPASINDUSTRIESROMANIA",
                        Database = "postgres",
                        ProcessedFilesTable = "public.processed_files"
                    }
                };

                var importer = new ExcelToPostgresImporter();
                var allResults = importer.ExecuteImport(config);

                Console.WriteLine("\n=== Final Import Report ===");
                Console.WriteLine($"Total folders processed: {config.FolderConfigs.Count}");

                foreach (var folderResult in allResults.GroupBy(r => r.FolderPath))
                {
                    Console.WriteLine($"\nFolder: {folderResult.Key}");
                    Console.WriteLine($"Successful files: {folderResult.Count(r => r.Success)}");
                    Console.WriteLine($"Failed files: {folderResult.Count(r => !r.Success)}");

                    foreach (var error in folderResult.Where(r => !r.Success).GroupBy(r => r.ErrorMessage))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"  Error: {error.Key}");
                        Console.WriteLine($"  Files affected: {error.Count()}");
                        Console.ResetColor();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($" Critical error: {ex.Message}");
                Console.ResetColor();
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
            }
            finally
            {
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }
    }

    public class ExcelToPostgresImporter
    {
        public List<ImportResult> ExecuteImport(AppConfig config)
        {
            var results = new List<ImportResult>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using var conn = new NpgsqlConnection(BuildConnectionString(config.DatabaseConfig));
            conn.Open();

            EnsureProcessedFilesTable(conn, config.DatabaseConfig.ProcessedFilesTable);

            foreach (var folderConfig in config.FolderConfigs)
            {
                results.AddRange(ProcessFolder(conn, folderConfig, config.DatabaseConfig.ProcessedFilesTable));
            }

            return results;
        }

        private void EnsureProcessedFilesTable(NpgsqlConnection conn, string tableName)
        {
            var sql = $@"
                CREATE TABLE IF NOT EXISTS {tableName} (
                    id SERIAL PRIMARY KEY,
                    file_name VARCHAR(255) NOT NULL,
                    file_hash VARCHAR(32) NOT NULL,
                    file_size BIGINT NOT NULL,
                    target_table VARCHAR(255) NOT NULL,
                    processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    CONSTRAINT unique_file_per_table UNIQUE (file_name, target_table)
                )";

            ExecuteNonQuery(conn, sql);
        }

        private List<ImportResult> ProcessFolder(NpgsqlConnection conn, FolderConfig folderConfig, string processedFilesTable)
        {
            var results = new List<ImportResult>();

            try
            {
                Console.WriteLine($"\n=== Processing folder: {folderConfig.FolderPath} ===");

                if (!Directory.Exists(folderConfig.FolderPath))
                {
                    results.Add(new ImportResult
                    {
                        FolderPath = folderConfig.FolderPath,
                        ErrorMessage = $"Folder not found: {folderConfig.FolderPath}"
                    });
                    return results;
                }

                var excelFiles = Directory.GetFiles(folderConfig.FolderPath, "*.xlsx")
                                       .Concat(Directory.GetFiles(folderConfig.FolderPath, "*.xls"))
                                       .ToArray();

                Console.WriteLine($"Found {excelFiles.Length} Excel files");

                foreach (var filePath in excelFiles)
                {
                    var result = ProcessExcelFile(conn, folderConfig, processedFilesTable, filePath);
                    results.Add(result);
                }
            }
            catch (Exception ex)
            {
                results.Add(new ImportResult
                {
                    FolderPath = folderConfig.FolderPath,
                    ErrorMessage = $"Folder processing error: {ex.Message}"
                });
            }

            return results;
        }

        private ImportResult ProcessExcelFile(NpgsqlConnection conn, FolderConfig folderConfig, string processedFilesTable, string filePath)
        {
            var fileName = Path.GetFileName(filePath);
            var result = new ImportResult
            {
                FileName = fileName,
                FolderPath = folderConfig.FolderPath,
                TableName = folderConfig.TableName
            };

            try
            {
                byte[] fileBytes;
                string fileHash;
                long fileSize;

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var memoryStream = new MemoryStream())
                {
                    fileStream.CopyTo(memoryStream);
                    fileBytes = memoryStream.ToArray();
                    fileSize = fileStream.Length;

                    memoryStream.Position = 0;
                    using var md5 = MD5.Create();
                    fileHash = BitConverter.ToString(md5.ComputeHash(memoryStream)).Replace("-", "").ToLowerInvariant();
                }

                if (IsFileProcessed(conn, processedFilesTable, fileName, fileHash, folderConfig.TableName))
                {
                    Console.WriteLine($"Skipping already processed file: {fileName}");
                    result.ErrorMessage = "File already processed";
                    return result;
                }

                Console.WriteLine($"\nProcessing: {fileName}");

                DataTable dataTable;
                using (var memoryStream = new MemoryStream(fileBytes))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(memoryStream))
                    {
                        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true,
                                FilterRow = rowReader => {
                                    for (var i = 0; i < rowReader.FieldCount; i++)
                                    {
                                        if (!string.IsNullOrEmpty(rowReader[i]?.ToString()))
                                            return true;
                                    }
                                    return false;
                                }
                            }
                        });

                        if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                        {
                            result.ErrorMessage = "Empty file - no data found";
                            return result;
                        }

                        dataTable = dataSet.Tables[0];
                    }
                }

                var excelColumns = dataTable.Columns.Cast<DataColumn>()
                    .Select(c => c.ColumnName.Trim())
                    .Distinct()
                    .ToList();

                var missingColumns = folderConfig.MandatoryColumns
                    .Where(c => !excelColumns.Contains(c))
                    .ToList();

                if (missingColumns.Any())
                {
                    result.ErrorMessage = $"Missing mandatory columns: {string.Join(", ", missingColumns)}";
                    return result;
                }

                result.AdditionalColumnsCount = excelColumns.Count - folderConfig.MandatoryColumns.Length;
                Console.WriteLine($"Detected {excelColumns.Count} columns ({result.AdditionalColumnsCount} additional)");

                EnsureTableStructure(conn, folderConfig.TableName, excelColumns);

                result.RowsImported = ImportData(conn, folderConfig.TableName, dataTable, excelColumns, filePath, fileHash);
                MarkFileAsProcessed(conn, processedFilesTable, fileName, fileHash, fileSize, folderConfig.TableName);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        private void EnsureTableStructure(NpgsqlConnection conn, string tableName, List<string> columns)
        {
            var existingColumns = GetTableColumns(conn, tableName);

            foreach (var column in columns)
            {
                var sanitizedColumn = SanitizeColumnName(column);
                if (!existingColumns.Contains(sanitizedColumn))
                {
                    try
                    {
                        var addColumnSql = $@"ALTER TABLE {tableName} ADD COLUMN ""{sanitizedColumn}"" TEXT";
                        ExecuteNonQuery(conn, addColumnSql);
                        Console.WriteLine($"Added column: {sanitizedColumn} to {tableName}");
                    }
                    catch (PostgresException ex) when (ex.SqlState == "42701") // Column already exists
                    {
                        // Ignore this specific error
                        Console.WriteLine($"Column {sanitizedColumn} already exists in {tableName} - skipping");
                    }
                }
            }
        }

        private List<string> GetTableColumns(NpgsqlConnection conn, string tableName)
        {
            var columns = new List<string>();
            var schema = tableName.Split('.')[0];
            var table = tableName.Split('.')[1];

            var sql = $@"
                SELECT column_name 
                FROM information_schema.columns 
                WHERE table_schema = '{schema}'
                AND table_name = '{table}'";

            using var cmd = new NpgsqlCommand(sql, conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                columns.Add(reader.GetString(0));
            }
            return columns;
        }

        private int ImportData(NpgsqlConnection conn, string tableName, DataTable data, List<string> columns, string filePath, string fileHash)
        {
            using var transaction = conn.BeginTransaction();
            try
            {
                int rowsImported = 0;
                var insertSql = BuildInsertStatement(tableName, columns);

                foreach (DataRow row in data.Rows)
                {
                    using var cmd = new NpgsqlCommand(insertSql, conn, transaction);

                    for (int i = 0; i < columns.Count; i++)
                    {
                        var value = row[i];
                        cmd.Parameters.AddWithValue($"@p{i}", value == DBNull.Value ? (object)DBNull.Value : value.ToString().Trim());
                    }

                    rowsImported += cmd.ExecuteNonQuery();
                    if (rowsImported % 100 == 0) Console.Write(".");
                }

                transaction.Commit();
                return rowsImported;
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        private string BuildInsertStatement(string tableName, List<string> columns)
        {
            var columnList = columns.Select(c => $@"""{SanitizeColumnName(c)}""");
            var valueParams = columns.Select((_, i) => $"@p{i}");

            return $@"
        INSERT INTO {tableName} ({string.Join(", ", columnList)})
        VALUES ({string.Join(", ", valueParams)})";
        }

        private bool IsFileProcessed(NpgsqlConnection conn, string tableName, string fileName, string fileHash, string targetTable)
        {
            var sql = $@"
                SELECT 1 FROM {tableName} 
                WHERE file_name = @name 
                AND target_table = @table
                AND file_hash = @hash";

            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@name", fileName);
            cmd.Parameters.AddWithValue("@table", targetTable);
            cmd.Parameters.AddWithValue("@hash", fileHash);

            using var reader = cmd.ExecuteReader();
            return reader.HasRows;
        }

        private void MarkFileAsProcessed(NpgsqlConnection conn, string tableName, string fileName, string fileHash, long fileSize, string targetTable)
        {
            var sql = $@"
                INSERT INTO {tableName} (file_name, file_hash, file_size, target_table)
                VALUES (@name, @hash, @size, @table)
                ON CONFLICT (file_name, target_table)
                DO UPDATE SET 
                    file_hash = EXCLUDED.file_hash,
                    file_size = EXCLUDED.file_size,
                    processed_at = CURRENT_TIMESTAMP";

            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@name", fileName);
            cmd.Parameters.AddWithValue("@hash", fileHash);
            cmd.Parameters.AddWithValue("@size", fileSize);
            cmd.Parameters.AddWithValue("@table", targetTable);
            cmd.ExecuteNonQuery();
        }

        private string SanitizeColumnName(string columnName)
        {
            return new string(columnName
                .Where(c => char.IsLetterOrDigit(c) || c == '_')
                .ToArray())
                .Trim()
                .ToLower();
        }

        private void ExecuteNonQuery(NpgsqlConnection conn, string sql)
        {
            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        private string BuildConnectionString(DatabaseConfig dbConfig)
        {
            return new NpgsqlConnectionStringBuilder
            {
                Host = dbConfig.Host,
                Port = dbConfig.Port,
                Username = dbConfig.Username,
                Password = dbConfig.Password,
                Database = dbConfig.Database,
                SslMode = SslMode.Require,
                TrustServerCertificate = true,
                Pooling = true,
                MinPoolSize = 1,
                MaxPoolSize = 20,
                Timeout = 30
            }.ToString();
        }
    }

    public class AppConfig
    {
        public List<FolderConfig> FolderConfigs { get; set; }
        public DatabaseConfig DatabaseConfig { get; set; }
    }

    public class FolderConfig
    {
        public string FolderPath { get; set; }
        public string TableName { get; set; }
        public string[] MandatoryColumns { get; set; }
    }

    public class DatabaseConfig
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string Database { get; set; }
        public string ProcessedFilesTable { get; set; }
    }

    public class ImportResult
    {
        public bool Success { get; set; }
        public string FileName { get; set; }
        public string FolderPath { get; set; }
        public string TableName { get; set; }
        public int RowsImported { get; set; }
        public int AdditionalColumnsCount { get; set; }
        public string ErrorMessage { get; set; }
    }
}