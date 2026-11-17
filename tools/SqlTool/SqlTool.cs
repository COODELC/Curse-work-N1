#:package Microsoft.Data.SqlClient@6.1.6
#:package Microsoft.Data.Sqlite@11.0.0

// SQL-инструмент для курсового проекта.
// Запуск:  dotnet run tools/SqlTool/SqlTool.cs -- --query "SELECT ..."
//          dotnet run tools/SqlTool/SqlTool.cs -- --file script.sql
//          dotnet run tools/SqlTool/SqlTool.cs -- --list
//          dotnet run tools/SqlTool/SqlTool.cs -- --dbs
//          dotnet run tools/SqlTool/SqlTool.cs -- --sqlite <путь к .db> --query "..."
//          dotnet run tools/SqlTool/SqlTool.cs -- (запрос читается из STDIN)

using System.Data.Common;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;

const int MaxCellWidth = 200;

const string HelpText = """
    Использование:
      dotnet run tools/SqlTool/SqlTool.cs -- --query "SELECT * FROM Users"
      dotnet run tools/SqlTool/SqlTool.cs -- --file script.sql
      dotnet run tools/SqlTool/SqlTool.cs -- --list            список таблиц и строк
      dotnet run tools/SqlTool/SqlTool.cs -- --dbs             список баз данных сервера
      dotnet run tools/SqlTool/SqlTool.cs -- --sqlite <db> --query "SELECT ..."
      echo "SELECT 1" | dotnet run tools/SqlTool/SqlTool.cs --

    Параметры:
      --query <sql>   SQL-запрос
      --file <path>   SQL-файл со скриптом
      --conn <path>   путь к Connection.txt или строка подключения
      --sqlite <path> работать с локальной базой SQLite вместо SQL Server
      --csv           вывод в CSV
      --wide          не обрезать длинные значения
    """;

var options = ParseArgs(args);

if (options.ContainsKey("help"))
{
    Console.WriteLine(HelpText);
    return;
}

var repoRoot = FindRepoRoot();

if (options.TryGetValue("sqlite", out var sqlitePath))
{
    await RunSqlite(sqlitePath, options);
    return;
}

var connectionString = ResolveConnectionString(repoRoot, options);

if (options.ContainsKey("dbs"))
{
    RunDatabaseList(connectionString);
    return;
}

if (options.ContainsKey("list"))
{
    RunSchemaList(connectionString);
    return;
}

var query = options.GetValueOrDefault("query");
var file = options.GetValueOrDefault("file");

if (string.IsNullOrWhiteSpace(query) && string.IsNullOrWhiteSpace(file))
{
    query = Console.In.ReadToEnd();
}

if (string.IsNullOrWhiteSpace(query) && string.IsNullOrWhiteSpace(file))
{
    Console.Error.WriteLine("Нет запроса. Укажите --query, --file или подайте текст в STDIN.");
    Console.Error.WriteLine(HelpText);
    Environment.Exit(2);
}

await using var connection = new SqlConnection(connectionString);

try
{
    await connection.OpenAsync();
}
catch (Exception ex)
{
    Console.Error.WriteLine("ОШИБКА ПОДКЛЮЧЕНИЯ: " + ex.Message);
    Environment.Exit(1);
}

Console.WriteLine($"-- Подключено: {connection.Database} @ {connection.DataSource}");
Console.WriteLine();

var script = file is not null ? await File.ReadAllTextAsync(file) : query ?? string.Empty;

await RunScriptAsync(connection, script, options.ContainsKey("csv"), options.ContainsKey("wide"));

async Task RunScriptAsync(DbConnection conn, string text, bool asCsv, bool wide)
{
    var batches = SplitBatches(text);
    var total = batches.Count;
    var index = 0;

    foreach (var batch in batches)
    {
        index++;

        if (total > 1)
        {
            Console.WriteLine($"=== Батч {index} из {total} ===");
        }

        await using var command = conn.CreateCommand();
        command.CommandText = batch;
        command.CommandTimeout = 120;

        try
        {
            using var reader = await command.ExecuteReaderAsync();

            var printed = await PrintResults(reader, asCsv, wide);

            while (await reader.NextResultAsync())
            {
                printed = await PrintResults(reader, asCsv, wide) || printed;
            }

            if (!printed)
            {
                Console.WriteLine("(выполнено, строки не возвращены)");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("ОШИБКА SQL: " + ex.Message);
            Environment.Exit(1);
        }

        Console.WriteLine();
    }
}

List<string> SplitBatches(string text)
{
    var batches = new List<string>();
    var current = new List<string>();

    foreach (var line in text.Replace("\r\n", "\n").Split('\n'))
    {
        if (line.Trim().Equals("GO", StringComparison.OrdinalIgnoreCase))
        {
            AddBatch();
            continue;
        }

        current.Add(line);
    }

    AddBatch();

    return batches;

    void AddBatch()
    {
        var batch = string.Join("\n", current).Trim();

        if (batch.Length > 0)
        {
            batches.Add(batch);
        }

        current.Clear();
    }
}

async Task<bool> PrintResults(DbDataReader reader, bool asCsv, bool wide)
{
    var columns = new string[reader.FieldCount];
    for (var i = 0; i < reader.FieldCount; i++)
    {
        columns[i] = reader.GetName(i);
    }

    if (reader.FieldCount == 0)
    {
        return false;
    }

    if (asCsv)
    {
        Console.WriteLine(string.Join(",", columns.Select(c => "\"" + c.Replace("\"", "\"\"") + "\"")));
    }
    else
    {
        Console.WriteLine(string.Join(" | ", columns));
        Console.WriteLine(string.Join("-+-", columns.Select(c => new string('-', Math.Min(c.Length, 30)))));
    }

    var count = 0;

    while (await reader.ReadAsync())
    {
        var values = new string[reader.FieldCount];

        for (var i = 0; i < reader.FieldCount; i++)
        {
            var value = await reader.IsDBNullAsync(i)
                ? "NULL"
                : Convert.ToString(reader.GetValue(i)) ?? string.Empty;

            values[i] = value;
        }

        if (asCsv)
        {
            Console.WriteLine(string.Join(",", values.Select(v => "\"" + v.Replace("\"", "\"\"") + "\"")));
        }
        else
        {
            Console.WriteLine(string.Join(" | ", values.Select(v =>
                wide || v.Length <= MaxCellWidth ? v : v[..MaxCellWidth] + "...")));
        }

        count++;
    }

    Console.WriteLine();
    Console.WriteLine($"(строк: {count})");

    return true;
}

void RunDatabaseList(string connectionString)
{
    var master = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = "master" };

    using var connection = new SqlConnection(master.ConnectionString);
    connection.Open();

    using var command = new SqlCommand("SELECT name FROM sys.databases ORDER BY name", connection);
    using var reader = command.ExecuteReader();

    Console.WriteLine("Базы данных на сервере:");
    while (reader.Read())
    {
        Console.WriteLine("  " + reader.GetString(0));
    }
}

void RunSchemaList(string connectionString)
{
    using var connection = new SqlConnection(connectionString);
    connection.Open();

    var tables = new List<string>();

    using (var command = new SqlCommand("SELECT name FROM sys.tables ORDER BY name", connection))
    using (var reader = command.ExecuteReader())
    {
        while (reader.Read())
        {
            tables.Add(reader.GetString(0));
        }
    }

    Console.WriteLine("Таблицы базы:");

    foreach (var table in tables)
    {
        int rows = 0;
        int columns = 0;

        using (var command = new SqlCommand($"SELECT COUNT(*) FROM [{table}]", connection))
        {
            rows = Convert.ToInt32(command.ExecuteScalar());
        }

        using (var command = new SqlCommand(
            "SELECT COUNT(*) FROM sys.columns c JOIN sys.tables t ON t.object_id = c.object_id WHERE t.name = @t",
            connection))
        {
            command.Parameters.AddWithValue("@t", table);
            columns = Convert.ToInt32(command.ExecuteScalar());
        }

        Console.WriteLine($"  {table,-16} строк: {rows,-6} колонок: {columns}");
    }
}

async Task RunSqlite(string path, Dictionary<string, string> opts)
{
    var fullPath = Path.IsPathRooted(path) ? path : Path.Combine(repoRoot, path);

    if (!File.Exists(fullPath))
    {
        Console.Error.WriteLine("Файл базы не найден: " + fullPath);
        Environment.Exit(1);
    }

    await using var sqliteConnection = new SqliteConnection($"Data Source={fullPath}");
    await sqliteConnection.OpenAsync();

    Console.WriteLine($"-- SQLite: {fullPath}");
    Console.WriteLine();

    if (opts.ContainsKey("list"))
    {
        var names = new List<string>();

        using (var listCommand = sqliteConnection.CreateCommand())
        {
            listCommand.CommandText =
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name";

            using var listReader = await listCommand.ExecuteReaderAsync();

            while (await listReader.ReadAsync())
            {
                names.Add(listReader.GetString(0));
            }
        }

        Console.WriteLine("Таблицы базы:");

        foreach (var name in names)
        {
            using var countCommand = sqliteConnection.CreateCommand();
            countCommand.CommandText = $"SELECT COUNT(*) FROM [{name}]";

            var rows = Convert.ToInt32(await countCommand.ExecuteScalarAsync());

            Console.WriteLine($"  {name,-16} строк: {rows}");
        }

        return;
    }

    var query = opts.GetValueOrDefault("query");
    var file = opts.GetValueOrDefault("file");

    if (string.IsNullOrWhiteSpace(query) && string.IsNullOrWhiteSpace(file))
    {
        query = Console.In.ReadToEnd();
    }

    if (string.IsNullOrWhiteSpace(query))
    {
        Console.Error.WriteLine("Нет запроса. Укажите --query или --file.");
        Environment.Exit(2);
    }

    var script = file is not null ? await File.ReadAllTextAsync(file) : query ?? string.Empty;

    await RunScriptAsync(sqliteConnection, script, opts.ContainsKey("csv"), opts.ContainsKey("wide"));
}

string ResolveConnectionString(string repoRoot, Dictionary<string, string> opts)
{
    if (opts.TryGetValue("conn", out var explicitConn))
    {
        var text = File.Exists(explicitConn) ? File.ReadAllText(explicitConn) : explicitConn;
        return text.Trim();
    }

    var env = Environment.GetEnvironmentVariable("CURSORVERK_CONN");
    if (!string.IsNullOrWhiteSpace(env))
    {
        return env.Trim();
    }

    var candidate = Path.Combine(repoRoot, "scr", "Ktpt.CourseOrder", "Connection.txt");
    if (File.Exists(candidate))
    {
        return File.ReadAllText(candidate).Trim();
    }

    throw new FileNotFoundException("Не найден Connection.txt", candidate);
}

Dictionary<string, string> ParseArgs(string[] raw)
{
    var result = new Dictionary<string, string>();

    for (var i = 0; i < raw.Length; i++)
    {
        var token = raw[i];

        if (!token.StartsWith("--"))
        {
            continue;
        }

        var key = token[2..].ToLowerInvariant();

        if (key is "csv" or "wide" or "list" or "dbs" or "help")
        {
            result[key] = "1";
            continue;
        }

        if (i + 1 < raw.Length && !raw[i + 1].StartsWith("--"))
        {
            result[key] = raw[i + 1];
            i++;
        }
        else
        {
            result[key] = "1";
        }
    }

    return result;
}

string FindRepoRoot()
{
    var dir = new DirectoryInfo(Directory.GetCurrentDirectory());

    while (dir is not null)
    {
        if (Directory.Exists(Path.Combine(dir.FullName, "sql")) &&
            Directory.Exists(Path.Combine(dir.FullName, "scr")))
        {
            return dir.FullName;
        }

        dir = dir.Parent;
    }

    return Directory.GetCurrentDirectory();
}
