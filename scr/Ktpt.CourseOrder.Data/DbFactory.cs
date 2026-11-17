using System.IO;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Data;


public sealed class DbFactory
{
    private readonly string _connectionString;

    private bool _useLocalFile;

    public DbFactory(string connectionString)
    {
        _connectionString = connectionString.Trim();

        if (_connectionString.Length == 0)
        {
            throw new InvalidOperationException(
                "В файле Connection.txt не указана строка подключения.");
        }
    }

    public static DbFactory CreateFromFile()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Connection.txt");

        if (!File.Exists(path))
        {
            throw new FileNotFoundException(
                "Файл Connection.txt не найден рядом с приложением.",
                path);
        }

        return new DbFactory(File.ReadAllText(path));
    }

    public bool IsLocalFileDatabase =>
        _useLocalFile
        || _connectionString.StartsWith("Data Source=", StringComparison.OrdinalIgnoreCase)
        || _connectionString.StartsWith("Filename=", StringComparison.OrdinalIgnoreCase);

    public string ConnectionString => _connectionString;


    public void SwitchToLocalFile()
    {
        _useLocalFile = true;
    }

    public WorkshopDbContext Create()
    {
        var builder = new DbContextOptionsBuilder<WorkshopDbContext>();

        if (IsLocalFileDatabase)
        {
            var path = Path.Combine(AppContext.BaseDirectory, "autoshop.db");

            builder.UseSqlite($"Data Source={path}");
        }
        else
        {
            builder.UseSqlServer(_connectionString);
        }

        return new WorkshopDbContext(builder.Options);
    }
}
