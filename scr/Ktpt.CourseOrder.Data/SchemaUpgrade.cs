using System.Data;
using System.Data.Common;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Data;

/// <summary>
/// Обновление схемы уже существующей базы данных.
///
/// Метод EnsureCreated() создаёт только отсутствующие таблицы и не меняет
/// созданные ранее, поэтому новые колонки добавляются здесь вручную.
/// Каждое обновление выполняется один раз: повторный запуск ничего не меняет.
/// </summary>
public static class SchemaUpgrade
{
    /// <summary>
    /// Применяет все обновления схемы к базе.
    /// </summary>
    public static void Apply(WorkshopDbContext db)
    {
        AddReviewAuthorColumn(db);
        AddItemDescriptionColumn(db);
    }

    /// <summary>
    /// Добавляет в таблицу Item колонку Description с описанием товара.
    /// </summary>
    private static void AddItemDescriptionColumn(WorkshopDbContext db)
    {
        if (ColumnExists(db, "Item", "Description"))
        {
            return;
        }

        var provider = db.Database.ProviderName;

        if (provider == "Microsoft.EntityFrameworkCore.SqlServer")
        {
            db.Database.ExecuteSqlRaw(
                "ALTER TABLE [Item] ADD [Description] nvarchar(1000) NULL;");
        }
        else if (provider == "Microsoft.EntityFrameworkCore.Sqlite")
        {
            db.Database.ExecuteSqlRaw(
                "ALTER TABLE \"Item\" ADD COLUMN \"Description\" TEXT NULL;");
        }
    }

    /// <summary>
    /// Добавляет в таблицу ItemReview колонку User_Id со ссылкой на Users.
    /// </summary>
    private static void AddReviewAuthorColumn(WorkshopDbContext db)
    {
        if (ColumnExists(db, "ItemReview", "User_Id"))
        {
            return;
        }

        var provider = db.Database.ProviderName;

        if (provider == "Microsoft.EntityFrameworkCore.SqlServer")
        {
            db.Database.ExecuteSqlRaw(
                "ALTER TABLE [ItemReview] ADD [User_Id] int NULL;");

            db.Database.ExecuteSqlRaw(
                "ALTER TABLE [ItemReview] ADD CONSTRAINT [FK_ItemReview_Users_User_Id] " +
                "FOREIGN KEY ([User_Id]) REFERENCES [Users] ([Id_User]) ON DELETE SET NULL;");

            return;
        }

        if (provider == "Microsoft.EntityFrameworkCore.Sqlite")
        {
            RebuildItemReviewTable(db);
        }
    }

    /// <summary>
    /// SQLite не умеет добавлять внешний ключ к существующей таблице,
    /// поэтому таблица пересоздаётся целиком.
    /// </summary>
    private static void RebuildItemReviewTable(WorkshopDbContext db)
    {
        var connection = db.Database.GetDbConnection();
        var wasOpen = connection.State == ConnectionState.Open;

        if (!wasOpen)
        {
            connection.Open();
        }

        try
        {
            Execute(connection, "PRAGMA foreign_keys=OFF;");

            Execute(connection, """
                CREATE TABLE "ItemReview__new" (
                    "Id_Review" INTEGER NOT NULL CONSTRAINT "PK_ItemReview" PRIMARY KEY AUTOINCREMENT,
                    "Review" TEXT NOT NULL,
                    "Item_Id" INTEGER NOT NULL,
                    "User_Id" INTEGER NULL,
                    CONSTRAINT "FK_ItemReview_Item_Item_Id" FOREIGN KEY ("Item_Id") REFERENCES "Item" ("Id_Item") ON DELETE CASCADE,
                    CONSTRAINT "FK_ItemReview_Users_User_Id" FOREIGN KEY ("User_Id") REFERENCES "Users" ("Id_User") ON DELETE SET NULL
                );
                """);

            Execute(connection, """
                INSERT INTO "ItemReview__new" ("Id_Review", "Review", "Item_Id", "User_Id")
                SELECT "Id_Review", "Review", "Item_Id", NULL FROM "ItemReview";
                """);

            Execute(connection, "DROP TABLE \"ItemReview\";");

            Execute(connection, "ALTER TABLE \"ItemReview__new\" RENAME TO \"ItemReview\";");

            Execute(connection, "PRAGMA foreign_keys=ON;");
        }
        finally
        {
            if (!wasOpen)
            {
                connection.Close();
            }
        }
    }

    private static void Execute(DbConnection connection, string sql)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Проверяет наличие колонки в таблице.
    /// </summary>
    private static bool ColumnExists(WorkshopDbContext db, string table, string column)
    {
        var provider = db.Database.ProviderName;

        var connection = db.Database.GetDbConnection();
        var wasOpen = connection.State == ConnectionState.Open;

        if (!wasOpen)
        {
            connection.Open();
        }

        try
        {
            using var command = connection.CreateCommand();

            if (provider == "Microsoft.EntityFrameworkCore.SqlServer")
            {
                command.CommandText =
                    "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS " +
                    "WHERE TABLE_NAME = @table AND COLUMN_NAME = @column;";

                AddParameter(command, "@table", table);
                AddParameter(command, "@column", column);

                return Convert.ToInt32(command.ExecuteScalar()) > 0;
            }

            if (provider == "Microsoft.EntityFrameworkCore.Sqlite")
            {
                command.CommandText = $"PRAGMA table_info(\"{table}\");";

                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var name = reader.GetString(1);

                    if (string.Equals(name, column, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }

            return false;
        }
        finally
        {
            if (!wasOpen)
            {
                connection.Close();
            }
        }
    }

    private static void AddParameter(DbCommand command, string name, string value)
    {
        var parameter = command.CreateParameter();
        parameter.ParameterName = name;
        parameter.Value = value;

        command.Parameters.Add(parameter);
    }
}
