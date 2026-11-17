using Ktpt.CourseOrder.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Data;

/// <summary>
/// Контекст Entity Framework Core — описание логической модели данных.
/// Один и тот же контекст работает и с SQL Server, и с локальной базой SQLite.
/// </summary>
public class WorkshopDbContext : DbContext
{
    public WorkshopDbContext(DbContextOptions<WorkshopDbContext> options)
        : base(options)
    {
    }

    public DbSet<Roles> Roles => Set<Roles>();

    public DbSet<Users> Users => Set<Users>();

    public DbSet<Category> Categories => Set<Category>();

    public DbSet<Item> Items => Set<Item>();

    public DbSet<ItemReview> ItemReviews => Set<ItemReview>();

    public DbSet<OrderStatus> OrderStatuses => Set<OrderStatus>();

    public DbSet<Orders> Orders => Set<Orders>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<Roles>(entity =>
        {
            entity.ToTable("Roles");

            entity.HasKey(role => role.Id_Role);

            entity.Property(role => role.Id_Role)
                .ValueGeneratedOnAdd();

            entity.Property(role => role.Name_Role)
                .HasMaxLength(30)
                .IsRequired();

            entity.HasMany(role => role.Users)
                .WithOne(user => user.Role)
                .HasForeignKey(user => user.Role_Id)
                .OnDelete(DeleteBehavior.Restrict);
        });

        modelBuilder.Entity<Users>(entity =>
        {
            entity.ToTable("Users");

            entity.HasKey(user => user.Id_User);

            entity.Property(user => user.Id_User)
                .ValueGeneratedOnAdd();

            entity.Property(user => user.FirstName)
                .HasMaxLength(30)
                .IsRequired();

            entity.Property(user => user.SecondName)
                .HasMaxLength(30)
                .IsRequired(false);

            entity.Property(user => user.Age)
                .IsRequired(false);

            entity.Property(user => user.Login)
                .HasMaxLength(30)
                .IsRequired();

            entity.Property(user => user.Password)
                .HasMaxLength(255)
                .IsRequired();
        });

        modelBuilder.Entity<Category>(entity =>
        {
            entity.ToTable("Category");

            entity.HasKey(category => category.Id_Category);

            entity.Property(category => category.Id_Category)
                .ValueGeneratedOnAdd();

            entity.Property(category => category.Name_Category)
                .HasMaxLength(50)
                .IsRequired();

            entity.HasMany(category => category.Items)
                .WithOne(item => item.Category)
                .HasForeignKey(item => item.Category_Id)
                .OnDelete(DeleteBehavior.Restrict);
        });

        modelBuilder.Entity<Item>(entity =>
        {
            entity.ToTable("Item");

            entity.HasKey(item => item.Id_Item);

            entity.Property(item => item.Id_Item)
                .ValueGeneratedOnAdd();

            entity.Property(item => item.Name_Item)
                .HasMaxLength(80)
                .IsRequired();

            entity.Property(item => item.NameArticule)
                .HasMaxLength(80)
                .IsRequired(false);

            entity.Property(item => item.Price)
                .HasPrecision(18, 2)
                .IsRequired();

            entity.Property(item => item.Quantity)
                .IsRequired();

            entity.Property(item => item.Description)
                .HasMaxLength(1000)
                .IsRequired(false);

            entity.Property(item => item.ImagePath)
                .HasMaxLength(200)
                .IsRequired(false);

            entity.HasMany(item => item.Reviews)
                .WithOne(review => review.Item)
                .HasForeignKey(review => review.Item_Id)
                .OnDelete(DeleteBehavior.Cascade);
        });

        modelBuilder.Entity<ItemReview>(entity =>
        {
            entity.ToTable("ItemReview");

            entity.HasKey(review => review.Id_Review);

            entity.Property(review => review.Id_Review)
                .ValueGeneratedOnAdd();

            entity.Property(review => review.Review)
                .IsRequired();

            // Автор отзыва. При удалении пользователя отзыв остаётся, но становится гостевым.
            entity.HasOne(review => review.User)
                .WithMany(user => user.Reviews)
                .HasForeignKey(review => review.User_Id)
                .OnDelete(DeleteBehavior.SetNull);
        });

        modelBuilder.Entity<OrderStatus>(entity =>
        {
            entity.ToTable("OrderStatus");

            entity.HasKey(status => status.Id_status);

            entity.Property(status => status.Id_status)
                .ValueGeneratedOnAdd();

            entity.Property(status => status.Name_Status)
                .HasMaxLength(100)
                .IsRequired();
        });

        modelBuilder.Entity<Orders>(entity =>
        {
            entity.ToTable("Orders");

            entity.HasKey(order => order.Id_Order);

            entity.Property(order => order.Id_Order)
                .ValueGeneratedOnAdd();

            entity.Property(order => order.Date_Order)
                .IsRequired();

            entity.Property(order => order.Status_Id)
                .HasDefaultValue(1)
                .IsRequired();

            entity.HasOne(order => order.Item)
                .WithMany(item => item.Orders)
                .HasForeignKey(order => order.Item_Id)
                .OnDelete(DeleteBehavior.Restrict);

            entity.HasOne(order => order.User)
                .WithMany(user => user.Orders)
                .HasForeignKey(order => order.UserId)
                .OnDelete(DeleteBehavior.Restrict);

            entity.HasOne(order => order.Status)
                .WithMany(status => status.Orders)
                .HasForeignKey(order => order.Status_Id)
                .OnDelete(DeleteBehavior.Restrict);
        });
    }
}
