namespace Ktpt.CourseOrder.Auth;

public enum UserRole
{
    Guest,
    Customer,
    Manager,
    Administrator
}

public sealed record AppUser(
    int Id,
    string Login,
    string FirstName,
    string? SecondName,
    int? Age,
    UserRole Role);
