namespace Ktpt.CourseOrder.Auth;

public enum Permission
{
    ViewCatalog,
    PlaceOrder,
    ViewOwnOrders,
    EditOwnProfile,
    AddReview,
    ManageProducts,
    ManageCategories,
    ManageOrders,
    ManageUsers
}

public static class Permissions
{
    public static bool Allows(
        UserRole role,
        Permission permission)
    {
        return permission switch
        {
            Permission.ViewCatalog => true,

            Permission.PlaceOrder
                or Permission.ViewOwnOrders
                or Permission.EditOwnProfile
                or Permission.AddReview
                => role is UserRole.Customer
                    or UserRole.Manager
                    or UserRole.Administrator,

            // Изменять каталог товаров может только администратор.
            Permission.ManageProducts
                or Permission.ManageCategories
                => role == UserRole.Administrator,

            // Менеджер может обрабатывать заказы, но не изменять товары.
            Permission.ManageOrders
                => role is UserRole.Manager
                    or UserRole.Administrator,

            Permission.ManageUsers
                => role == UserRole.Administrator,

            _ => false
        };
    }
}
