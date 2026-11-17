using System;

namespace Ktpt.CourseOrder.Auth;

public sealed class AuthSession
{
    public AppUser? CurrentUser { get; private set; }

    public UserRole Role =>
        CurrentUser?.Role ?? UserRole.Guest;

    public bool IsAuthenticated =>
        CurrentUser is not null;

    public event Action? Changed;

    public void SignIn(AppUser user)
    {
        CurrentUser = user;
        Changed?.Invoke();
    }

    public void SignOut()
    {
        CurrentUser = null;
        Changed?.Invoke();
    }

    public bool Can(Permission permission)
    {
        return Permissions.Allows(Role, permission);
    }

    public void Demand(Permission permission)
    {
        if (!Can(permission))
        {
            throw new UnauthorizedAccessException(
                "У вас нет прав для выполнения этой операции.");
        }
    }
}