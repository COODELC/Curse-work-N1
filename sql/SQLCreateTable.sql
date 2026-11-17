-- =============================================================
-- Создание базы данных «Автомастерская» для SQL Server.
-- База: cursevork
-- Файл соответствует моделям EF Core в Ktpt.CourseOrder.Data.
-- =============================================================

USE cursevork;
GO

-- Удаление таблиц в порядке, обратном зависимостям.
DROP TABLE IF EXISTS Orders;
DROP TABLE IF EXISTS ItemReview;
DROP TABLE IF EXISTS OrderStatus;
DROP TABLE IF EXISTS Item;
DROP TABLE IF EXISTS Category;
DROP TABLE IF EXISTS Users;
DROP TABLE IF EXISTS Roles;
GO

CREATE TABLE Roles (
    Id_Role INT IDENTITY(1, 1) NOT NULL,
    Name_Role NVARCHAR(30) NOT NULL,

    CONSTRAINT PK_Roles PRIMARY KEY (Id_Role)
);
GO

CREATE TABLE Users (
    Id_User INT IDENTITY(1, 1) NOT NULL,
    FirstName NVARCHAR(30) NOT NULL,
    SecondName NVARCHAR(30) NULL,
    Age INT NULL,
    Login NVARCHAR(30) NOT NULL,
    Password NVARCHAR(255) NOT NULL,
    Role_Id INT NOT NULL,

    CONSTRAINT PK_Users PRIMARY KEY (Id_User),
    CONSTRAINT UQ_Users_Login UNIQUE (Login),

    CONSTRAINT FK_Users_Roles_Role_Id
        FOREIGN KEY (Role_Id) REFERENCES Roles (Id_Role)
        ON DELETE NO ACTION
);
GO

CREATE TABLE Category (
    Id_Category INT IDENTITY(1, 1) NOT NULL,
    Name_Category NVARCHAR(50) NOT NULL,

    CONSTRAINT PK_Category PRIMARY KEY (Id_Category)
);
GO

CREATE TABLE Item (
    Id_Item INT IDENTITY(1, 1) NOT NULL,
    Name_Item NVARCHAR(80) NOT NULL,
    NameArticule NVARCHAR(80) NULL,
    Category_Id INT NOT NULL,
    Price DECIMAL(18, 2) NOT NULL CONSTRAINT DF_Item_Price DEFAULT (0),
    Quantity INT NOT NULL CONSTRAINT DF_Item_Quantity DEFAULT (0),
    ImagePath NVARCHAR(200) NULL,

    CONSTRAINT PK_Item PRIMARY KEY (Id_Item),

    CONSTRAINT FK_Item_Category_Category_Id
        FOREIGN KEY (Category_Id) REFERENCES Category (Id_Category)
        ON DELETE NO ACTION
);
GO

CREATE TABLE ItemReview (
    Id_Review INT IDENTITY(1, 1) NOT NULL,
    Review NVARCHAR(MAX) NOT NULL,
    Item_Id INT NOT NULL,
    User_Id INT NULL,

    CONSTRAINT PK_ItemReview PRIMARY KEY (Id_Review),

    CONSTRAINT FK_ItemReview_Item_Item_Id
        FOREIGN KEY (Item_Id) REFERENCES Item (Id_Item)
        ON DELETE CASCADE,

    CONSTRAINT FK_ItemReview_Users_User_Id
        FOREIGN KEY (User_Id) REFERENCES Users (Id_User)
        ON DELETE SET NULL
);
GO

CREATE TABLE OrderStatus (
    Id_status INT IDENTITY(1, 1) NOT NULL,
    Name_Status NVARCHAR(100) NOT NULL,

    CONSTRAINT PK_OrderStatus PRIMARY KEY (Id_status)
);
GO

CREATE TABLE Orders (
    Id_Order INT IDENTITY(1, 1) NOT NULL,
    Date_Order DATETIME2 NOT NULL,
    Item_Id INT NOT NULL,
    UserId INT NOT NULL,
    Status_Id INT NOT NULL CONSTRAINT DF_Orders_Status_Id DEFAULT (1),

    CONSTRAINT PK_Orders PRIMARY KEY (Id_Order),

    CONSTRAINT FK_Orders_Item_Item_Id
        FOREIGN KEY (Item_Id) REFERENCES Item (Id_Item)
        ON DELETE NO ACTION,

    CONSTRAINT FK_Orders_Users_UserId
        FOREIGN KEY (UserId) REFERENCES Users (Id_User)
        ON DELETE NO ACTION,

    CONSTRAINT FK_Orders_OrderStatus_Status_Id
        FOREIGN KEY (Status_Id) REFERENCES OrderStatus (Id_status)
        ON DELETE NO ACTION
);
GO

CREATE INDEX IX_Users_Login ON Users (Login);
CREATE INDEX IX_Item_Category_Id ON Item (Category_Id);
CREATE INDEX IX_ItemReview_Item_Id ON ItemReview (Item_Id);
CREATE INDEX IX_ItemReview_User_Id ON ItemReview (User_Id);
CREATE INDEX IX_Orders_Item_Id ON Orders (Item_Id);
CREATE INDEX IX_Orders_UserId ON Orders (UserId);
CREATE INDEX IX_Orders_Status_Id ON Orders (Status_Id);
GO