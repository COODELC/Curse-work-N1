-- Резервная копия базы cursevork ДО пересоздания по модели приложения
-- Создано: 2026-11-13 21:48:53
SET NOCOUNT ON;

-- Roles
DELETE FROM Roles;
INSERT INTO Roles (Id_Role, Name_Role) VALUES (1, N'Гость');
INSERT INTO Roles (Id_Role, Name_Role) VALUES (2, N'Покупатель');
INSERT INTO Roles (Id_Role, Name_Role) VALUES (3, N'Менеджер');
INSERT INTO Roles (Id_Role, Name_Role) VALUES (4, N'Администратор');

-- Users
DELETE FROM Users;
INSERT INTO Users (Id_User, FirstName, SecondName, Age, Login, Password, Role_Id) VALUES (1, N'Андрей', N'Копейкин', 18, N'ander', N'8942', 4);
INSERT INTO Users (Id_User, FirstName, SecondName, Age, Login, Password, Role_Id) VALUES (2, N'Мария', N'Иванова', 25, N'maria', N'1234', 2);
INSERT INTO Users (Id_User, FirstName, SecondName, Age, Login, Password, Role_Id) VALUES (3, N'Иван', N'Петров', 30, N'ivan', N'5678', 3);
INSERT INTO Users (Id_User, FirstName, SecondName, Age, Login, Password, Role_Id) VALUES (4, N'Ольга', N'Смирнова', 22, N'olga', N'abcd', 2);
INSERT INTO Users (Id_User, FirstName, SecondName, Age, Login, Password, Role_Id) VALUES (5, N'Сергей', N'Волков', 41, N'sergey', N'pass1', 4);

-- Category
DELETE FROM Category;
INSERT INTO Category (Id_Category, Name_Category) VALUES (1, N'Двигатель');
INSERT INTO Category (Id_Category, Name_Category) VALUES (2, N'Тормозная система');
INSERT INTO Category (Id_Category, Name_Category) VALUES (3, N'Подвеска');
INSERT INTO Category (Id_Category, Name_Category) VALUES (4, N'Электрика');
INSERT INTO Category (Id_Category, Name_Category) VALUES (5, N'Масла и жидкости');
INSERT INTO Category (Id_Category, Name_Category) VALUES (6, N'Кузовные детали');

-- Item
DELETE FROM Item;
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (1, N'Масляный насос', N'Помпа', 1);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (2, N'Воздушный фильтр', N'Фильтр очистки воздуха', 1);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (3, N'Тормозные колодки', N'Колодки', 2);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (4, N'Тормозной диск', N'Тормозной ротор', 2);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (5, N'Амортизатор', N'Стойка', 3);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (6, N'Шаровая опора', N'Опора шаровая', 3);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (7, N'Аккумулятор', N'АКБ', 4);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (8, N'Свеча зажигания', N'Свеча', 4);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (9, N'Моторное масло', N'Машинное масло', 5);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (10, N'Антифриз', N'Охлаждающая жидкость', 5);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (11, N'Бампер', N'Буфер', 6);
INSERT INTO Item (Id_Item, Name_Item, NameArticule, Category_Id) VALUES (12, N'Крыло', N'Оперение', 6);

-- Orders
DELETE FROM Orders;
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (1, N'2026-10-01 00:00:00', 1, 2);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (2, N'2026-10-03 00:00:00', 3, 2);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (3, N'2026-10-05 00:00:00', 5, 3);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (4, N'2026-10-08 00:00:00', 7, 4);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (5, N'2026-10-12 00:00:00', 9, 5);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (6, N'2026-10-15 00:00:00', 2, 2);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (7, N'2026-10-18 00:00:00', 4, 3);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (8, N'2026-10-21 00:00:00', 6, 4);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (9, N'2026-10-25 00:00:00', 8, 5);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (10, N'2026-10-28 00:00:00', 10, 2);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (11, N'2026-10-29 00:00:00', 11, 3);
INSERT INTO Orders (Id_Order, Date_Order, Item_Id, UserId) VALUES (12, N'2026-10-30 00:00:00', 12, 4);

