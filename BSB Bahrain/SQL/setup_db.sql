-- BSB Stationery Inventory System - Database Setup
-- Run this in SSMS before launching the app

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'BSBStationery')
BEGIN
    CREATE DATABASE BSBStationery;
END
GO

USE BSBStationery;
GO

-- Users table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Users' AND xtype='U')
CREATE TABLE Users (
    UserID INT IDENTITY(1,1) PRIMARY KEY,
    Username NVARCHAR(50) NOT NULL UNIQUE,
    PasswordHash NVARCHAR(255) NOT NULL,
    Role NVARCHAR(10) NOT NULL CHECK (Role IN ('admin', 'staff')),
    FullName NVARCHAR(100),
    CreatedAt DATETIME DEFAULT GETDATE()
);
GO

-- Parent Categories
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ParentCategories' AND xtype='U')
CREATE TABLE ParentCategories (
    ParentCatID INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL UNIQUE
);
GO

-- Sub Categories
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='SubCategories' AND xtype='U')
CREATE TABLE SubCategories (
    SubCatID INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(100) NOT NULL,
    ParentCatID INT FOREIGN KEY REFERENCES ParentCategories(ParentCatID) ON DELETE CASCADE
);
GO

-- Items
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Items' AND xtype='U')
CREATE TABLE Items (
    ItemID INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(150) NOT NULL,
    SubCatID INT FOREIGN KEY REFERENCES SubCategories(SubCatID) ON DELETE SET NULL,
    Quantity INT DEFAULT 0,
    Unit NVARCHAR(30) DEFAULT 'pcs',
    UnitPrice DECIMAL(10,3) DEFAULT 0,
    MinStock INT DEFAULT 5,
    CreatedAt DATETIME DEFAULT GETDATE()
);
GO

-- Issue Records
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IssueRecords' AND xtype='U')
CREATE TABLE IssueRecords (
    IssueID INT IDENTITY(1,1) PRIMARY KEY,
    IssuedBy NVARCHAR(100),
    RecipientName NVARCHAR(100),
    Department NVARCHAR(100),
    Notes NVARCHAR(500),
    IssuedAt DATETIME DEFAULT GETDATE()
);
GO

-- Issue Items (line items per issue)
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='IssueItems' AND xtype='U')
CREATE TABLE IssueItems (
    IssueItemID INT IDENTITY(1,1) PRIMARY KEY,
    IssueID INT FOREIGN KEY REFERENCES IssueRecords(IssueID) ON DELETE CASCADE,
    ItemID INT FOREIGN KEY REFERENCES Items(ItemID) ON DELETE SET NULL,
    ItemName NVARCHAR(150),
    Quantity INT,
    UnitPrice DECIMAL(10,3)
);
GO

-- Seed default users (password: admin123 and staff123)
IF NOT EXISTS (SELECT * FROM Users WHERE Username = 'admin')
INSERT INTO Users (Username, PasswordHash, Role, FullName)
VALUES ('admin', 'admin123', 'admin', 'Administrator');

IF NOT EXISTS (SELECT * FROM Users WHERE Username = 'staff')
INSERT INTO Users (Username, PasswordHash, Role, FullName)
VALUES ('staff', 'staff123', 'staff', 'Staff Member');
GO

-- Seed parent categories
IF NOT EXISTS (SELECT * FROM ParentCategories WHERE Name = 'STATIONERY')
INSERT INTO ParentCategories (Name) VALUES 
('STATIONERY'), ('PANTRY'), ('FACILITIES');
GO

-- Seed sub categories
IF NOT EXISTS (SELECT * FROM SubCategories WHERE Name = 'Pens & Pencils')
BEGIN
    INSERT INTO SubCategories (Name, ParentCatID) VALUES
    ('Pens & Pencils', 1), ('Papers', 1), ('Files & Folders', 1),
    ('Markers', 1), ('Correction', 1), ('Clips & Pins', 1),
    ('Hot Drinks', 2), ('Cold Drinks', 2), ('Snacks', 2),
    ('Cleaning', 3), ('Equipment', 3);
END
GO

-- Seed sample items
IF NOT EXISTS (SELECT * FROM Items WHERE Name = 'Ballpoint Pen Blue')
BEGIN
    INSERT INTO Items (Name, SubCatID, Quantity, Unit, UnitPrice, MinStock) VALUES
    ('Ballpoint Pen Blue', 1, 200, 'pcs', 0.050, 20),
    ('Ballpoint Pen Red', 1, 100, 'pcs', 0.050, 10),
    ('HB Pencil', 1, 150, 'pcs', 0.030, 15),
    ('A4 Paper 80gsm', 2, 50, 'ream', 1.800, 10),
    ('A3 Paper', 2, 20, 'ream', 2.500, 5),
    ('Ring Binder A4', 3, 30, 'pcs', 0.800, 5),
    ('Whiteboard Marker Black', 4, 40, 'pcs', 0.250, 10),
    ('Highlighter Yellow', 4, 60, 'pcs', 0.200, 10),
    ('Correction Fluid', 5, 25, 'pcs', 0.350, 5),
    ('Staples Box', 6, 20, 'box', 0.400, 5),
    ('Nescafe Jar', 7, 10, 'jar', 2.500, 3),
    ('Tea Bags Box', 7, 15, 'box', 1.200, 3),
    ('Sugar 1kg', 7, 8, 'bag', 0.500, 3),
    ('Water 1.5L', 8, 48, 'bottle', 0.150, 12),
    ('Tissue Box', 10, 20, 'box', 0.600, 5);
END
GO

PRINT 'BSBStationery database setup complete!';
GO
