-- ════════════════════════════════════════════════════════════
-- BSB Stationery System — Database Upgrade
-- Adds: emails on accounts, staff/admin accounts, Orders, Messages (mail),
--       and Issue-Log completion tracking.
-- Safe to run multiple times.
-- ════════════════════════════════════════════════════════════
USE BSBStationery;
GO

-- 1) Add Email column to Users (if missing)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'Email' AND Object_ID = Object_ID(N'Users'))
    ALTER TABLE Users ADD Email NVARCHAR(150) NULL;
GO

-- 2) Seed / update accounts (admins + staff) with emails.
--    Passwords are stored exactly like the existing ones (plain text in PasswordHash).
MERGE Users AS tgt
USING (VALUES
    -- Admins (procurement / office)
    ('admin',        'admin123',  'admin', 'Administrator',        'admin@bsbahrain.com'),
    ('procurement',  'proc123',   'admin', 'Procurement Head',     'procurement@bsbahrain.com'),
    ('office',       'office123', 'admin', 'Office Manager',       'office@bsbahrain.com'),
    -- Staff (teachers)
    ('staff',        'staff123',  'staff', 'Staff Member',         'staff@bsbahrain.com'),
    ('jsmith',       'pass123',   'staff', 'J. Smith',             'j.smith@bsbahrain.com'),
    ('mkhan',        'pass123',   'staff', 'M. Khan',              'm.khan@bsbahrain.com'),
    ('rwilson',      'pass123',   'staff', 'R. Wilson',            'r.wilson@bsbahrain.com')
) AS src (Username, PasswordHash, Role, FullName, Email)
ON tgt.Username = src.Username
WHEN MATCHED THEN
    UPDATE SET Role = src.Role, FullName = src.FullName, Email = src.Email
WHEN NOT MATCHED THEN
    INSERT (Username, PasswordHash, Role, FullName, Email)
    VALUES (src.Username, src.PasswordHash, src.Role, src.FullName, src.Email);
GO

-- 3) Orders (staff order -> admin arranges -> collected)
IF OBJECT_ID('Orders','U') IS NULL
BEGIN
    CREATE TABLE Orders (
        OrderID        INT IDENTITY(1,1) PRIMARY KEY,
        OrderedBy      NVARCHAR(100) NOT NULL,   -- username
        OrderedByName  NVARCHAR(150) NOT NULL,   -- full name
        Department     NVARCHAR(150) NULL,
        Status         NVARCHAR(20)  NOT NULL DEFAULT 'Pending', -- Pending / Arranged / Collected
        Notes          NVARCHAR(400) NULL,
        CreatedAt      DATETIME      NOT NULL DEFAULT GETDATE()
    );
END
GO

IF OBJECT_ID('OrderItems','U') IS NULL
BEGIN
    CREATE TABLE OrderItems (
        OrderItemID INT IDENTITY(1,1) PRIMARY KEY,
        OrderID     INT NOT NULL,
        ItemID      INT NULL,
        ItemName    NVARCHAR(200) NOT NULL,
        Quantity    INT NOT NULL,
        Unit        NVARCHAR(60) NULL,
        CONSTRAINT FK_OrderItems_Orders FOREIGN KEY (OrderID)
            REFERENCES Orders(OrderID) ON DELETE CASCADE
    );
END
GO

-- 4) Messages (staff -> admin mail)
IF OBJECT_ID('Messages','U') IS NULL
BEGIN
    CREATE TABLE Messages (
        MessageID  INT IDENTITY(1,1) PRIMARY KEY,
        FromUser   NVARCHAR(100) NOT NULL,  -- username
        FromName   NVARCHAR(150) NOT NULL,
        ToUser     NVARCHAR(100) NOT NULL,  -- admin username (or 'all')
        Subject    NVARCHAR(200) NULL,
        Body       NVARCHAR(MAX) NULL,
        IsRead     BIT NOT NULL DEFAULT 0,
        SentAt     DATETIME NOT NULL DEFAULT GETDATE()
    );
END
GO

-- 5) Issue-Log completion: add a Status column to IssueRecords for tick-off
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'Completed' AND Object_ID = Object_ID(N'IssueRecords'))
    ALTER TABLE IssueRecords ADD Completed BIT NOT NULL DEFAULT 0;
GO

PRINT 'Upgrade complete. Accounts, Orders, OrderItems, Messages, and Issue completion are ready.';
GO
