-- Kings School The Crown Egypt — Database Setup
-- Run this in SQL Server Management Studio
-- Database: KingsSchoolEgypt

USE master;
GO
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'KingsSchoolEgypt')
    CREATE DATABASE KingsSchoolEgypt;
GO
USE KingsSchoolEgypt;
GO

-- ── Users & Auth ──────────────────────────────────────────
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='Users')
CREATE TABLE Users (
    UserID        INT IDENTITY PRIMARY KEY,
    Username      NVARCHAR(200) NOT NULL UNIQUE,
    PasswordHash  NVARCHAR(200) NOT NULL,
    Role          NVARCHAR(20)  NOT NULL DEFAULT 'staff'
                  CHECK (Role IN ('admin','hr','it','operation')),
    FullName      NVARCHAR(200) NULL,
    Email         NVARCHAR(200) NULL,
    Department    NVARCHAR(100) NULL,
    EmployeeNumber NVARCHAR(50) NULL,
    MustChangePassword BIT NOT NULL DEFAULT 1,
    LastLogin     DATETIME NULL,
    CreatedAt     DATETIME NOT NULL DEFAULT GETDATE()
);
GO

-- Seed default accounts
IF NOT EXISTS (SELECT 1 FROM Users WHERE Username='admin')
    INSERT INTO Users(Username,PasswordHash,Role,FullName,MustChangePassword)
    VALUES('admin','admin123','admin','Administrator',0);
IF NOT EXISTS (SELECT 1 FROM Users WHERE Username='hr')
    INSERT INTO Users(Username,PasswordHash,Role,FullName,MustChangePassword)
    VALUES('hr','hr123','hr','HR Manager',0);
IF NOT EXISTS (SELECT 1 FROM Users WHERE Username='it')
    INSERT INTO Users(Username,PasswordHash,Role,FullName,MustChangePassword)
    VALUES('it','it123','it','IT Manager',0);
IF NOT EXISTS (SELECT 1 FROM Users WHERE Username='operation')
    INSERT INTO Users(Username,PasswordHash,Role,FullName,MustChangePassword)
    VALUES('operation','op123','operation','Operations Manager',0);
GO

-- ── Employees ─────────────────────────────────────────────
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='Employees')
CREATE TABLE Employees (
    EmpID               INT IDENTITY PRIMARY KEY,
    -- Personal
    EmployeeID          NVARCHAR(50)  NULL,
    FirstName           NVARCHAR(100) NULL,
    SecondName          NVARCHAR(100) NULL,
    ThirdName           NVARCHAR(100) NULL,
    FamilyName          NVARCHAR(100) NULL,
    FullNameArabic      NVARCHAR(200) NULL,
    NationalID          NVARCHAR(50)  NULL,
    DateOfBirth         DATE          NULL,
    Gender              NVARCHAR(20)  NULL,
    Nationality         NVARCHAR(100) NULL,
    MaritalStatus       NVARCHAR(30)  NULL,
    MobileNumber        NVARCHAR(30)  NULL,
    PersonalEmail       NVARCHAR(200) NULL,
    Address             NVARCHAR(500) NULL,
    -- Employment
    EmployeeNumber      NVARCHAR(50)  NULL,
    DateOfJoining       DATE          NULL,
    EmploymentStatus    NVARCHAR(50)  NULL,
    EmploymentType      NVARCHAR(50)  NULL,
    JobTitle            NVARCHAR(100) NULL,
    Department          NVARCHAR(100) NULL,
    Branch              NVARCHAR(100) NULL,
    ReportingManager    NVARCHAR(100) NULL,
    ContractStartDate   DATE          NULL,
    ContractEndDate     DATE          NULL,
    ProbationPeriod     NVARCHAR(50)  NULL,
    ConfirmationDate    DATE          NULL,
    WorkingHours        NVARCHAR(50)  NULL,
    WorkSchedule        NVARCHAR(100) NULL,
    -- Salary
    SalaryType          NVARCHAR(30)  NULL,
    BasicSalary         DECIMAL(12,2) NULL DEFAULT 0,
    GrossSalary         DECIMAL(12,2) NULL DEFAULT 0,
    BasicInsuranceSalary DECIMAL(12,2) NULL DEFAULT 0,
    HousingAllowance    DECIMAL(12,2) NULL DEFAULT 0,
    TransportAllowance  DECIMAL(12,2) NULL DEFAULT 0,
    OtherFixedAllowances DECIMAL(12,2) NULL DEFAULT 0,
    VariableAllowances  DECIMAL(12,2) NULL DEFAULT 0,
    Overtime            DECIMAL(12,2) NULL DEFAULT 0,
    Bonuses             DECIMAL(12,2) NULL DEFAULT 0,
    Commissions         DECIMAL(12,2) NULL DEFAULT 0,
    OtherEarnings       DECIMAL(12,2) NULL DEFAULT 0,
    RecurringDeductions DECIMAL(12,2) NULL DEFAULT 0,
    EffectiveFromDate   DATE          NULL,
    PaymentMethod       NVARCHAR(50)  NULL,
    Currency            NVARCHAR(10)  NULL DEFAULT 'EGP',
    -- Bank
    BankName            NVARCHAR(100) NULL,
    BankAccountNumber   NVARCHAR(50)  NULL,
    IBAN                NVARCHAR(50)  NULL,
    -- Tax
    IsTaxable           BIT           NOT NULL DEFAULT 1,
    TaxRegistrationID   NVARCHAR(50)  NULL,
    TaxTreatment        NVARCHAR(100) NULL,
    TaxStatus           NVARCHAR(50)  NULL,
    OtherTaxExemptions  DECIMAL(12,2) NULL DEFAULT 0,
    -- Social Insurance
    InsuranceNumber     NVARCHAR(50)  NULL,
    InsuranceOffice     NVARCHAR(100) NULL,
    InsuranceStartDate  DATE          NULL,
    PreviousInsurancePeriod NVARCHAR(100) NULL,
    PreviousEmployer    NVARCHAR(200) NULL,
    InsuranceWage       DECIMAL(12,2) NULL DEFAULT 0,
    InsuranceStatus     NVARCHAR(50)  NULL,
    InsuranceCategory   NVARCHAR(100) NULL,
    MedicalFitnessCert  NVARCHAR(200) NULL,
    -- Documents
    DocNationalID       NVARCHAR(200) NULL,
    DocBirthCert        NVARCHAR(200) NULL,
    DocContract         NVARCHAR(200) NULL,
    DocEducationCert    NVARCHAR(200) NULL,
    DocSyndicate        NVARCHAR(200) NULL,
    DocMedicalForm111   NVARCHAR(200) NULL,
    DocSocialInsurance  NVARCHAR(200) NULL,
    DocBank             NVARCHAR(200) NULL,
    DocPassport         NVARCHAR(200) NULL,
    DocWorkPermit       NVARCHAR(200) NULL,
    DocOther            NVARCHAR(200) NULL,
    -- Meta
    CreatedAt           DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt           DATETIME NULL
);
GO

-- ── IT Assets (same as BSB) ───────────────────────────────
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='ITCategories')
CREATE TABLE ITCategories (
    ITCatID INT IDENTITY PRIMARY KEY,
    Name    NVARCHAR(100) NOT NULL
);
GO
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='ITAssets')
CREATE TABLE ITAssets (
    AssetID     INT IDENTITY PRIMARY KEY,
    Name        NVARCHAR(200) NOT NULL,
    Model       NVARCHAR(200) NULL,
    Serial      NVARCHAR(200) NULL,
    AssetTag    NVARCHAR(100) NULL,
    Location    NVARCHAR(200) NULL,
    ITCatID     INT NULL REFERENCES ITCategories(ITCatID),
    Stock       INT NOT NULL DEFAULT 1,
    Unit        NVARCHAR(30) NULL DEFAULT 'pcs',
    DetailsJSON NVARCHAR(MAX) NULL,
    CreatedAt   DATETIME NOT NULL DEFAULT GETDATE()
);
GO
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='ITAssignments')
CREATE TABLE ITAssignments (
    AssignID    INT IDENTITY PRIMARY KEY,
    AssetID     INT NULL REFERENCES ITAssets(AssetID),
    AssetLabel  NVARCHAR(200) NULL,
    AssignedTo  NVARCHAR(200) NULL,
    AssignedBy  NVARCHAR(200) NULL,
    AssignDate  DATETIME NOT NULL DEFAULT GETDATE(),
    DueDate     NVARCHAR(50) NULL,
    Status      NVARCHAR(30) NULL DEFAULT 'Assigned',
    Notes       NVARCHAR(1000) NULL
);
GO

-- ── Operations (same as BSB) ──────────────────────────────
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='OpsAssets')
CREATE TABLE OpsAssets (
    OpsID           INT IDENTITY PRIMARY KEY,
    AssetNumber     NVARCHAR(100) NULL,
    AssetCode       NVARCHAR(100) NULL,
    AssetDescription NVARCHAR(300) NULL,
    Category        NVARCHAR(100) NULL,
    SubCategory     NVARCHAR(100) NULL,
    Location        NVARCHAR(200) NULL,
    Condition       NVARCHAR(50)  NULL,
    SerialCode      NVARCHAR(100) NULL,
    Brand           NVARCHAR(100) NULL,
    Model           NVARCHAR(100) NULL,
    CreatedAt       DATETIME NOT NULL DEFAULT GETDATE()
);
GO

PRINT 'Kings School Egypt DB setup complete.';
SELECT 'Users' AS [Table], COUNT(*) AS Rows FROM Users
UNION ALL SELECT 'Employees', COUNT(*) FROM Employees
UNION ALL SELECT 'ITCategories', COUNT(*) FROM ITCategories
UNION ALL SELECT 'ITAssets', COUNT(*) FROM ITAssets
UNION ALL SELECT 'OpsAssets', COUNT(*) FROM OpsAssets;
GO
