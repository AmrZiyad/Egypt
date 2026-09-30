USE KingsSchoolEgypt;
GO

-- Add missing fields to Employees table
IF COL_LENGTH('Employees','Commissions') IS NULL
    ALTER TABLE Employees ADD Commissions DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','TaxableSalary') IS NULL
    ALTER TABLE Employees ADD TaxableSalary DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','EmployeeSIDeduction') IS NULL
    ALTER TABLE Employees ADD EmployeeSIDeduction DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','PersonalAllowance') IS NULL
    ALTER TABLE Employees ADD PersonalAllowance DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','TaxableBenefits') IS NULL
    ALTER TABLE Employees ADD TaxableBenefits DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','MonthlyIncomeTax') IS NULL
    ALTER TABLE Employees ADD MonthlyIncomeTax DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','AnnualTaxAdjustment') IS NULL
    ALTER TABLE Employees ADD AnnualTaxAdjustment DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','PreviousPensionInfo') IS NULL
    ALTER TABLE Employees ADD PreviousPensionInfo NVARCHAR(500) NULL;
GO

PRINT 'New fields added successfully.';
GO
