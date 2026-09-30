USE KingsSchoolEgypt;
GO
IF COL_LENGTH('Employees','ReImbursement') IS NULL
    ALTER TABLE Employees ADD ReImbursement DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','FlightAllowance') IS NULL
    ALTER TABLE Employees ADD FlightAllowance DECIMAL(12,2) NULL DEFAULT 0;
IF COL_LENGTH('Employees','ReAllocationAllowance') IS NULL
    ALTER TABLE Employees ADD ReAllocationAllowance DECIMAL(12,2) NULL DEFAULT 0;
GO
PRINT 'New salary fields added.';
GO
