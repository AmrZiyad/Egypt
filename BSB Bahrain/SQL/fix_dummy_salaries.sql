USE KingsSchoolEgypt;
GO

-- Reset Variable Allowances and Overtime to 0 for all employees
-- (these should only be filled when HR manually sets them per employee)
UPDATE Employees SET
    VariableAllowances = 0,
    Overtime = 0,
    ReImbursement = 0,
    FlightAllowance = 0,
    ReAllocationAllowance = 0
WHERE VariableAllowances > 0 OR Overtime > 0;
GO

PRINT 'Fixed. Current salary overview:';
SELECT EmpID, FirstName+' '+FamilyName AS Name,
       BasicSalary, HousingAllowance, TransportAllowance,
       OtherFixedAllowances, Bonuses,
       BasicSalary+HousingAllowance+TransportAllowance+OtherFixedAllowances+Bonuses AS TotalSalary
FROM Employees ORDER BY EmpID;
GO
