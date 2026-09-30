-- Run this to clean duplicate employees, keeping only the lowest EmpID per person
USE KingsSchoolEgypt;
GO

-- Delete duplicates keeping the first (lowest EmpID) occurrence
WITH CTE AS (
    SELECT EmpID,
           ROW_NUMBER() OVER (
               PARTITION BY FirstName, FamilyName, NationalID
               ORDER BY EmpID ASC
           ) AS rn
    FROM Employees
)
DELETE FROM CTE WHERE rn > 1;
GO

PRINT 'Duplicates removed. Remaining employees:';
SELECT EmpID, FirstName+' '+FamilyName AS Name, JobTitle, Department
FROM Employees ORDER BY EmpID;
GO
