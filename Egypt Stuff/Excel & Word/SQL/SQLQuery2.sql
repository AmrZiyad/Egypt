USE KingsSchoolEgypt;

-- Delete payroll runs first
ALTER TABLE PayrollRuns NOCHECK CONSTRAINT ALL;
DELETE FROM PayrollRuns;
ALTER TABLE PayrollRuns CHECK CONSTRAINT ALL;

-- Now delete duplicates keeping lowest EmpID
SELECT MIN(EmpID) AS KeepID
INTO #ToKeep
FROM Employees
GROUP BY FirstName, FamilyName, NationalID;

DELETE FROM Employees
WHERE EmpID NOT IN (SELECT KeepID FROM #ToKeep);

DROP TABLE #ToKeep;

PRINT 'Done.';
SELECT EmpID, FirstName+' '+FamilyName AS Name FROM Employees ORDER BY EmpID;