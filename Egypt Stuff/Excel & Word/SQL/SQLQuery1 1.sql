USE KingsSchoolEgypt;

-- Clear everything
DELETE FROM PayrollRuns;
DELETE FROM ITAssignments;
DELETE FROM ITAssets;
DELETE FROM OpsAssets;
DELETE FROM Employees;

-- Reset ID counters back to 1
DBCC CHECKIDENT ('Employees', RESEED, 0);
DBCC CHECKIDENT ('PayrollRuns', RESEED, 0);
DBCC CHECKIDENT ('ITAssets', RESEED, 0);

PRINT 'All done. Database is clean and ready.';