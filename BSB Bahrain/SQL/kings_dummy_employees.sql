USE KingsSchoolEgypt;
GO

INSERT INTO Employees (
    EmployeeID, FirstName, FamilyName, FullNameArabic,
    NationalID, DateOfBirth, Gender, Nationality, MaritalStatus,
    MobileNumber, PersonalEmail, Address,
    EmployeeNumber, DateOfJoining, EmploymentStatus, EmploymentType,
    JobTitle, Department, Branch, ReportingManager,
    ContractStartDate, ContractEndDate, ProbationPeriod, WorkingHours, WorkSchedule,
    SalaryType, BasicSalary, HousingAllowance, TransportAllowance,
    OtherFixedAllowances, Bonuses, OtherEarnings,
    PaymentMethod, Currency,
    IsTaxable, TaxStatus,
    CreatedAt
)
VALUES
(
    'EMP-001', 'Amr', 'Ziyad', 'عمرو زياد',
    '12345678901234', '1995-06-15', 'Male', 'Egyptian', 'Single',
    '01001234567', 'amr.ziyad@kings.edu.eg', 'Cairo, Egypt',
    'E001', '2023-09-01', 'Active', 'Full-Time',
    'IT Manager', 'IT', 'Cairo Campus', 'Ahmed Hassan',
    '2023-09-01', '2025-09-01', '3 Months', '8', 'Sunday to Thursday',
    'Monthly', 40000, 1000, 500, 500, 500, 0,
    'Bank Transfer', 'EGP',
    1, 'Active',
    GETDATE()
),
(
    'EMP-002', 'Ahmed', 'Hassan', 'أحمد حسن',
    '23456789012345', '1988-03-22', 'Male', 'Egyptian', 'Married',
    '01112345678', 'ahmed.hassan@kings.edu.eg', 'Giza, Egypt',
    'E002', '2022-01-15', 'Active', 'Full-Time',
    'HR Manager', 'HR', 'Cairo Campus', 'Director',
    '2022-01-15', '2024-01-15', '3 Months', '8', 'Sunday to Thursday',
    'Monthly', 55000, 2000, 500, 1000, 1000, 0,
    'Bank Transfer', 'EGP',
    1, 'Active',
    GETDATE()
),
(
    'EMP-003', 'Sarah', 'Smith', 'سارة سميث',
    '34567890123456', '1991-11-08', 'Female', 'British', 'Single',
    '01223456789', 'sarah.smith@kings.edu.eg', 'New Cairo, Egypt',
    'E003', '2023-03-01', 'Active', 'Full-Time',
    'English Teacher', 'Teaching', 'Cairo Campus', 'Ahmed Hassan',
    '2023-03-01', '2025-03-01', '3 Months', '8', 'Sunday to Thursday',
    'Monthly', 70000, 5000, 1000, 2000, 2000, 500,
    'Bank Transfer', 'EGP',
    1, 'Active',
    GETDATE()
);
GO

PRINT 'Dummy employees inserted successfully.';
SELECT EmpID, FirstName+' '+FamilyName AS Name, JobTitle, BasicSalary
FROM Employees ORDER BY EmpID;
GO
