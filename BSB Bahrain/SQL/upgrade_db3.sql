USE BSBStationery;
GO

-- Add email to Orders table so we can email staff when arranged
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name=N'OrderedByEmail' AND Object_ID=Object_ID(N'Orders'))
    ALTER TABLE Orders ADD OrderedByEmail NVARCHAR(150) NULL;
GO

-- Update Orders to store email from user at time of ordering
-- (will be populated by app going forward)

-- Seed the 5 department accounts for testing
MERGE Users AS tgt
USING (VALUES
    ('admin',  'admin123',  'admin',       'Administrator',      'admin@thebsbh.com',       0),
    ('proc',   'proc123',   'procurement', 'Procurement Head',   'proc@thebsbh.com',        0),
    ('it',     'it123',     'it',          'IT Manager',         'it@thebsbh.com',          0),
    ('op',     'op123',     'operation',   'Operations Manager', 'op@thebsbh.com',          0),
    ('staff',  'staff123',  'staff',       'Staff Member',       'staff@thebsbh.com',       0)
) AS src(Username,PasswordHash,Role,FullName,Email,MustChangePassword)
ON tgt.Username=src.Username
WHEN MATCHED THEN
    UPDATE SET Role=src.Role,FullName=src.FullName,Email=src.Email
WHEN NOT MATCHED THEN
    INSERT(Username,PasswordHash,Role,FullName,Email,MustChangePassword)
    VALUES(src.Username,src.PasswordHash,src.Role,src.FullName,src.Email,src.MustChangePassword);
GO

PRINT 'Upgrade 3 complete.';
GO
