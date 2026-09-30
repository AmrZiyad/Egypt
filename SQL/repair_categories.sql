-- BSB Resource Manager — Repair: fix categories not showing in sidebar
-- Safe to run multiple times. Re-links any subcategory whose parent is
-- missing/NULL to a valid parent so it appears in the sidebar.
USE BSBStationery;
GO

-- 1) Make sure at least one parent category exists.
IF NOT EXISTS (SELECT 1 FROM ParentCategories)
    INSERT INTO ParentCategories(Name) VALUES ('STATIONERY');
GO

-- 2) Pick the STATIONERY parent (or the first parent) as the home for orphans.
DECLARE @pid INT;
SELECT TOP 1 @pid = ParentCatID FROM ParentCategories WHERE Name = 'STATIONERY';
IF @pid IS NULL
    SELECT TOP 1 @pid = ParentCatID FROM ParentCategories ORDER BY ParentCatID;

-- 3) Re-link any subcategory that has NULL parent OR points to a non-existent parent.
UPDATE SubCategories
SET ParentCatID = @pid
WHERE ParentCatID IS NULL
   OR ParentCatID NOT IN (SELECT ParentCatID FROM ParentCategories);
GO

-- 4) Report what we have now.
SELECT p.Name AS Parent, COUNT(s.SubCatID) AS SubCategories
FROM ParentCategories p
LEFT JOIN SubCategories s ON s.ParentCatID = p.ParentCatID
GROUP BY p.Name
ORDER BY p.Name;

SELECT s.Name AS SubCategory,
       (SELECT COUNT(*) FROM Items i WHERE i.SubCatID = s.SubCatID) AS Items
FROM SubCategories s
ORDER BY s.Name;
GO

PRINT 'Repair complete — all subcategories now linked to a parent.';
GO
