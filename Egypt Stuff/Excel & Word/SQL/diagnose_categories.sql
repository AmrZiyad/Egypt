-- Run this in SSMS to see WHY categories aren't showing.
USE BSBStationery;
GO

PRINT '=== Parent categories ===';
SELECT ParentCatID, Name FROM ParentCategories ORDER BY Name;

PRINT '=== Sub categories + their parent link ===';
SELECT s.SubCatID, s.Name AS SubCategory, s.ParentCatID,
       p.Name AS ParentName,
       (SELECT COUNT(*) FROM Items i WHERE i.SubCatID = s.SubCatID) AS Items
FROM SubCategories s
LEFT JOIN ParentCategories p ON s.ParentCatID = p.ParentCatID
ORDER BY s.Name;

PRINT '=== ORPHANS (these are why the sidebar hides them) ===';
SELECT s.SubCatID, s.Name, s.ParentCatID
FROM SubCategories s
WHERE s.ParentCatID IS NULL
   OR s.ParentCatID NOT IN (SELECT ParentCatID FROM ParentCategories);
GO
