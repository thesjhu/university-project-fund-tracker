-- University Project Activity Tracker (UPAT)
-- Reporting queries: examples of the questions the schema answers.
-- Paste one statement at a time into Create > Query Design > SQL View.
-- Dialect: Microsoft Access SQL. Access runs one statement per query,
-- so run each statement on its own (or use the buttons on the automation form).

-- Q1. Fund balance by project
-- Total fund balance and number of funds for every project. The LEFT JOIN keeps projects that have no fund yet.
SELECT p.ProjectID, p.ProjectTitle,
       COUNT(f.FundCode)      AS FundCount,
       Nz(SUM(f.FundBalance), 0) AS TotalFundBalance
FROM PROJECT AS p
LEFT JOIN FUND AS f ON p.ProjectID = f.ProjectID
GROUP BY p.ProjectID, p.ProjectTitle
ORDER BY Nz(SUM(f.FundBalance), 0) DESC;

-- Q2. Projects with no fund (exception report)
-- The business rule says every project must have at least one fund. This lists the projects that break it.
SELECT p.ProjectID, p.ProjectTitle
FROM PROJECT AS p
LEFT JOIN FUND AS f ON p.ProjectID = f.ProjectID
WHERE f.FundCode IS NULL
ORDER BY p.ProjectID;

-- Q3. Funds and their sponsors
-- One row per fund-sponsor pair. A fund with two sponsors appears twice, so do not total FundBalance from this result.
SELECT f.FundCode, f.FundDescription, f.FundBalance, s.SponsorName
FROM (FUND AS f
INNER JOIN FUNDSPONSOR AS fs ON f.FundCode = fs.FundCode)
INNER JOIN SPONSOR AS s ON fs.SponsorID = s.SponsorID
ORDER BY f.FundCode, s.SponsorName;

-- Q4. Activity count by department and fiscal year
-- How many project activities each department ran in each fiscal year.
SELECT d.DepartmentName, a.FiscalYear, COUNT(*) AS ActivityCount
FROM ACTIVITY AS a
INNER JOIN DEPARTMENT AS d ON a.DepartmentID = d.DepartmentID
GROUP BY d.DepartmentName, a.FiscalYear
ORDER BY d.DepartmentName, a.FiscalYear;

-- Q5. Department summary
-- Faculty headcount and number of distinct projects each department takes part in. Subqueries avoid double counting across the two joins.
SELECT d.DepartmentID, d.DepartmentName,
       (SELECT COUNT(*) FROM FACULTY AS fa
         WHERE fa.DepartmentID = d.DepartmentID) AS FacultyCount,
       (SELECT COUNT(*) FROM (SELECT DISTINCT a.DepartmentID, a.ProjectID FROM ACTIVITY AS a) AS dp
         WHERE dp.DepartmentID = d.DepartmentID) AS ProjectCount
FROM DEPARTMENT AS d
ORDER BY d.DepartmentName;

-- Q6. Activities open on a given date
-- Access prompts for the date. Lists the activities whose start and end dates span it.
PARAMETERS [As of date] DATETIME;
SELECT a.ProjectID, p.ProjectTitle, a.ActivityID,
       a.ActivityStartDate, a.ActivityEndDate, d.DepartmentName
FROM (ACTIVITY AS a
INNER JOIN PROJECT AS p ON a.ProjectID = p.ProjectID)
INNER JOIN DEPARTMENT AS d ON a.DepartmentID = d.DepartmentID
WHERE a.ActivityStartDate <= [As of date]
  AND a.ActivityEndDate >= [As of date]
ORDER BY a.ProjectID, a.ActivityID;

