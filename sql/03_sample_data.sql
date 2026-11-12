-- University Project Activity Tracker (UPAT)
-- Step 3 of 3: sample data. All names, IDs and amounts are fictional.
-- Insert order matters: parent tables first (ZIPCODE before DEPARTMENT, and so on).
-- Dialect: Microsoft Access SQL. Access runs one statement per query,
-- so run each statement on its own (or use the buttons on the automation form).

-- ZIPCODE
INSERT INTO ZIPCODE ( ZipCode, City, State )
    VALUES ( '10010', 'New York', 'NY' );
INSERT INTO ZIPCODE ( ZipCode, City, State )
    VALUES ( '10011', 'New York', 'NY' );
INSERT INTO ZIPCODE ( ZipCode, City, State )
    VALUES ( '10012', 'New York', 'NY' );
INSERT INTO ZIPCODE ( ZipCode, City, State )
    VALUES ( '10013', 'New York', 'NY' );

-- DEPARTMENT
INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode )
    VALUES ( 'D101', 'Epidemiology and Biostatistics', '555-123-4567', 'epi.biostats@university.example', '321 Public Health Lane', '10010' );
INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode )
    VALUES ( 'D102', 'Health Policy and Management', '555-234-5678', 'health.policy@university.example', '456 Policy Plaza', '10011' );
INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode )
    VALUES ( 'D103', 'Environmental and Occupational Health', '555-345-6789', 'env.occup.health@university.example', '789 Environmental Avenue', '10012' );
INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode )
    VALUES ( 'D104', 'Community and Global Health', '555-456-7890', 'comm.global.health@university.example', '101 Community Street', '10013' );

-- PROJECT
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1001', 'Regional Health Data Consortium', 'Award ref AWD-0001' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1002', 'HIV Care Expansion Program - Country Office', 'Award ref AWD-0002' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1003', 'Maternal and Child Health Technical Assistance - HQ', 'Award ref AWD-0003' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1004', 'Health Workforce Capacity Building', 'Award ref AWD-0004' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1005', 'Laboratory Systems Technical Assistance', 'Award ref AWD-0005' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1006', 'Population Health Impact Surveys', 'Award ref AWD-0006' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1007', 'Health System Strengthening Initiative', 'Award ref AWD-0007' );
INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription )
    VALUES ( 'PRJ-1008', 'Global Technical Assistance Program', 'Award ref AWD-0008' );

-- FUND
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '68000', 'CAPITAL EQUIPMENT', 118500, 'PRJ-1001' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '65000', 'SUBCONTRACTS', 1250000, 'PRJ-1002' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '61900', 'MINOR EQUIPMENT', 96400, 'PRJ-1003' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '59000', 'FRINGE BENEFIT CHARGE', 412300, 'PRJ-1004' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '61400', 'SUPPLIES AND MATERIALS', 264750, 'PRJ-1005' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '64500', 'CONSULTANTS', 380000, 'PRJ-1007' );
INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID )
    VALUES ( '50000', 'SALARY AND WAGES', 1845200, 'PRJ-1007' );

-- SPONSOR
INSERT INTO SPONSOR ( SponsorID, SponsorName )
    VALUES ( 'S01', 'Federal Health Agency Grant' );
INSERT INTO SPONSOR ( SponsorID, SponsorName )
    VALUES ( 'S02', 'Private Foundation Award' );
INSERT INTO SPONSOR ( SponsorID, SponsorName )
    VALUES ( 'S03', 'University Internal Research Fund' );
INSERT INTO SPONSOR ( SponsorID, SponsorName )
    VALUES ( 'S04', 'Partner Institution Subaward' );

-- FUNDSPONSOR
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '68000', 'S02' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '65000', 'S04' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '61900', 'S01' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '59000', 'S02' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '59000', 'S04' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '61400', 'S01' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '64500', 'S03' );
INSERT INTO FUNDSPONSOR ( FundCode, SponsorID )
    VALUES ( '50000', 'S01' );

-- ACTIVITY
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1001', '01', #7/30/2004#, #12/31/2010#, 2012, 'D101' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1001', '02', #8/1/2007#, #7/31/2008#, 2013, 'D101' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1001', '03', #9/28/2007#, #9/27/2011#, 2014, 'D104' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1001', '04', #12/1/2012#, #2/28/2018#, 2015, 'D104' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1002', '01', #7/30/2004#, #12/31/2010#, 2012, 'D102' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1002', '02', #8/1/2007#, #7/31/2008#, 2013, 'D103' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1002', '03', #9/28/2007#, #9/27/2011#, 2014, 'D104' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1003', '01', #7/30/2004#, #12/31/2010#, 2012, 'D101' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1003', '02', #8/1/2007#, #7/31/2008#, 2013, 'D102' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1003', '04', #12/1/2012#, #2/28/2018#, 2021, 'D102' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1004', '01', #7/30/2004#, #12/31/2010#, 2014, 'D101' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1004', '02', #8/1/2007#, #7/31/2008#, 2015, 'D104' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1004', '04', #12/1/2012#, #2/28/2018#, 2017, 'D104' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1005', '01', #7/30/2004#, #12/31/2010#, 2014, 'D101' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1005', '02', #8/1/2007#, #7/31/2008#, 2015, 'D102' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1006', '01', #7/30/2004#, #12/31/2010#, 2015, 'D102' );
INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID )
    VALUES ( 'PRJ-1006', '04', #12/1/2012#, #2/28/2018#, 2018, 'D102' );

-- FACULTY
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID001', 'Olivia', 'Anderson', 'Professor', '555-111-2233', 'olivia.anderson@university.example', 'D101' );
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID002', 'Marcus', 'Rodriguez', 'Associate Professor', '555-222-3344', 'marcus.rodriguez@university.example', 'D102' );
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID003', 'Sophia', 'Carter', 'Assistant Professor', '555-333-4455', 'sophia.carter@university.example', 'D103' );
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID004', 'Daniel', 'Walker', 'Professor', '555-444-5566', 'daniel.walker@university.example', 'D102' );
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID005', 'Emily', 'Foster', 'Associate Professor', '555-555-6677', 'emily.foster@university.example', 'D101' );
INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID )
    VALUES ( 'FID006', 'Robert', 'Turner', 'Assistant Professor', '555-666-7788', 'robert.turner@university.example', 'D104' );

