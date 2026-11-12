' University Project Activity Tracker (UPAT)
' Code-behind for the "automation form" in the Access database.
' Access cannot run several SQL statements in one query, so two buttons do the work:
'   Command0 - runs every saved query whose name starts with "cre" (CREATE TABLE),
'              then every query whose name starts with "alt" (ALTER TABLE / foreign keys)
'   Command1 - loads the sample data in dependency order
' All sample names, IDs and amounts are fictional.

Option Compare Database
Option Explicit

Private Sub Command0_Click()
    MsgBox "Build all tables"

    ' suppress the confirmation prompt for each action query
    DoCmd.SetWarnings False

    Dim q As Object

    ' create the tables first
    For Each q In CurrentDb.QueryDefs
        If Left(q.Name, 3) = "cre" Then
            DoCmd.OpenQuery q.Name
        End If
    Next

    ' then add the foreign keys
    For Each q In CurrentDb.QueryDefs
        If Left(q.Name, 3) = "alt" Then
            DoCmd.OpenQuery q.Name
        End If
    Next

    DoCmd.SetWarnings True
End Sub

Private Sub Command1_Click()
    Dim db As DAO.Database
    Set db = CurrentDb

    ' ZIPCODE
    db.Execute "INSERT INTO ZIPCODE ( ZipCode, City, State ) " & _
               "VALUES ( '10010', 'New York', 'NY' );"
    db.Execute "INSERT INTO ZIPCODE ( ZipCode, City, State ) " & _
               "VALUES ( '10011', 'New York', 'NY' );"
    db.Execute "INSERT INTO ZIPCODE ( ZipCode, City, State ) " & _
               "VALUES ( '10012', 'New York', 'NY' );"
    db.Execute "INSERT INTO ZIPCODE ( ZipCode, City, State ) " & _
               "VALUES ( '10013', 'New York', 'NY' );"

    ' DEPARTMENT
    db.Execute "INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode ) " & _
               "VALUES ( 'D101', 'Epidemiology and Biostatistics', '555-123-4567', 'epi.biostats@university.example', '321 Public Health Lane', '10010' );"
    db.Execute "INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode ) " & _
               "VALUES ( 'D102', 'Health Policy and Management', '555-234-5678', 'health.policy@university.example', '456 Policy Plaza', '10011' );"
    db.Execute "INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode ) " & _
               "VALUES ( 'D103', 'Environmental and Occupational Health', '555-345-6789', 'env.occup.health@university.example', '789 Environmental Avenue', '10012' );"
    db.Execute "INSERT INTO DEPARTMENT ( DepartmentID, DepartmentName, Phone, Email, Street, ZipCode ) " & _
               "VALUES ( 'D104', 'Community and Global Health', '555-456-7890', 'comm.global.health@university.example', '101 Community Street', '10013' );"

    ' PROJECT
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1001', 'Regional Health Data Consortium', 'Award ref AWD-0001' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1002', 'HIV Care Expansion Program - Country Office', 'Award ref AWD-0002' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1003', 'Maternal and Child Health Technical Assistance - HQ', 'Award ref AWD-0003' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1004', 'Health Workforce Capacity Building', 'Award ref AWD-0004' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1005', 'Laboratory Systems Technical Assistance', 'Award ref AWD-0005' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1006', 'Population Health Impact Surveys', 'Award ref AWD-0006' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1007', 'Health System Strengthening Initiative', 'Award ref AWD-0007' );"
    db.Execute "INSERT INTO PROJECT ( ProjectID, ProjectTitle, ProjectDescription ) " & _
               "VALUES ( 'PRJ-1008', 'Global Technical Assistance Program', 'Award ref AWD-0008' );"

    ' FUND
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '68000', 'CAPITAL EQUIPMENT', 118500, 'PRJ-1001' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '65000', 'SUBCONTRACTS', 1250000, 'PRJ-1002' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '61900', 'MINOR EQUIPMENT', 96400, 'PRJ-1003' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '59000', 'FRINGE BENEFIT CHARGE', 412300, 'PRJ-1004' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '61400', 'SUPPLIES AND MATERIALS', 264750, 'PRJ-1005' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '64500', 'CONSULTANTS', 380000, 'PRJ-1007' );"
    db.Execute "INSERT INTO FUND ( FundCode, FundDescription, FundBalance, ProjectID ) " & _
               "VALUES ( '50000', 'SALARY AND WAGES', 1845200, 'PRJ-1007' );"

    ' SPONSOR
    db.Execute "INSERT INTO SPONSOR ( SponsorID, SponsorName ) " & _
               "VALUES ( 'S01', 'Federal Health Agency Grant' );"
    db.Execute "INSERT INTO SPONSOR ( SponsorID, SponsorName ) " & _
               "VALUES ( 'S02', 'Private Foundation Award' );"
    db.Execute "INSERT INTO SPONSOR ( SponsorID, SponsorName ) " & _
               "VALUES ( 'S03', 'University Internal Research Fund' );"
    db.Execute "INSERT INTO SPONSOR ( SponsorID, SponsorName ) " & _
               "VALUES ( 'S04', 'Partner Institution Subaward' );"

    ' FUNDSPONSOR
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '68000', 'S02' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '65000', 'S04' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '61900', 'S01' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '59000', 'S02' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '59000', 'S04' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '61400', 'S01' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '64500', 'S03' );"
    db.Execute "INSERT INTO FUNDSPONSOR ( FundCode, SponsorID ) " & _
               "VALUES ( '50000', 'S01' );"

    ' ACTIVITY
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1001', '01', #7/30/2004#, #12/31/2010#, 2012, 'D101' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1001', '02', #8/1/2007#, #7/31/2008#, 2013, 'D101' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1001', '03', #9/28/2007#, #9/27/2011#, 2014, 'D104' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1001', '04', #12/1/2012#, #2/28/2018#, 2015, 'D104' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1002', '01', #7/30/2004#, #12/31/2010#, 2012, 'D102' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1002', '02', #8/1/2007#, #7/31/2008#, 2013, 'D103' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1002', '03', #9/28/2007#, #9/27/2011#, 2014, 'D104' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1003', '01', #7/30/2004#, #12/31/2010#, 2012, 'D101' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1003', '02', #8/1/2007#, #7/31/2008#, 2013, 'D102' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1003', '04', #12/1/2012#, #2/28/2018#, 2021, 'D102' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1004', '01', #7/30/2004#, #12/31/2010#, 2014, 'D101' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1004', '02', #8/1/2007#, #7/31/2008#, 2015, 'D104' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1004', '04', #12/1/2012#, #2/28/2018#, 2017, 'D104' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1005', '01', #7/30/2004#, #12/31/2010#, 2014, 'D101' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1005', '02', #8/1/2007#, #7/31/2008#, 2015, 'D102' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1006', '01', #7/30/2004#, #12/31/2010#, 2015, 'D102' );"
    db.Execute "INSERT INTO ACTIVITY ( ProjectID, ActivityID, ActivityStartDate, ActivityEndDate, FiscalYear, DepartmentID ) " & _
               "VALUES ( 'PRJ-1006', '04', #12/1/2012#, #2/28/2018#, 2018, 'D102' );"

    ' FACULTY
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID001', 'Olivia', 'Anderson', 'Professor', '555-111-2233', 'olivia.anderson@university.example', 'D101' );"
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID002', 'Marcus', 'Rodriguez', 'Associate Professor', '555-222-3344', 'marcus.rodriguez@university.example', 'D102' );"
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID003', 'Sophia', 'Carter', 'Assistant Professor', '555-333-4455', 'sophia.carter@university.example', 'D103' );"
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID004', 'Daniel', 'Walker', 'Professor', '555-444-5566', 'daniel.walker@university.example', 'D102' );"
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID005', 'Emily', 'Foster', 'Associate Professor', '555-555-6677', 'emily.foster@university.example', 'D101' );"
    db.Execute "INSERT INTO FACULTY ( FacultyID, FirstName, LastName, Title, Phone, Email, DepartmentID ) " & _
               "VALUES ( 'FID006', 'Robert', 'Turner', 'Assistant Professor', '555-666-7788', 'robert.turner@university.example', 'D104' );"

    MsgBox "Sample data loaded"
End Sub
