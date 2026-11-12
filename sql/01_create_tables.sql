-- University Project Activity Tracker (UPAT)
-- Step 1 of 3: create the eight tables with primary keys
-- Dialect: Microsoft Access SQL. Access runs one statement per query,
-- so run each statement on its own (or use the buttons on the automation form).

CREATE TABLE ZIPCODE
(
    ZipCode VARCHAR(10) NOT NULL,
    City    VARCHAR(25),
    State   VARCHAR(2),
    CONSTRAINT PK_ZIPCODE PRIMARY KEY (ZipCode)
);

CREATE TABLE DEPARTMENT
(
    DepartmentID   VARCHAR(10) NOT NULL,
    DepartmentName VARCHAR(50) NOT NULL,
    Phone          VARCHAR(12),
    Email          VARCHAR(50),
    Street         VARCHAR(50),
    ZipCode        VARCHAR(10) NOT NULL,
    CONSTRAINT PK_DEPARTMENT PRIMARY KEY (DepartmentID)
);

CREATE TABLE FACULTY
(
    FacultyID    VARCHAR(10) NOT NULL,
    FirstName    VARCHAR(25) NOT NULL,
    LastName     VARCHAR(25) NOT NULL,
    Title        VARCHAR(25) NOT NULL,
    Phone        VARCHAR(12),
    Email        VARCHAR(50),
    DepartmentID VARCHAR(10) NOT NULL,
    CONSTRAINT PK_FACULTY PRIMARY KEY (FacultyID)
);

CREATE TABLE SPONSOR
(
    SponsorID   VARCHAR(10) NOT NULL,
    SponsorName VARCHAR(50) NOT NULL,
    CONSTRAINT PK_SPONSOR PRIMARY KEY (SponsorID)
);

CREATE TABLE PROJECT
(
    ProjectID          VARCHAR(15)  NOT NULL,
    ProjectTitle       VARCHAR(255) NOT NULL,
    ProjectDescription VARCHAR(50)  NOT NULL,
    CONSTRAINT PK_PROJECT PRIMARY KEY (ProjectID)
);

CREATE TABLE FUND
(
    FundCode        VARCHAR(10) NOT NULL,
    FundDescription VARCHAR(50) NOT NULL,
    FundBalance     CURRENCY    NOT NULL,
    ProjectID       VARCHAR(15) NOT NULL,
    CONSTRAINT PK_FUND PRIMARY KEY (FundCode)
);

CREATE TABLE FUNDSPONSOR
(
    FundCode  VARCHAR(10) NOT NULL,
    SponsorID VARCHAR(10) NOT NULL,
    CONSTRAINT PK_FUNDSPONSOR PRIMARY KEY (FundCode, SponsorID)
);

CREATE TABLE ACTIVITY
(
    ProjectID         VARCHAR(15) NOT NULL,
    ActivityID        VARCHAR(4)  NOT NULL,
    ActivityStartDate DATETIME    NOT NULL,
    ActivityEndDate   DATETIME    NOT NULL,
    FiscalYear        SHORT       NOT NULL,
    DepartmentID      VARCHAR(10) NOT NULL,
    CONSTRAINT PK_ACTIVITY PRIMARY KEY (ProjectID, ActivityID)
);
