# University Project Activity Tracker (UPAT)

A relational database that lets a university track sponsored projects, the activities under each project, the funds that pay for them, the sponsors behind those funds, and the departments and faculty doing the work. It replaces the spreadsheets a finance or grants office typically uses for this.

Built in Microsoft Access with SQL and VBA as the term project for a graduate database course at Baruch College (Zicklin School of Business), November 2023. The project covers the full design cycle: business requirements, E-R model, relational model, normalization, SQL schema, and a working Access application.

## What it answers

- How much money is on each project, and across how many funds?
- Which projects have no fund attached?
- Which sponsors stand behind each fund?
- Which departments ran which activities, in which fiscal year?
- What is open as of a given date?

Example, from [`sql/04_reporting_queries.sql`](sql/04_reporting_queries.sql):

```sql
SELECT p.ProjectID, p.ProjectTitle,
       COUNT(f.FundCode)         AS FundCount,
       Nz(SUM(f.FundBalance), 0) AS TotalFundBalance
FROM PROJECT AS p
LEFT JOIN FUND AS f ON p.ProjectID = f.ProjectID
GROUP BY p.ProjectID, p.ProjectTitle
ORDER BY Nz(SUM(f.FundBalance), 0) DESC;
```

| ProjectID | ProjectTitle | FundCount | TotalFundBalance |
| :--- | :--- | :--- | ---: |
| PRJ-1007 | Health System Strengthening Initiative | 2 | $2,225,200 |
| PRJ-1002 | HIV Care Expansion Program - Country Office | 1 | $1,250,000 |
| PRJ-1004 | Health Workforce Capacity Building | 1 | $412,300 |
| PRJ-1005 | Laboratory Systems Technical Assistance | 1 | $264,750 |
| PRJ-1001 | Regional Health Data Consortium | 1 | $118,500 |
| PRJ-1003 | Maternal and Child Health Technical Assistance - HQ | 1 | $96,400 |
| PRJ-1006 | Population Health Impact Surveys | 0 | $0 |
| PRJ-1008 | Global Technical Assistance Program | 0 | $0 |

The two projects with a zero balance are the ones the exception query (Q2 in the same file) flags as having no fund.

## Data model

![E-R diagram](screenshot1.png)

| Table | Primary key | Foreign keys | Purpose |
| :--- | :--- | :--- | :--- |
| `PROJECT` | ProjectID | | A sponsored project |
| `ACTIVITY` | ProjectID + ActivityID | ProjectID, DepartmentID | A dated phase of a project, carried out by one department in one fiscal year |
| `FUND` | FundCode | ProjectID | A budget line and its balance, tied to one project |
| `SPONSOR` | SponsorID | | A funding source |
| `FUNDSPONSOR` | FundCode + SponsorID | FundCode, SponsorID | Junction table for the many-to-many link between funds and sponsors |
| `DEPARTMENT` | DepartmentID | ZipCode | An academic department |
| `FACULTY` | FacultyID | DepartmentID | A faculty member, employed by one department |
| `ZIPCODE` | ZipCode | | City and state for each ZIP code |

Design decisions:

- **Normalized to BCNF.** `DEPARTMENT` originally held City and State, which depend on ZipCode and not on the department. Splitting out `ZIPCODE` removed that transitive dependency. The report walks through the functional dependencies for every table.
- **Composite key on `ACTIVITY`.** Activity numbers restart at 01 for each project, so the key is ProjectID + ActivityID.
- **Junction table for sponsors.** A fund can have several sponsors and a sponsor can back several funds, so `FUNDSPONSOR` holds the pairs.
- **Referential integrity.** Every foreign key uses `ON UPDATE CASCADE` and `ON DELETE NO ACTION`, so a project with funds or activities cannot be deleted by accident.
- **Text types for codes.** ZIP codes, phone numbers and fund codes are stored as text so leading zeros survive.

## Automation

Access runs only one SQL statement per query, so building the schema by hand means opening about fifteen queries in the right order. A small form with two buttons does it instead ([`vba/automation_form.bas`](vba/automation_form.bas)):

1. **Build tables** loops over the saved queries, running every `CREATE TABLE` first and every `ALTER TABLE` second.
2. **Insert values** loads the sample data in dependency order (parent tables before child tables).

The application opens on a navigation form with one tab per table.

## Repository contents

| Path | What it is |
| :--- | :--- |
| `university-project-activities-database.accdb` | The Access database: tables, saved queries, automation form, navigation form |
| `sql/01_create_tables.sql` | `CREATE TABLE` statements with primary keys |
| `sql/02_add_foreign_keys.sql` | `ALTER TABLE` statements adding the seven foreign keys |
| `sql/03_sample_data.sql` | Sample data as `INSERT` statements |
| `sql/04_reporting_queries.sql` | Six reporting queries |
| `vba/automation_form.bas` | VBA behind the two automation buttons |
| `The University Project Activities Database Project-Sijie Hu.pdf` | Full design report (22 pages): requirements, E-R model, normalization, SQL, application |
| `university-project-activities-database.pptx` | Editable source of the E-R diagram |

The `.sql` files are there so the schema can be read on GitHub without opening Access. They use Access SQL (`CURRENCY`, `SHORT`, `#date#` literals, `Nz()`).

## Running it

Requires Microsoft Access 2016 or later on Windows.

1. Download `university-project-activities-database.accdb` and open it. Choose **Enable Content** so the VBA can run.
2. The navigation form opens with the tables already built and loaded.
3. To rebuild from scratch, delete the tables, open the automation form, and click **Build tables**, then **Insert values**.
4. To try a reporting query, go to Create > Query Design > SQL View and paste one statement from `sql/04_reporting_queries.sql`.

## Sample data

All sample data is fictional: project names, IDs, award references, sponsors, people, contact details and balances were made up to exercise the relationships. There are 8 projects, 17 activities, 7 funds, 4 sponsors, 4 departments and 6 faculty.

## Limitations

- Reports are not built yet. The six queries in `sql/04_reporting_queries.sql` are the starting point for them.
- Each fund belongs to exactly one project, and fund balances are a single figure with no transaction history.
- Access is a single-file desktop database. The same schema would move to SQL Server or PostgreSQL with minor type changes.

## Author

Sijie Hu, MS Information Systems (Data Analytics), Baruch College. [LinkedIn](https://linkedin.com/in/thesjhu)
