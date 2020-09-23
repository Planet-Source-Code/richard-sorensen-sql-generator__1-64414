<div align="center">

## SQL Generator


</div>

### Description

This application allows you to quickly build an industrial strength database program, and adds many safeguards and features that you would otherwise have to code for yourself. The app accepts database schema definitions in the form of an XML file and generates all of the code necessary to access SQL databases in the form of a VB module that you insert into your application. Operations that are performed for you include the following:

* Handles all of the issues related to initialization strings and opening the database(s).

* Generates correctly formatted SQL statements for both queries and updates (errors are caught at generation time rather than when the code is run). You can rerun the generator at any time when the schema changes, and the code will then always be correct.

* Handles SQL error recovery, display, and logging of errors automatically.

* Deals with many SQL gotchas, such as escaping string field values containing single quote chatacters, handling maximum text field lengths, date field formatting for queries and inserts, verifying numeric field values, recovery from null field errors, etc.

* Resolves many of the subtle issues related to differences between various SQL systems (MySQL, Oracle, Access, Postgres, etc.)

* Provides a very easy-to-use programming interface to the SQL system.

The logic handles many database programming situations such as multiple database connections, multiple table buffers, using SQL DML v.s. recordset operations, defining the system to generate only the desired type of access for each table, etc. The generated code can also be used for VBA applications.

This system has been in operation for a number of years, and therefore is very solid and well-tested. Full documentation and a sample is also included.
 
### More Info
 
This is a complete application for generating a VB code module to access one or more databases.

Please read the included documentation to understand how this app is used.


<span>             |<span>
---                |---
**Submitted On**   |2005-11-21 12:02:48
**By**             |[Richard Sorensen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-sorensen.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[SQL\_Genera1975452222006\.zip](https://github.com/Planet-Source-Code/richard-sorensen-sql-generator__1-64414/archive/master.zip)

### API Declarations

Everything is included





