 
 
    SQL ADO Database Access Code Generator for Visual Basic 5/6 and .Net
 
 
    By:  Richard B. Sorensen
         Westernesse Corporation
         rich@westernesse.com
         www.westernesse.com
 

    INTRODUCTION
    ------------ 

    This system generates code to perform all of the routines that are commonly
    needed in a database application to access a series of database tables.  
 
    These routines are used to perform common database access functions (query,
    add, update, and delete) as well as all logging and error-handling.  The 
    database tables which are the source and target of these routines may exist
    in any type of database system that is accessible via MS ADO, with limitations
    as discussed below.
  
    All of the code necessary to perform SQL functions is automatically generated.
    The developer merely creates an XML file defining the database schema, as
    shown below.  Then the SQLGenerator program is run, and the name of XML file
    file is specified.  The generator will then create a VB module file with the
    code and data blocks necessary to access the database.

    Following are some of the benefits of using a generation system:

    1.  All of the logic is generated from the schema definition, so as your SQL
        schema evolves, all you have to do is rerun the generator to support the 
        changes.  The system uses the same internal names for fields as the schema
        definition, so that if you change the SQL definition you will receive  
        compile-time errors, and can therefore easily locate and fix the problems
        (rather than having these errors appear in the run-time logic after you
        have shipped the code to your customers).

    2.  The code automatically provides record buffers for each table, so you don't
        have to create your own working variables to handle SQL data.  You can also
        easily create additional buffers without having to define any of the fields
        in the table.

    3.  Many applications use strings in their code to perform SQL operations
        (such as "SELECT custName FROM CUSTOMER WHERE ..."), which are not verified
        until that code is run.  If field names are misspelled or no longer exist, or
        if the syntax is incorrect, the error is not caught until run-time.  All
        queries (for deleting, selecting, and updating) can be defined for the 
        application, and if so defined, the generator will catch these problems at
        generation time, and will insure that all of your SQL code is correctly
        formatted. Using this facility therefore means more inherent robustness and
        higher quality.  However, you also code your SQL strings directly if you
        prefer to do it that way.

    4.  There are a large number of SQL "gotchas" that need to dealt with in a robust 
        application.  Issues such as surrounding text fields with single quotes, 
        making sure that text fields are properly escaped to handle embedded single 
        quote characters (different of different database systems), truncating text 
        fields to fit within database-specified limitations, making sure that only
        numeric values are supplied for numeric fields, properly formatting date values
        for both queries and updates, automatic recovery from attempting to read null
        values, etc. are all handled for you automatically.

    5.  Error handling and logging is automatically performed - no more SQL crashes or
        hung apps.  You can specify either silent mode (logging only) for unattended
        batch processing applications, or logging plus message displays depending on
        what kind of app you are writing.

    6.  To a large extent, differences between SQL platforms are masked.  Not all 
        differences can be handled automatically, but most are.  Furthermore, the
        generator uses a text template-file, so that additional logic to support
        database-specific situations can easily be added.

    7.  For insert and update functions, you can choose between the recordset-based
        operations (such as AddNew, Delete, and Update) or SQL DML (Insert, Delete
        From, and Update Set).  For each table you can specify only the routines 
        that you need (i.e., some tables in your database may be read-only).

    8.  Utility functions such as retrieving record counts, searching a resultset,
        determining the primary key of an inserted record (for autonumber keys), 
        doing compaction and repairs on Access databases, and others are supplied.
 

    These routines are intended for add/modify/delete database applications,
    involving straight-forward, non-complex queries.  They are NOT intended for
    decision support systems or data warehousing applications where the primary 
    purpose of the application is the construction of complex queries involving
    SQL joins, group-bys and other summary operations.  However, if complex
    queries have been defined in the target database system as stored procedures,
    then these routines can be easily used.
   
    In particular, the generated code does not support SQL join operations (all
    queries in the code will target only one table at a time), and the queries
    always return all of the defined fields from the specified table (i.e., a 
    SELECT * is always used).  The reason for these limitations is that the code
    generates a record buffer containing fields for each table, and this buffer 
    is the primary interface between the SQL code and the user's application; if
    selected fields and/or multiple table queries would be allowed, the record
    buffers would only be partially valid, and many problems would result.

    However, If you wish to use this generator and also perform join operations
    or other complex queries, they can easily be done - you simply use the
    connection and recordset objects generated for you, and supply your own SQL
    query logic.
 
 
 
    SPECIFYING THE DATABASE SCHEMA
    ---------- --- -------- ------
 
    The schema for the database is specified in an XML-formatted file, which
    should be created with a plain-text ASCII text editor.  The format of the
    file is as follows:
 
        <Output module-file-name>

        <Options ...>

        <Table TableName1 [BufferCount] [AccessType]>
            FieldName1    String(length) [AutoKey] [Primary]
            FieldName2    Boolean
            FieldName3    Double
            ...
        </Table>

        <Table TableName2>
            ...
        </Table>
        
        ... (more table definitions)
        
        [optional query definitions here - see below]


    For example:

        <Output modSQL.bas>

        <Option vb6 DeleteFrom Execute Insert RecordCount UpdateInto>

        <Table Customer>
            RecordID      Long AutoKey
            RecType       Integer Index Unique
            Name          String(50) Index
            AddressLine1  String(30)
            MaxCredit     Double
            LastOrderDate Date
            ZipCode       String
            CurrentCust   Boolean
            ...
        </Table>
        
        <Table Order>
            RecordID      Long Autokey
            CustomerID    Long
            OrderAmt      Double
            ...
        </Table>
                
        [optional query definitions here - see below]
        

    The VB code module, specified in the Output statement above, will be generated
    by this process, and can then be included in a Visual Basic project and used.
    The file extension of the file name should be "bas" in keeping with the VB file
    naming conventions. The code may also be generated for either Visual Basic
    version 6 or VB.Net formats - for VB.Net add the keyword "VBNet" to the option
    values, and for VB6 add the keyword "VB6" or don't use a keyword (VB.Net support
    may not be available).  

    If you wish to use this generator for VBA applications, use the keyword "VBA" in
    the Option statement.  This uses the VB6 code base with minor modifications.

    If you wish to insert comments in the schema definition or "comment-out" certain
    parts of it, use the tokens "<!--" to begin the comment and "-->" to end it.
    Also, lines beginning with a single quote character ("'") will be ignored.
 
    Multiple tables may be defined as shown above, and each one must be started
    with a "<Table TableName>" line, followed by a series of lines defining each
    field in the table.  The data type of each field must be specified - the valid
    types are:
 
        Boolean  - 8 bit true/false
        Byte     - 8 bit integer
        Currency - 64 bit integer
        Date     - 8 byte date/time
        Double   - 8 byte floating point
        Int8     - 8 bit integer
        Int16    - 16 bit integer
        Int32    - 32 bit integer
        Int64    - 64 bit integer
        Integer  - 16 bit integer for VB6, 32 bit for .Net
        Long     - 32 bit integer for VB6, 64 bit for .Net
        Memo     - long string
        Single   - 4 byte floating point
        Short    - 8 bit integer for VB6, 16 bit for .Net
        String   - short string

    Various database system support different type of fields and various naming
    conventions for the field types.  For the most part this is automatically
    handled by the ADO logic.  However, there may be no direct analog for certain
    types in Visual Basic (such as timestamp, IP address, unsigned integer, blob,
    etc.), so those fields cannot be used unless they will map directly to a
    supported Visual Basic data type.  Note that "variants" are not allowed.

    For String and Memo fields, the maximum desired database size of the fields
    can be specified, as follows:

        String(50) (can be used in indexes, whereas memo fields cannot)
        Memo(2000) (maximum size in MicroSoft Access is 65,536 bytes)

    Most database systems impose a farily small maximum length on string/text
    fields (usually 255 characters).

    The benefit of indicating the length is that the generated SQL code will then
    automatically truncate text beyond the maximum length, thus avoiding SQL 
    errors when insert and update operations are performed.  If no size is given,
    the automatic truncation feature will not be performed for that field.

    In addition to the field type, other characteristics of each field may be
    specified by adding keywords after the data type, as follows:

        AutoKey
        Primary
        Index
        Unique

    If the database field consists of an automatically generated value (e.g., the
    "AutoNumber" field type in Access), then the keyword "AutoKey" must be added
    after the data type (e.g., "KeyField Long AutoKey").  This is necessary to
    prevent the Insert and UpdateInto functions from attempting to store a value
    in that field.

    The "Primary" keyword indicates that the field is the primary key of the table
    and that a unique index exists for the field.  This value is only used in the
    Create Table logic.

    The "Index" keyword indicates that the field is indexed in the table (the
    "Primary" and "Index" keywords are mutually exclusive).  If the "Unique" 
    keyword is also given, then values for this field must be unique in the table.
    This value is only used in the Create Table logic.
 
    Each time the generator is run it will regenerate and overlay the contents of
    specified code module.  The code module should therefore not be manually 
    modified as the changes will be wiped out the next time that the generator is
    run.  If you wish to define additional SQL-related routines place them in 
    another module.
 
    The values on the Option statement above are used to include optional
    routines.  Multiple option values are entered on the same line and must be
    separated by one or more spaces.  Following are the possible values.
 
        AddNew        - includes routines for each table to insert new SQL records
                        using a recordset.
 
        Clear         - includes routines for each table to clear the fields in the
                        record buffers for the table.
 
        CompactRepair - includes a routine for performing compaction and repair
                        on Access Jet databases (Access 2000 & above)
                     
        Delete        - includes routines for each table to delete existing records
                        from a recordset.
 
        DeleteFrom    - includes routines for each table to delete existing records
                        without requiring a recordset.

        Insert        - includes routines for each table to insert new SQL records
                        without using a recordset.
 
        Parameter     - includes code to generate ADO parameters for each field,
                        if the develop wants to perform inserts and updates with
                        command parameters.

        RecordCount   - includes code to perform record count queries on each table.
 
        Search        - includes routines for each table to search a recordset for
                        a specific record.
 
        Transaction   - includes transaction-related routines for each connection
                        (BeginTrans, CommitTrans and RollbackTrans).
 
        Update        - includes routines for each table to update existing records
                        from a recordset.
 
        UpdateInto    - includes routines for each table to update existing records
                        without requiring a recordset.

        VB6           - generate the code for compatibility with Visual Basic version 6

        VBA           - generate the code for compatibility with Visual Basic for 
                        Applications

        VBNet         - generate the code for compatibility with Visual Basic .Net
 
 
        Note that the routines to add, modify, and delete records are all optional.
        This was done so that the user can select the routines which require a recordset
        for database modification (AddNew, Delete, and Update) or the SQL DML routines
        which do not require a recordset (Insert, DeleteFrom, and UpdateInto).  Any of
        these routines can be included in any combination that the user desires.  
        Furthermore, SQL execute routines are automatically supplied, and can be used to
        perform free-form SQL statements, such as altering or dropping tables, setting
        SQL security, running stored procedures, etc.



    DEFINING MULTIPLE DATABASE CONNECTIONS
    -------- -------- -------- -----------
 
    In the above example, all tables will use the same default connection object, 
    and therefore all tables must exist in the same database.  However, if the tables
    are from different databases, or tables in the same database need to be accessed
    from more than one connection, then multiple connection objects are needed.
    The following structure must be used for multiple connections:
 
        <Connection ConnectionNameA>
        <Table TableNameA1>
            ...
        </Table>
        
        <Table TableNameA2
            ...
        </Table>
        </Connection>
        
        <Connection ConnectionNameB>
        <Table TableNameB1>
            ...
        </Table>
        
        <Table TableNameB2>
            ...
        </Table>
        </Connection>
 
    The "ConnectionName" above is an arbitrary value supplied by the developer to
    identify the database connection.  This value must be different for each connection,
    and the value will be added to the beginning of each table name in the generated
    code.  For example, for a table named "Customer", the routine name used to add a
    record if no connection is specified would be "SqlInsertCustomer".  If, however, a
    connection with the arbitrary name of "Dbx" is defined for the customer table, the
    insert routine will be named "SqlInsertDbx_Customer".

    It should be noted that the Microsoft ADO documentation strongly discourages the
    use of multiple connections on the same database unless they are absolutely
    necessary. 

 
 
    DEFINING QUERIES
    -------- -------

    Following the option statement, queries can be defined.  Each query will create a
    callable routine that will run the query.  It is not necessary to create any queries
    - you can construct your own WHERE clauses as string variables and call the 
    SELECT routines directly for the associated tables.  However, The benefits of
    defining your queries is that the system will verify that the field names exist in
    the table, and will automatically generate the correct syntax.  For example, text
    values will automatically be surrounded by single quote characters, boolean values
    will be specified by the strings "true" or "false", etc.  Furthermore, the system
    will scan string parameters for unacceptable characters, such as single quotes, and
    automatically escape them.

    The general form of a query is as follows:

        [<QueryConnection ConnectionName>]
        
        <QuerySelect TableName Queryname [BufferNumber]>
            (      (optional)
            FieldName Operator (=, <>, >, <, >=, <=, LIKE) % ["format string"]
        
                   (following lines are optional)
            AND or OR
            FieldName Operator (=, <>, >, <, >=, <=, LIKE) fixed value
            )      (optional)
            ...
	           (following lines are optional)
            ORDER BY
            FieldName
            ...
        </Query>
        
        ... more query definitions
        
        [</QueryConnection]

    Following is an example, using the Customer table from above:

        <QuerySelect Customer Credit>
            (
            MaxCredit > %
            AND
            RecType = 1
            AND
            LastOrderDate < % "mm/dd/yyyy"
            )
            ORDER BY
            ZipCode
            Name
        </Query>

    The system will generate the following callable routine from the above 
    definition:

        Public Function SQLQuerySelectCustomer_Credit(MaxCredit1 As Double, LastOrderDate2 As Date) As Boolean
            ...
        End Function

    When this function is called it will generate the following SELECT statement,
    with the parameter values automatically inserted:

        SELECT * FROM Customer 
            WHERE (MaxCredit > % AND RecType = 1 AND LastOrderDate < #%%/%%/%%%%# )
            ORDER BY ZipCode, Name

    Note that the field "RecType" field above has a fixed parameter supplied with the
    query definition, so an external parameter will not be generated, whereas the
    code will insert the value for "MaxCredit" and "LastOrderDate" from the values
    specified when the routine is called.

    Any operator can be used (such as "LIKE") as long as it it valid for the database
    system in use.  

    The "(" and ")" lines will add left and right parentheses characters, and can be
    inserted into the flow wherever parentheses are necessary.  However, it is up to 
    the developer to insert them in the correct position along with the appropriate
    "AND" and "OR" conjunctions so that a valid SQL statement is generated.

    The developer can create as many queries as necessary, but each query must have
    a name that is unique within the queries defined for each table.

    If multiple connections are defined (see above), the query must be prefaced by
    a "<QueryConnection>" line giving the name of the connection to be used for this
    and subsequent queries until an "</QueryConnection>" line is reached.

    For select-type queries, a query can be defined to use other buffer areas, if
    more than one buffer has been defined for a table (see below),  The query will
    then populate the specified record buffer with the records from the database.  
    This is done by adding the corresponding buffer number (2, 3, etc.) following
    the table name in the query definition.  

    Queries for performing delete and update operations may also be defined.  For
    example:

        <QueryDelete Customer Expired>
            LastOrderDate < % "mm/dd/yyyy"
            AND
            CurrentCust = false
        </Query>

        The above query definition generates the following routine:

        Public Function SQLQueryDeleteCustomer_Expired(LastOrderDate1 As Date) As Boolean
            ...
        End Function

            and generates the following SQL statement:

            DELETE FROM Customer WHERE LastOrderDate < #%%/%%/%%%%# AND CurrentCust = false


        <QueryUpdate Customer All>
            RecordID = %
        </Query>

        The above query definition generates the following routine:

        Public Function SQLQueryUpdateCustomer_All(RecordID1 As Long) As Boolean
            ...
        End Function

            and generates the following SQL statement:

            UPDATE Customer SET <all fields from the Customer table>

        <QuerySet Customer NameAddress>
            RecordID = %
            SET
            Name = %
            AddressLine1 = %
            ZipCode = %
            CurrentCust = true
        </Query>

        The above query definition generates the following routine:

        Public Function SQLQuerySetCustomer_NameAddress(RecordID1 As Long, Name2 As String, AddressLine3 As String, ZipCode4 As String) As Boolean
            ...
        End Function

            and generates the following SQL statement:

            UPDATE Customer SET Name = '%', AddressLine1 = '%', ZipCode = '%', CurrentCust = true WHERE RecordID = %

    Note that there are two forms of update queries - "Update" and "Set".  The fomer 
    will automatically update all of the fields in the table from the associated
    record buffer (i.e., it will automatically build an SQL SET statement), whereas the
    latter will only update the fields that you indicate in the query definition.  

    If delete queries are defined, the "DeleteFrom" option above must be included, and 
    the "UpdateInto" option is required for update queries.  Also the associated table
    definitions must be defined to allow delete and/or update operations.

 
 
 
    DEFINING ALLOWABLE I/O OPERATIONS FOR EACH TABLE
    -------- --------- --- ---------- --- ---- -----
 
    The generator assumes that all tables will need all of the I/O routines that are
    specified in the options section above.  But some tables may be read-only, and update
    code for them is not necessary.  Therefore the developer may supply a keyword after
    the table name to indicate that some I/O routines should not be generated for that
    table.
 
    Select and read routines will always be generated, but add, modify, and delete 
    routines may be selectively removed.  Following are the keywords that may be used:

        Create   - causes a CreateTable routine to be generated for that table
 
        NoWrite  - removes all add, modify, and delete routines for that table
 
        NoAdd    - removes add-related routines for that table
 
        NoDelete - removes delete-related routines for that table
 
        NoModify - removes modify-related routines for that table
 
 
    Following is an example of how this facility is specified:
 
        <Table TableName NoWrite>
            ...
        </Table>
 
 
 
    DEFINING MULTIPLE RECORD BUFFERS FOR A TABLE
    -------- -------- ------ ------- --- - -----
 
    A numeric value may be specified after the table name, which represents the
    number of recordsets and record buffers that will be defined for the table
    (the default is one).  The developer may define additional buffers to contain
    table data, but all read options will populate only the system-defined 
    buffer(s).  See below for examples on how this facility is used
 
 

    UNDERSTANDING THE GENERATION PROCESS
    ------------- --- ---------- ------- 
 
    1.  The target VB program or activeX dll must specify "Microsoft ActiveX 
        Data Objects" under Projects.References.  The ADO dll becomes an additional
        dependency, which must be installed with the other elements of the
        application.  Also note that ADO requires that MDAC (Microsoft Data Access
        Components) Version 2.5 or above be installed as well.  Installers such as
        the MS Package and Deployment Wizard, InstallShield, Wise, etc. will usually
        detect that the project requires ADO, and will automatically include the
        MDAC stuff in the installation.  However, it should be noted that the Jet
        infrastructure necessary for Access databases is no longer included in 
        MDAC 2.6 and above, and must be installed separately.

 
    2.  Some database systems do not supply "provider" dlls for the Windows 
        environment and therefore must be accessed via an ODBC connection, using
        a DSN value.  These system include MySql, PostgresQL, and others.  It is
        the responsibility of the developer to include the necessary ODBC setup in
        any installation process.
 
 
    3.  The table and field names specified must match the actual names used in
        database, and the field types must be properly specified.  Names must not
        begin with a number or contain embedded spaces.
 
 
    4.  The VB module must be "checked out" if a source code archiving system is
        being used.  In other words, it must not have the read-only property set.
        Also, the developer must have both update and delete file rights on the
        code file, and on the directory in which the file is located.
 
 
    5.  The generated code module will be overwritten each time the generator is
        run.  The file will be placed in the same path as the XML schema file.
        If the database schema changes, the XML file must be updated to reflect
        the changes, and the generation process rerun.  This can be done as often
        as necessary.
 
 
    6.  The generator inserts a data structure for the table and defines a variable
        by which the developer can manipulate the fields from the current record, as 
        shown below.
 
 
    7.  The generator will support database table and field names that are not valid
        Visual Basic fields, such as "emp-name#" (VB variable names cannot contain
        dashes or other special characters).  As table and field names are parsed,
        an "internal" version of the name will be created by replacing all non-
        alphanumeric characters with underscores.  The internal version of the name
        must then be used when referring to the field in the VB application, and the
        external name will automatically be used for queries and updates sent to the
        database server.  For example, the field "emp-name#" in the table "User-Pro"
        would be referred to as "g_recUser_Pro.emp_name_".
 
 
    8.  Prior to running the generation process, the code module should be removed 
        from the VB project.  The generator can then be run and the code module
        readded to the project.  If this step is not followed, the modified version
        of the module will not be used until the project is terminated and restarted.
        Alternatively, you can terminate and rerun the VB project.


    9.  The developer can modify the generator code or add additional routines by
        modifying the file SQLCode.txt, which is the source code template file.
 
 
 
    UNDERSTANDING AND USING THE GENERATED CODE
    ------------- --- ----- --- --------- ----
 
    1.  The developer will not need to write any SQL code if his needs for SQL access
        are handled by these routines - see below for the limitations.
 
 
    2.  The generator inserts a "record buffer" for each table that will be used for
        fetching and storing the contents of the current record from that table.
        Insert, select, and update routines are always performed on every field in the
        table.  Following is a sample record buffer:
 
            Type typTableName
                FieldName1 As String
                FieldName2 As String
                FieldName3 As Double
                ...
            End Type
            Public g_recTableName As typTableName
            Public g_objRecordsetTableName As ADODB.Recordset
 
        For example, to get or set the value of "FieldNameX", the developer would refer
        to "g_recTableName.FieldNameX"
 
 
    3.  The following routines are provided (in all cases, the connection name is blank
        if no connections were declared):
 
        For each connection, the following routines are always generated:
 
            SQLOpenConnection[ConnectionName]    - opens a connection to a database
            SQLCloseConnection[ConnectionName]   - closes the database connection
            SQLCommand[ConnectionName]           - executes an SQL command for the connection
 
        If specified, the following optional routines are generated for each connection:
 
            SQLBeginTrans[ConnectionName]    - Begins a database transaction
            SQLCommitTrans[ConnectionName]   - Commits a database transaction
            SQLRollbackTrans[ConnectionName] - Rolls back a database transaction
 
        For each table, the following routines are always generated:
 
            SQLSelect[ConnectionName_]TableName - runs a SQL query to generate a recordset on the table
            SQLNext[ConnectionName_]TableName   - navigates to the next record in an open recordset
 
        If specified, the following optional routines are generated for each table:
 
            SQLAddNew[ConnectionName_]TableName      - adds a new record to the table using a recordset
            SQLClear[ConnectionName_]TableName       - clears the fields in the table's record buffer
            SQLCreateTable[ConnectionName_]TableName - creates the specified table
            SQLDelete[ConnectionName_]TableName      - deletes a record from an open recordset on the table
            SQLDeleteFrom[ConnectionName_]TableName  - deletes one or more records in the table
            SQLFetch[ConnectionName_]TableName       - populates the user buffer from the current record
            SQLInsert[ConnectionName_]TableName      - adds a new record to the table without a recordset
            SQLParameter[ConnectionName_]TableName   - executes an SQL command on the table (if parameters were declared)
            SQLRecordCount[ConnectionName_]TableName - returns a record count for the table
            SQLSearch[ConnectionName_]TableName      - searches an open recordset for a specific record
            SQLUpdate[ConnectionName_]TableName      - updates the current record and stores it in the table
            SQLUpdateInto[ConnectionName_]TableName  - updates one or more records in the table
 
        The following utility routines are also generated
 
            SQLCompactRepair - (optional) compacts and repairs a Jet database.
            SQLCreateTable   - creates an SQL Table
            SQLExecute       - executes an SQL command
            SQLErrorHandling - turns on/off error display and error logging.
            SQLErrorCount    - returns a count of errors since the first routine call.
            SQLErrorLast     - returns the error message string from the last SQL
                               function (returns null string if there were no errors). 
 
 
    4.  All database access routines above return a boolean result ("True" for
        success, "False" for failure), but in some cases failure is a normal
        condition.  For example, the Select and Next routines return false if
        the recordset is empty, or the last record in the recordset was reached.
        If an actual failure occurs, an error message will automatically be
        displayed by the routines (if error display has been turned on).
        To suppress error messages and instead log them to a file, use the 
        SqlErrorHandling function.  The SQLErrorLast function can be used to
        determine if an error actually occured.
 
 
    5.  Multiple record sets and record buffers may automatically be generated for
        each table as indicated above.  This is done by specifying a number after the
        table name in the "<Table>" line.  If no number is specified, only one recordset
        and one buffer is generated.  In the following example, the table definition
        specifies 3 recordsets and buffers, and will generate the code shown:
 
            <Table TableName 3>
                 ...
            <Table>
 
            Type typTableName
                ...
            End Type
            Public g_recTableName As typTableName
            Public g_objCommandTableName As ADODB.Command
            Public g_objRecordsetTableName As ADODB.Recordset
            Public g_recTableName_2 As typTableName
            Public g_objCommandTableName_2 As ADODB.Command
            Public g_objRecordsetTableName_2 As ADODB.Recordset
            Public g_recTableName_3 As typTableName
            Public g_objCommandTableName_3 As ADODB.Command
            Public g_objRecordsetTableName_3 As ADODB.Recordset
 
        The generated code will include routines for all of the recordsets and buffers.
        If only one buffer is defined (no number was specifed), the following will be
        generated for the table:
 
            SQLSelect[ConnectionName_]TableName
            ...
 
        However, if three buffers are defined as shown above, then routines for each
        buffer and recordset will be generated as follows:
 
            SQLSelect[ConnectionName_]TableName
            SQLSelect[ConnectionName_]TableName_2
            SQLSelect[ConnectionName_]TableName_3
            ...
 
        Each buffer is associated with a recordset as shown above, and when select or
        other SQL operations are performed, the associated record buffer will always
        be used as the source or destination of the recordset data.  The is true for
        the following routines:  Select, Next, Search, and Update.

        The developer may, of course, copy data from one buffer area to another, as
        needed in his application code and may create additional record buffer 
        instances as necessary. 

        Other routines (in particular the SQL DML routines Insert and UpdateInto)
        allow you to work with any buffer, as long as it is defined from the associated
        table definition type declaration.  The buffer area to be used is passed to the
        routine by reference.
 
        The multiple buffer facility may be used in conjuction with table I/O routine
        limitations, and the associated values may be specified in any order.
        For example:
 
            <Table TableName 3 NoDelete>
                ...
            </Table>
        
 
    6.  The SqlOpenConnection routine is called to open a connection to a database.  
        There are two ways to make the connection:
 
            DSN - supply a DSN value (the name of a control panel ODBC definition)
            ConnectString - supply a database connection string
 
        The first argument for the SQLOpenConnection function is the DSN value.
        A DSN (data source name) refers to an ODBC declaration set up in the ODBC
        control panel function (Win98 and NT), or the Data Sources function (Win2000
        and XP).  All of the parameters necessary to make the connection must be 
        specified in the DSN definition.
 
        If a DSN will not be used, this value should be left blank, and the the second
        argument (ConnectString) must be used, as well as the following ones (server
        name, database name, etc.)  Connect strings contain a series of keywords and
        values separated by semicolons (";") which specify the connection parameters.

        Care must be taken on connect strings to get the correct combination of
        attributes - different database systems and providers require different values.

        If you use a connect string there are a series of "replacement strings" that 
        can be added in the appropriate place in the string.  The SQL open routine will
        replace these strings with the correponding value from the calling parameters
        so that you won't have to construct the actual string.
         

            <DB>  - the name of the database, if any.  This string is replaced by the
                    database name value supplied in the SQLOpenConnection call.

            <Srv> - the name of the server, if any.  This string is replaced by the server
                    name value supplied in the SQLOpenConnection call.

            <UID> - the database username, if any.  This string is replaced by the username
                    value supplied in the SQLOpenConnection call.

            <Pwd> - the password, if any.  This string is replaced by the password value
                    supplied in the SQLOpenConnection call.


        Here are some common examples of connect strings for use with SQLOpenConnection
        (parameters are case insensitive):
         
            Access
                "PROVIDER=MSDASQL; DRIVER=Microsoft Access Driver (*.mdb); INITIAL CATALOG=<DB>; UID=<UID>; PWD=<Pwd>"
                "PROVIDER=Microsoft.Jet.OLEDB.4.0; DATA SOURCE=<DB>; Jet OLEDB:Database Password=<Pwd>"
                (the <DB> value must be the fully qualified name of the database)
 
            SQL Server:
                "PROVIDER=MSDASQL; DRIVER={Sql Server}; SERVER=<Srv>; DATABASE=<DB>; UID=<UID>; PASSWORD=<Pwd>"
                "PROVIDER=SQLOLEDB; DATA SOURCE=<Srv>; DATABASE=<DB>; USER ID=<UID>; PASSWORD=<Pwd>"
                (the server value is the machine name or the IP address of the SQL server computer)

            Oracle:
                "PROVIDER=MSDASQL; DRIVER={Microsoft ODBC for Oracle}; SERVER=<Srv>; DATABASE=<DB>; UID=<UID>; PASSWORD=<Pwd>"
                "PROVIDER=MSDORA; DATA SOURCE=<Srv>; USER ID=<UID>; PASSWORD=<Pwd>"
                "PROVIDER=OraOLEDB.Oracle; DATA SOURCE=<Srv>"
                (where the server name is the name of database declararation on the Oracle server)
 
 
        Other database systems such as Informix, MySql, PostgresQL, Progress, Sybase,
        etc. can easily be connected via OBDC, and then the DSN name can be used, as
        indicated above.  For more information, refer to the vendor or the MS ADO
        documentation.  
 
 	There is an optional parameter in the open routine call to specify the
 	database system being used (the default is "SqlServer" if no parameter is
 	specified).  The value is case-insensitive, and the choices are:
 
 	    Access
            DB2
            Informix
            Ingres
            MySql
 	    Oracle
            PostgresQL
 	    Progress
 	    SqlServer
            Sybase
 
        There are a number of subtle differences between the systems that affect the
        record set cursor types, the syntax used for SQL statements, the handling of
        specific data types during retrieval and update operations, and others.
        For example, in connections to an Oracle server the database cursor
 	type will be set to "static," which is the only cursor type supported
 	by Oracle.  Also the "ServerPrefix" property will be automatically set by
 	the Open routine if the database system is Oracle.  This is used when
        constructing SQL statements, and is prepended to table name, for example:
 
            SELECT * FROM DatabaseServer.TableName WHERE ...


        Most of the differences between database systems are handled, but there are
        some things that the user must take care of.  This is especially true of
        select syntax, which, for the most part is user-specified.  The garden 
        variety queries (such as "numField > 10 AND strField = 'XXX'") are the same
        in all systems, but more esoteric operators are different bewteen the systems.
        Things to be aware of include date comparisons, string matching operations 
        (using keywords such as LIKE), comparisons between values, aggregation 
        operators, etc.  This code provides no analysis of whether a given select
        will work on a given database system.  However, the template file can be
        extended to automatically handle various cases (it has not been tested
        with all possible queries on all database platforms).
 
        There are two optional timeout values that can be set in the Open call
        The Connect timeout specifies the number of seconds to wait to a connection
        from the database server before giving up.  The Command timeout value
        specifies the number of seconds allows for all SQL functions following
        the connection process to complete.  The default value is 15 seconds for
        both parameters.  One or both of these values may have to be increased if
        SQL timeout errors are occurring.

        Other optional parameters are provided to specify the timeout values for
        connection and command requests, and to specify the location of the database
        cursor for the connection.  The default is for the cursor and the recordset
        objects to exist on the server, but this parameter can be used to specify
        client-side cursors.


    7.  An SQLCommand routine is generated for each connection.  This routine is used
        to execute any SQL statement for the connection.
  
 
    8.  The optional SQLBeginTrans, SQLCommitTrans, and SQLRollbackTrans routines are
        generated if the Transaction option is specified.  These routines supply
        transactional support for applications needing this capability.
 
 
    9.  The SQLSelect routine is called to run an SQL query and generate a recordset
        from the table.  The query string is specified as it would appear in an
        SELECT statement, and consists of the text following the WHERE keyword
        (i.e., it is not necessary to include the "WHERE" keyword at the beginning of
        the query string).  The Query routine will automatically populate the table's
        record buffer with the first record, if the recordset contained at least one
        record.  The routine will return "True" if the query suceeded and at least
        one record existed in the result set, otherwise it will return "False".  See
        below for a usage sample.
 
        This routine takes three parameters.  The first is the "WHERE" and/or "ORDER BY"
        clause which will be used to quality the SELECT statement.  The keyword "WHERE"
        is not necessary, but if both WHERE and ORDER BY text are included, the WHERE
        text must come first, followed by the literal "ORDER BY" and the field ordering.
 
        The second parameter is optional and is used to specify various recordset
        characteristics.  By default, a server-side, updateable keyset-based (or whatever
        is the default for database system) recordset with optimistic locking will be
        returned.  However, the following keywords can be included to modify the recordset
        characteristics
 
            Keyset          - returns a keyset-based recordset (if possible).
 
            Dynamic         - returns a dynamic recordset (if possible).
 
            Static          - returns a static recordset  (if possible).
 
            ForwardOnly     - returns a recordset that can only be navigated top to bottom.
 
            ReadOnly        - returns a read-only recordset.
 
            Optimistic      - sets the record locking to optimistic.
 
            Pessimistic     - sets the record locking to pessimistic.
 
            BatchOptimistic - sets the record locking to batch optimistic.
 
            StoredProc      - the SQL statement represents a stored procedure name.
 
            Table           - returns a table-type recordset.
 
        Keywords are case insensitive and may be specified in any order.  However, they
        must be separated by at least one space.  See the Microsoft ADO and/or the
        database system documentation for more information about these choices and the
        impact they have on your application.

        The third parameter is also optional and is used to limit the number of records
        returned by the Select operation to the top "n" rows.
 
 
    10. The SQLNext routine is called to navigate to the next record in a recordset,
        and it populates the record buffer with all of the fields from the record.
 
 
    11. The SQLFetch routine moves the contents of the current record into the
        associated memory buffer for manipulation by the program.  This routine is
        not normally needed, because it is automatically called by the Query and
        the Next routines.  It is required only for cases in which the developer 
        uses an alternate recordset navigation strategy (e.g., moving backward
        through the recordset).
 
 
    12. The optional SQLAddNew routine adds new SQL records to a table using the
        "AddNew" method.  In contrast to the SQLInsert function, this routine requires
        an open recordset in order to add the desired record.  The recordset must be
        updateable and can be opened as a "table" type, or as a normal recordset.
 
 
    13. The optional SQLClear routine clears the fields in a table's record buffer.
        This is useful routine to call prior to inserting a record, where not all of the
        fields in the record are explicitly set.  Fields are cleared to the default
        value appropriate for the field type - string fields are set to "", numeric
        fields are set to zero, boolean fields are set to false, etc.
 
 
    14. The optional SQLDelete routine deletes the current record from the associated
        recordset.  Therefore it requires an open, updateable recordset to operate
        against.
 
 
    15. The optional SQLDeleteFrom routine deletes one or more records from a table
        without using a recordset.  The "WHERE" clause in this function must be
        carefully specified - if no clause is provided, every record in the table will
        be deleted.
 
 
    16. The SQLExecute routine is used to run any SQL operation not explicitly
        defined by the other routines.  It can be used to run stored procedures, run
        SQL DML statements such as "INSERT INTO ..." and "UPDATE ... SET ...", or run
        SQL DDL statements such as "CREATE TABLE ...", "ALTER TABLE ...", "GRANT ...",
        etc.  Parameters can also be supplied if needed, by passing a variant array to
        the function.  However, the SQLExecute function cannot be used to run a 
        stored procedure that returns a recordset.  The SQLSelect function with the
        "StoredProc" option must be used for that purpose.
 
 
    17. The optional SQLInsert routine adds a new record to the associated table,
        but does not use a recordset.  A SQL Insert DML command is used instead,
        along with an ADO Command object.  The record buffer must be populated with
        the field values for the new record prior to calling this function.


    18. The optional SQLParameter routines should be used to exeute SQL statements
        if the Parameter option has been specifed (uses ADO pararemter objects for
        defining each field).


    19. The optional SQLRecordCount routine returns a count of the number of records
        in the table.  It also accepts an optional WHERE clause.
 

    20. The optional SQLSearch routine searches an open recordset, and positions it
        to the matching entry, if one exists.
 
 
    21. The optional SQLUpdate routine updates the current record from the associated
        recordset.  Therefore it requires an open, updateable recordset to operate
        against.
 
 
    22. The optional SQLUpdateInto routine updates one or more records from a table
        without using a record set.  The "WHERE" clause in this function must be
        carefully specified - if no clause is provided, every record in the table will
        be updated.
 
 
    23. A facility is supplied to eliminate error message (MsgBox) displays for batch
        programs.  Instead of displaying error messages, they will be written to a log
        file.  To use this facility, call the SQLBatch routine AFTER calling
        SQLOpenConnection, and provide a fully qualified (i.e., including path name)
        error log file name in the function call.  Errors will automatically be written
        to that file in text format.  The SQLBatch routine will return True if the file
        could be opened, or False if not.
 
 
    24. The number of errors which have occurred since the connection was opened may
        be obtained by calling the SQLErrorCount function.
 
 
    25. The error message (if any) from the last SQL operation that was done can be
        obtained by calling the SQLErrorLast function.  This is especially useful in 
        determining if an error occured in a SQLSelect or SQLNext operation (these
        routines return "False" when they encounter an error, but also when there are
        no more records to be returned.
 
 
    26. Sample usage of these routines for a table named "Customer":
 
        Dim blnSQL As Boolean
        Dim StrSQL As String

        If Not SQLOpenConnection("", "xxx", "Server", "Database", "User", "Password") Then
            (Handle any connection error)
        End If
        ...
        strSQL = "CustType = '10' AND CustSales > 100000"
        blnSQL = SQLSelectCustomer(strSQL, "")
        Do While blnSQL
            g_recCustomer.CustType = 20
            ...
            If Not SQLUpdateCustomer() Then
                (Handle update error as appropriate)
            End If
            ...
            If Not SQLAddNewCustomer() Then
                (Handle insert error as appropriate)
            End If
            ...
            blnSQL = SQLNextCustomer()
        Loop   
 
 
    27. If the CompactRepair option is specified, a routine to perform this process
        for Access Jet databases is included.  Setting this option requires that the
        "Microsoft Jet Replication Objects" library, MSJRO.DLL be included in the 
        project as a reference.  This dll file then becomes a dependency, and must be
        installed along with everything else.  
 
        This function will always set the format of the database to be Jet 4x 
        (Access 2000/2002), so this it will not work properly with older version of
        Jet databases unless source code changes are made in the routine.  The specific
        modification necessary is the "Engine type" parameter, which specifies which
        version of Jet database will be generated.  See Microsoft's MSJRO reference
        knowledge base articles for more information.
 
 
        The database must be closed prior to running this function, and must be
        reopened after the routine is finished.  If this process is unsuccessful, the
        calling program will not be able to continue, and should terminate with an
        appropriate error message to the user.


    28. Two "callback" routines may be defined in another module:  

        Callback_MsgboxPreProcess - will be called prior to displaying a message box to
        perform any pre-processing necessary to handle a modal message box display.

        Callback_SQLConnectFailure(recConnection As typSQLConnection) - will be called
        following a connection failure to allow programmatic recovery.

        In order to implement this processing, the constant enumSQLCallback must be set
        to "True" in the template code file.
 
 
    LIMITATIONS OF THESE ROUTINES
    ----------- -- ----- --------
 
    1.  The SQL queries always retrieve ALL of the fields from the table (the
        "SELECT * FROM Tablename ..." syntax is always used), and SQL updates
        always write ALL fields back into the table (unless "set" queries are
        defined).  Partial retrievals, SQL joins (queries involving multiple
        tables), and other summary operations such as GROUP-BY are not directly
        supported through these routines.  Complex SQL operations can, of course,
        always be added by the developer, and the generated command and recordset
        objects can be used to facilitate the added functionality, as needed.
        Another alternative is defining all complex queries as stored procedures,
        if the database system being used supports them.

        If the developer wants to update only one or two fields, that can easily be
        accomplished by specifying a "Set" query as discussed above, or by calling
        the SQLCommand or SQLExecute function.
 
        In any case, this SQL generation facility can be combined with other 
        ADO-based access logic that is used outside the scope of these routines.
        Once the ADO connection has been opened, any valid ADO operation can be
        called.
 
 
    2.  Except for the SQLSearch function, no routines are generated for moving
        through a recordset other then in the forward direction (using the 
        SQLNext function).  However the developer is free to use the ADO Move
        methods as desired (MovePrevious, MoveFirst, etc.) directly in his
        application code, provided that tests for EOF and/or BOF conditions are
        included.  For example:
 
            g_objRecordsetTableName.MoveLast
            Do Until g_objRecordsetTableName.BOF
                Call SQLFetchTableName
                ... (process record)
                g_objRecordsetTableName.MovePrevious
            Loop
 
        Note that moves inside a recordset other than "move next" are not permitted
        by ADO if the recordset has been defined as "ForwardOnly".
 
 
    3.  While SQL is a fairly universal standard, there are still some subtle
        differences between database systems that could potentially affect the
        portablity of this code from one system to another.  Differences include
        the following:

            . Date comparisons in SQL WHERE clauses.  MSAccess, for example, requires
              date literals to be enclosed in pound signs ("#").

            . Date fields.  The code template may require modification for some
              database systems.  This may also be true for other esoteric data types.

            . Field pattern matching clauses using operators such as "BETWEEN" and
              "LIKE".

            . Stored procedure and trigger capabilities.

            . Blob (binary large object) fields are not supported in the generator.
 
        The developer needs to be aware of these differences and make sure that they 
        are handled in the generated code, or in the external application logic.

        If any differences are noted and can be included in these routines, the
        SQL code template file can be modified to handle it.

