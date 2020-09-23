Attribute VB_Name = "modSampleSQL"
Option Explicit
Option Base 0
'
'
'   SQL Database Access Routines
'
'
'   By:  Richard B. Sorensen
'
'
'   This code module was automatically generated from a schema definition.
'   DO NOT MODIFY THIS CODE, and DO NOT ADD LINES OR ROUTINES to this module.
'   The generation process will wipe out all changes the next time it is run.
'
'

'Types ---------------------------

'Build SQL statement record
Private Type typSQLBuild
    blnActive As Boolean
    intBuildType As Integer
    strBuildField As String
    strBuildValue As String
    strDatabaseSystem As String
    strEscapeQuote As String
End Type

'Connection record
Public Type typSQLConnection
    blnDSN As Boolean
    intCommandTimeout As Integer
    intConnectTimeout As Integer
    intOperationType As Integer
    lngCursorDefault As Long
    lngCursorLocation As Long
    objConnection As ADODB.Connection
    objCommand As ADODB.Command
    strConnectionName As String
    strConnectString As String
    strConnectErrorString As String
    strDatabase As String
    strDatabasePassword As String
    strDatabaseSystem As String
    strEscapeQuote As String
    strFileName As String
    strPassword As String
    strRecordSetParams As String
    strServer As String
    strServerPrefix As String
    strUserName As String
End Type

'Enumerations ---------------------

Public Enum enumSQLBuildType
    CreateTable = 1
    InsertTable = 2
    UpdateTable = 3
End Enum

Public Enum enumSQLOperationType
    ConnectOperation = 0
    ReadOperation = 1
    WriteOperation = 2
    CriticalWriteOperation = 3
End Enum

'Constants ------------------------

'Callback conditional
#Const enumSQLCallback = False

'Database system types
Public Const enumSQLDatabaseAccess As String = "access"
Public Const enumSQLDatabaseDB2 As String = "db2"
Public Const enumSQLDatabaseInformix As String = "informix"
Public Const enumSQLDatabaseIngres As String = "ingres"
Public Const enumSQLDatabaseMySql As String = "mysql"
Public Const enumSQLDatabaseOracle As String = "oracle"
Public Const enumSQLDatabasePostgresQL As String = "postgresql"
Public Const enumSQLDatabaseProgress As String = "progress"
Public Const enumSQLDatabaseSqlServer As String = "sqlserver"
Public Const enumSQLDatabaseSybase As String = "sybase"

'Recordset parameters
Public Const enumSQLRecordsetBatchOptimistic As String = "batchoptimistic"
Public Const enumSQLRecordsetDynamic As String = "dynamic"
Public Const enumSQLRecordsetForwardOnly As String = "forwardonly"
Public Const enumSQLRecordsetKeyset As String = "keyset"
Public Const enumSQLRecordsetOptimistic As String = "optimistic"
Public Const enumSQLRecordsetPessimistic As String = "pessimistic"
Public Const enumSQLRecordsetReadOnly As String = "readonly"
Public Const enumSQLRecordsetStatic As String = "static"
Public Const enumSQLRecordsetStoredProc As String = "storedproc"
Public Const enumSQLRecordsetTable As String = "table"

'Other constants
Public Const enumSQLErrorLogFile As String = "SQLErrorLog.txt"

'SQL error codes
Private Const enumSQLErrorConnectFail As Long = 3709
Private Const enumSQLErrorConnectFailAccess As Long = -2147217865
Private Const enumSQLErrorConnectFailPostgres As Long = -2147467259
Private Const enumSQLErrorNetwork As Long = 3043
Private Const enumSQLErrorNullValue As Long = 94
Private Const enumSQLErrorMaxRetries As Integer = 3
Private Const enumSQLErrorUpdateDupKey As Long = 7

'Table names
Public Const enumTableNameExtCustomer As String = "Customer"
Public Const enumTableNameIntCustomer As String = "Customer"
Public Const enumTableNameExtOrder As String = "Order"
Public Const enumTableNameIntOrder As String = "Order"
Public Const enumTableNameExtOrderDetail As String = "OrderDetail"
Public Const enumTableNameIntOrderDetail As String = "OrderDetail"
Public Const enumTableNameExtSystem As String = "System"
Public Const enumTableNameIntSystem As String = "System"
Public Const enumTableNameExtVersion As String = "Version"
Public Const enumTableNameIntVersion As String = "Version"

'Declarations -----------------------------

Private Declare Sub Sleep Lib "kernel32" _
    (ByVal lngMilliseconds As Long)

'Data --------------------------------------

'Connection records
Public g_recConnection As typSQLConnection

'Miscellaneous module-level variables
Private m_blnErrorDisplay As Boolean
Private m_blnErrorLog As Boolean
Private m_lngErrorCount As Long
Private m_recBuild(0 To 2) As typSQLBuild
Private m_strError As String
Private m_strErrorLogFile As String

'
'SQL Table Definitions   Automatically Generated  -  DO NOT MODIFY
'
Public Type typTableCustomer
    lngCustomerID As Long
    strName As String
    strAddress1 As String
    strAddress2 As String
    strCity As String
    strState As String
    strZip As String
End Type
Public g_recCustomer As typTableCustomer
Public g_objRecordsetCustomer As ADODB.Recordset

Public Type typTableOrder
    lngOrderID As Long
    lngCustomerID As Long
    strOrderNo As String
    datOrder As Date
    strSalesmanID As String
    blnShipped As Boolean
End Type
Public g_recOrder As typTableOrder
Public g_objRecordsetOrder As ADODB.Recordset

Public Type typTableOrderDetail
    lngOrderDetailID As Long
    lngOrderID As Long
    lngCustomerID As Long
    strItemID As String
    sngQuantity As Single
End Type
Public g_recOrderDetail As typTableOrderDetail
Public g_objRecordsetOrderDetail As ADODB.Recordset

Public Type typTableSystem
    blnOption1 As Boolean
    blnOption2 As Boolean
    blnOption3 As Boolean
    blnOption4 As Boolean
    blnOption5 As Boolean
    blnOption6 As Boolean
End Type
Public g_recSystem As typTableSystem
Public g_objRecordsetSystem As ADODB.Recordset

Public Type typTableVersion
    strSystemName As String
    strVersion As String
    intSchemaVersion As Integer
End Type
Public g_recVersion As typTableVersion
Public g_objRecordsetVersion As ADODB.Recordset

'
'SQL Public Connection & Transaction Routines   Automatically Generated  -  DO NOT MODIFY
'

Public Sub SQLCloseConnection()
'
'   Closes a SQL Connection
'
    On Error Resume Next
    With g_recConnection
        .intOperationType = enumSQLOperationType.ConnectOperation
        .objConnection.Close
        Set .objConnection = Nothing
    End With
    Err.Clear
End Sub

Public Function SQLCommand(strSQLStatement As String, Optional strParameters As String = "", Optional blnStoredProcedure As Boolean = False, Optional intOperationType As Integer) As Boolean
'
'   Executes an SQL command for the connection
'
'   Parameters:
'       strSQLStatement - The SQL statement to be run
'       strParameters - an optional parameter string, if parameters will be passed
'       blnStoredProcedure - an optional flag indicating that the command will run a stored procedure
'       intOperationType - an optional value indicating the type of operation (from enumSQLOperationType)
'   
    SQLCommand = SQLExecute(g_recConnection, "Connection ", strSQLStatement, strParameters, blnStoredProcedure, intOperationType)
End Function

Public Function SQLOpenConnection(strDSN As String, Optional strConnectString As String, Optional strServer As String, Optional strDatabase As String, Optional strDatabasePassword As String, Optional strUserName As String, Optional strPassword As String, Optional strDatabaseSystem As String = "", Optional strRecordSetParams As String = "", Optional intConnectTimeout As Integer = 15, Optional intCommandTimeout As Integer = 15, Optional blnCursorClient As Boolean = False) As Boolean
'
'   Opens a SQL Connection
'
'   Parameters:
'       strDSN - the DSN string if the connection is being opened with a Data Source Name (may be blank)
'       strConnectString - the connect string used to open the connection (may be blank and may include replacement string values - see the SQLOpen routine)
'       strServer - optional name of the database server
'       strDatabase - optional name of the database
'       strDatabasePassword - optional database password
'       strUserName - optional username
'       strPassword - optional username password
'       strDatabaseSystem - optional database system being accessed (from enumSQLDatabase enumerations)
'       strRecordSetParams - optional default recordset parameters separated by spaces (e.g., "forwardonly" - from enumSQLRecordset enumerations)
'       intConnectTimeout - optional connect timeout value
'       intCommandTimeout - optional command timeout value
'       blnCursorClient - optional database cursor location flag (default is server-side cursor)
'
    On Error Resume Next
    With g_recConnection
        .intOperationType = enumSQLOperationType.ConnectOperation
        .strConnectionName = ""
    End With
    SQLOpenConnection = SQLOpen(g_recConnection, strDSN, strConnectString, strServer, strDatabase, strDatabasePassword, strUserName, strPassword, strDatabaseSystem, strRecordSetParams, intConnectTimeout, intCommandTimeout, blnCursorClient)
End Function

'
'SQL Public Data Access Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLCloseRecordSet(objRecordset As ADODB.Recordset)
'
'   Closes a recordset
'
'   Parameters:
'       objRecordset - the recordset object to be used
'
    On Error Resume Next
    objRecordset.Close
    Set objRecordset = Nothing
    Err.Clear
End Sub

Public Function SQLConnect(recConnection As typSQLConnection, Optional blnInitialConnect As Boolean = False, Optional blnSkipErrorHandling As Boolean = False) As Boolean
'
'   Performs the actual SQL connect/reconnect logic
'
'   Parameters:
'       recConnection - the connection record to be used
'       blnInitialConnect - an optional flag indicating if this is the initial or a reconnect operation
'       blnSkipErrorHandling - an optional flag indicating that error displays will be suppressed for this connect attempt
'
    Dim blnMouseHourglass As Boolean
    Dim blnResult As Boolean
    Dim intRetryCount As Integer
    Dim strConnection As String

    On Error Resume Next
    With recConnection
        If Screen.Mousepointer <> vbHourglass Then
            blnMouseHourglass = True
            Screen.Mousepointer = vbHourglass
        End If
        .objConnection.Close
        Set .objConnection = Nothing
        Err.Clear
        Set .objConnection = New ADODB.Connection
        .objConnection.CommandTimeout = .intCommandTimeout
        .objConnection.ConnectionTimeout = .intConnectTimeout
        .objConnection.Errors.Clear
        .objConnection.Open .strConnectString
        .intOperationType = enumSQLOperationType.ConnectOperation
        If Err.Number = 0 Then
            Set .objCommand = Nothing
            Set .objCommand = New ADODB.Command
            Set .objCommand.ActiveConnection = .objConnection
            Err.Clear
            blnResult = True
        ElseIf Not blnSkipErrorHandling Then
            Call SQLErrorProcess(recConnection, IIf(blnInitialConnect, "SQLConnectInitial", "SQLConnect"), "", .strConnectErrorString)
        End If
        If blnMouseHourglass Then
            Screen.Mousepointer = vbDefault
        End If
    End With
    SQLConnect = blnResult
End Function

Public Function SQLCreateTable(recConnection As typSQLConnection, strTableNameInternal As String, strTableNameExternal As String, strSQLStatement As String) As Boolean
'
'   Creates a table
'
'   Parameters:
'       recConnection - the connection record to be used
'       strTableNameInternal - the internal name of the SQL table
'       strTableNameExternal - the external name of the SQL table
'       strSQLStatement - the SQL create statement
'
    Dim intX As Integer
    Dim strLine() As String
    Dim strStatement As String

    On Error Resume Next
    If strSQLStatement <> "" Then
        strLine = Split(strSQLStatement, "|")
        strStatement = "CREATE TABLE " & strTableNameExternal & " (" & strLine(0) & ");"
        If SQLExecute(recConnection, strTableNameInternal, strStatement) Then
            For intX = 1 To UBound(strLine)
                If Left$(strLine(intX), 1) = "*" Then
                    strStatement = "CREATE UNIQUE INDEX "
                    strLine(intX) = Mid$(strLine(intX), 2)
                Else
                    strStatement = "CREATE INDEX "
                End If
                strStatement = strStatement & strLine(intX) & " ON " & strTableNameExternal & " (" & strLine(intX) & ");"
                Call SQLExecute(recConnection, strTableNameInternal, strStatement, , , enumSQLOperationType.CriticalWriteOperation)
            Next
            SQLCreateTable = True
        End If
    End If
    Err.Clear
End Function

Public Function SQLDeleteFrom(recConnection As typSQLConnection, strTableNameInternal As String, strTableNameExternal As String, strWhereClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
'
'   Deletes a record without using a recordset
'
'   Parameters:
'       recConnection - the connection record to be used
'       strTableNameInternal - the internal name of the SQL table
'       strTableNameExternal - the external name of the SQL table
'       strWhereClause - the SQL statment indicating the records to be deleted
'       blnCriticalWrite - an optional flag indicating that is a critical write operation
'
    Dim blnResult As Boolean
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        .intOperationType = IIf(blnCriticalWrite, enumSQLOperationType.CriticalWriteOperation, enumSQLOperationType.WriteOperation)
        strStatement = "DELETE FROM " & .strServerPrefix & strTableNameExternal
        If strWhereClause <> "" And LCase$(Left$(Trim$(strWhereClause), 5)) <> "where" Then
            strStatement = strStatement & " WHERE"
        End If
        strStatement = strStatement & " " & strWhereClause
        Do
            .objCommand.CommandType = adCmdText
            .objCommand.CommandText = strStatement
            .objCommand.Execute
            If Err.Number = 0 Then
                blnResult = True
                Exit Do
            ElseIf SQLErrorProcess(recConnection, "SQLDeleteFrom", strTableNameInternal, strStatement, "", True) Then
                Exit Do
            End If
        Loop
    End With
    SQLDeleteFrom = blnResult
End Function

Public Function SQLExecute(recConnection As typSQLConnection, strTableNameInternal As String, strSQLStatement As String, Optional strParameters As String = "", Optional blnStoredProcedure As Boolean = False, Optional intOperationType As Integer = 0) As Boolean
'
'   Executes an SQL command
'
'   Parameters:
'       recConnection - the connection record to be used
'       strTableNameInternal - the internal name of the SQL table being used
'       strSQLStatement - The SQL statement to be run
'       strParameters - an optional parameter string, if parameters will be passed
'       blnStoredProcedure - an optional flag indicating that the command will run a stored procedure
'       intOperationType - an optional value indicating the type of operation (from enumSQLOperationType)
'   
    Dim blnResult As Boolean

    On Error Resume Next
    With recConnection
        .objConnection.Errors.Clear
        .intOperationType = intOperationType
    End With
    With recConnection.objCommand
        If blnStoredProcedure Then
            .CommandType = adCmdStoredProc
        Else
            .CommandType = adCmdText
        End If
        .CommandText = strSQLStatement
        Do
            If strParameters <> "" Then
                .Execute , strParameters
            Else
                .Execute
            End If
            If Err.Number = 0 Then
                blnResult = True
                Exit Do
            ElseIf SQLErrorProcess(recConnection, "SQLExecute", strTableNameInternal, strSQLStatement, "", True) Then
                Exit Do
            End If
        Loop
    End With
    SQLExecute = blnResult
End Function

Public Function SQLGetErrorCount() As Long
'
'   Returns the SQL error count
'
    SQLGetErrorCount = m_lngErrorCount
End Function

Public Function SQLGetErrorLast() As String
'
'   Returns the error message on the last function called (if any)
'
    SQLGetErrorLast = m_strError
End Function

Public Function SQLGetErrorLogFileName() As String
'
'   Returns the name of the SQL error log file
'
    SQLGetErrorLogFileName = m_strErrorLogFile
End Function

Public Function SQLInsert(recConnection As typSQLConnection, strTableNameInternal As String, strTableNameExternal As String, strInsertClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
'
'   Inserts a new record into a table without using a recordset
'
'   Parameters:
'       recConnection - the connection record to be used
'       strTableNameInternal - the internal name of the SQL table
'       strTableNameExternal - the external name of the SQL table
'       strInsertClause - the SQL statment describing the insert processing
'       blnCriticalWrite - an optional flag indicating that is a critical write operation
'
    Dim blnResult As Boolean
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        .intOperationType = IIf(blnCriticalWrite, enumSQLOperationType.CriticalWriteOperation, enumSQLOperationType.WriteOperation)
        strStatement = "INSERT INTO " & .strServerPrefix & strTableNameExternal & " " & strInsertClause
        Do
            .objCommand.CommandType = adCmdText
            .objCommand.CommandText = strStatement
            .objCommand.Execute
            If Err.Number = 0 Then
                blnResult = True
                Exit Do
            ElseIf SQLErrorProcess(recConnection, "SQLInsert", strTableNameInternal, strStatement, "", True) Then
                Exit Do
            End If
        Loop
    End With
    SQLInsert = blnResult
End Function

Public Function SQLNext(objRecordset As ADODB.Recordset) As Boolean
'
'   Reads the next record in a recordset
'
'   Parameters:
'       objRecordset - the recordset object to be used
'
    On Error Resume Next
    objRecordset.MoveNext
    If Err.Number <> 0 Or objRecordset.EOF Then
        Err.Clear
    Else
        SQLNext = True
    End If
End Function

Public Function SQLSelect(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strTableNameExternal As String, strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
'
'   Performs a query and generates a recordset
'
'   Parameters:
'       recConnection - the connection record to be used
'       objRecordset - the recordset object to be used
'       strTableNameInternal - the internal name of the SQL table
'       strTableNameExternal - the external name of the SQL table
'       strWhereClause - the SQL statement indicating the WHERE and ORDER BY strings
'       lngTopCount - an optional value used to return only the first "n" results in the record set
'       strRecordSetParams - optional recordset parameters for this SELECT separated by spaces (e.g., "readonly" - from the enumSQLRecordset enumerations).  If blank, the default connection value is used
'
    Dim blnResult As Boolean
    Dim intX As Integer
    Dim lngCursorLocation As Long
    Dim lngCursorType As Long
    Dim lngLockType As Long
    Dim lngOption As Long
    Dim strArray() As String
    Dim strParameters As String
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        .intOperationType = enumSQLOperationType.ReadOperation
        lngCursorLocation = .lngCursorLocation
        lngCursorType = .lngCursorDefault
        lngLockType = adLockOptimistic
        lngOption = adCmdText
        strParameters = strRecordSetParams
        If strParameters = "" Then
            strParameters = .strRecordSetParams
        End If
        If strParameters <> "" Then
            strArray = Split(LCase$(strParameters), " ")
            For intX = 0 To UBound(strArray)
                Select Case Trim$(strArray(intX))
                    Case Is = enumSQLRecordsetForwardOnly
                        lngCursorType = adOpenForwardOnly
                    Case Is = enumSQLRecordsetKeyset
                        lngCursorType = adOpenKeyset
                    Case Is = enumSQLRecordsetDynamic
                        lngCursorType = adOpenDynamic
                    Case Is = enumSQLRecordsetStatic
                        lngCursorType = adOpenStatic
                    Case Is = enumSQLRecordsetReadOnly
                        lngLockType = adLockReadOnly
                    Case Is = enumSQLRecordsetPessimistic
                        lngLockType = adLockPessimistic
                    Case Is = enumSQLRecordsetOptimistic
                        lngLockType = adLockOptimistic
                    Case Is = enumSQLRecordsetBatchOptimistic
                        lngLockType = adLockBatchOptimistic
                    Case Is = enumSQLRecordsetStoredProc
                        lngOption = adCmdStoredProc
                    Case Is = enumSQLRecordsetTable
                        lngOption = adCmdTable
                End Select
            Next
        End If
        strStatement = "SELECT"
        If lngTopCount > 0 Then
            strStatement = strStatement & " TOP " & CStr(lngTopCount)
        End If
        strStatement = strStatement & " * FROM " & .strServerPrefix & strTableNameExternal
        If strWhereClause <> "" Then
            If LCase$(Left$(Trim$(strWhereClause), 5)) <> "where" And LCase$(Left$(Trim$(strWhereClause), 8)) <> "order by" Then
                strStatement = strStatement & " WHERE"
            End If
        End If
        strStatement = strStatement & " " & strWhereClause
        Do
            Set objRecordset = Nothing
            Set objRecordset = New ADODB.RecordSet
            objRecordset.CursorLocation = lngCursorLocation
            objRecordset.Open strStatement, .objConnection, lngCursorType, lngLockType, lngOption
            If Err.Number = 0 Then
                blnResult = Not objRecordset.EOF
                Exit Do
            ElseIf SQLErrorProcess(recConnection, "SQLSelect", strTableNameInternal, strStatement, "", True) Then
                Exit Do
            End If
        Loop
    End With
    SQLSelect = blnResult
End Function

Public Sub SQLSetErrorHandling(blnErrorDisplay As Boolean, blnErrorLog As Boolean, Optional strLogPathName As String = "")
'
'   Sets the error handling characteristics
'
'   Parameters:
'       blnErrorDisplay - a flag turning error displays (message boxes) on or off (errors are stilled logged)
'       blnErrorLog - a flag indicating if errors should be logged
'       strLogPathName - the path to be used for the error file (a blank path retains the current value)
'
    On Error Resume Next
    m_blnErrorDisplay = blnErrorDisplay
    m_blnErrorLog = blnErrorLog
    If blnErrorLog Then
        If strLogPathName <> "" Then
            m_strErrorLogFile = strLogPathName & enumSQLErrorLogFile
        End If
    End If
End Sub

Public Function SQLUpdateInto(recConnection As typSQLConnection, strTableNameInternal As String, strTableNameExternal As String, strUpdateClause As String, strWhereClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
'
'   Updates a record without using a recordset
'
'   Parameters:
'       recConnection - the connection record to be used
'       strTableNameInternal - the internal name of the SQL table
'       strTableNameExternal - the external name of the SQL table
'       strUpdateClause - the SQL SET statement
'       strWhereClause - the SQL WHERE statement indicating the records to be updated
'       blnCriticalWrite - an optional flag indicating that is a critical write operation
'
    Dim blnResult As Boolean
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        .intOperationType = IIf(blnCriticalWrite, enumSQLOperationType.CriticalWriteOperation, enumSQLOperationType.WriteOperation)
        strStatement = "UPDATE " & .strServerPrefix & strTableNameExternal & " " & strUpdateClause
        If strWhereClause <> "" And LCase$(Left$(Trim$(strWhereClause), 5)) <> "where" Then
            strStatement = strStatement & " WHERE"
        End If
        strStatement = strStatement & " " & strWhereClause
        Do
            .objCommand.CommandType = adCmdText
            .objCommand.CommandText = strStatement
            .objCommand.Execute
            If Err.Number = 0 Then
                blnResult = True
                Exit Do
            ElseIf SQLErrorProcess(recConnection, "SQLUpdateInto", strTableNameInternal, strStatement, "", True) Then
                Exit Do
            End If
        Loop
    End With
    SQLUpdateInto = blnResult
End Function


'
'SQL Private Routines   Automatically Generated  -  DO NOT MODIFY
'

Private Function SQLErrorProcess(recConnection As typSQLConnection, strOperation As string, strTableName As String, strStatement As String, Optional strMsg As String = "", Optional blnRetryConnectError As Boolean) As Boolean
'
'   Handles SQL errors and checks to see if function should be retried
'
    Const lngRetrySleepValue As Long = 1000

    Static blnErrorBeingDisplayed As Boolean
    Static intConnectFailureRetryCount As Integer

    Dim blnConnectFailure As Boolean
    Dim intFileNo As Integer
    Dim lngErrorNo As Long
    Dim lngErrorNoActual As Long
    Dim strErrorMsg As String
    Dim strErrorString As String

    If Err.Number <> 0 Then
        lngErrorNo = Err.Number
        strErrorString = Err.Description & "  (" & CStr(Err.Number) & ")"
    End If
    Err.Clear
    Screen.Mousepointer = vbDefault
    On Error Resume Next
    With recConnection.objConnection.Errors
        If .Count >= 1 Then
            lngErrorNoActual = .Item(0).NativeError
        Else
            lngErrorNoActual = 0
        End If
        .Clear
    End With
    Err.Clear
    If (strOperation = "SQLInsert" Or strOperation = "SQLUpdateInto") _
        And lngErrorNoActual = enumSQLErrorUpdateDupKey Then
        'native error codes for dup key inserts and updates from all database systems are needed here
        SQLErrorProcess = True
        Exit Function
    End If
    If lngErrorNo = enumSQLErrorConnectFail _
        Or lngErrorNo = enumSQLErrorConnectFailAccess _
        Or lngErrorNo = enumSQLErrorConnectFailPostgres _
        Or lngErrorNo = enumSQLErrorNetwork Then
        If blnRetryConnectError And intConnectFailureRetryCount <= enumSQLErrorMaxRetries Then
            intConnectFailureRetryCount = intConnectFailureRetryCount + 1
            Call Sleep(lngRetrySleepValue)
            If SQLConnect(recConnection, False, True) Then
                Exit Function
            End If
        End If
        blnConnectFailure = True
    End If
    intConnectFailureRetryCount = 0
    If strMsg <> "" Then
        strErrorString = strErrorString & IIf(strErrorString = "", "", vbCrLf) & strMsg
    End If
    m_lngErrorCount = m_lngErrorCount + 1
    With recConnection
        m_strError = strErrorString
        strErrorMsg = "An error occurred during the function '" & strOperation & "'" & _
            IIf(.strConnectionName = "", "", " on the connection '" & .strConnectionName & "'") & _
            IIf(strTableName = "", "", " for the table '" & strTableName & "'") & vbCrLf & vbCrLf & _
            "The specific error was:  " & strErrorString
        If strStatement <> "" Then
            strErrorMsg = strErrorMsg & vbCrLf & vbCrLf & "The SQL Statement was:" & vbCrLf & strStatement
        End If
    End With
    If m_blnErrorLog Then
        intFileNo = FreeFile
        Open m_strErrorLogFile For Append As #intFileNo
        With App
            Print #intFileNo, Format$(Now, "mm/dd/yyyy  hh:nn:ss"); "    "; _
                .FileDescription; "  "; CStr(.Major); "."; CStr(.Minor); "."; CStr(.Revision); vbCrLf; vbCrLf; _
                strErrorMsg
        End With
        Print #intFileNo, strErrorMsg
        Print #intFileNo, String$(80, "=")
        Close #intFileNo
        Err.Clear
    End If
    If Not blnErrorBeingDisplayed Then
        blnErrorBeingDisplayed = True
        If blnConnectFailure And strOperation <> "SQLConnectInitial" Then
            #If enumSQLCallback Then
            Call Callback_SQLConnectFailure(recConnection)
            #End If
        ElseIf m_blnErrorDisplay And Left$(strOperation, 10) <> "SQLConnect" Then
            #If enumSQLCallback Then
            Call Callback_MsgboxPreProcess
            #End If
            MsgBox "SQL Database Error" & vbCrLf & vbCrLf & strErrorMsg, vbExclamation
        End If
        blnErrorBeingDisplayed = False
    End If
    SQLErrorProcess = True
End Function


Private Function SQLFieldBuildBegin(recConnection As typSQLConnection, intBuildType As Integer) As Integer
'
'   Start an Insert/Update string build operation
'
    Dim IntX As Integer

    On Error Resume Next
    For intX = 0 to UBound(m_recBuild)
        If Not m_recBuild(intX).blnActive Then
            Exit For
        End If
    Next
    If intX > UBound(m_recBuild) Then
        Call SQLErrorProcess(recConnection, "SQLBuild", "", "", "SQL Build Array Overflow")
	intX = -1
    Else
        With m_recBuild(intX)
            .blnActive = True
            .intBuildType = intBuildType
            .strBuildField = ""
            .strBuildValue = ""
            .strDatabaseSystem = recConnection.strDatabaseSystem
            .strEscapeQuote = recConnection.strEscapeQuote
        End With
    End If
    SQLFieldBuildBegin = intX
    Err.Clear
End Function

Private Function SQLFieldBuildEnd(intBuildIndex As Integer) As String
'
'   End an Insert/Update string build operation
'
    On Error Resume Next
    With m_recBuild(intBuildIndex)
        Select Case .intBuildType
            Case Is = enumSQLBuildType.CreateTable
                SQLFieldBuildEnd = .strBuildValue & IIf(.strBuildField = "", "", "|" & .strBuildField)
            Case Is = enumSQLBuildType.InsertTable
                SQLFieldBuildEnd = " (" & .strBuildField & ") VALUES (" & .strBuildValue & ")"
            Case Is = enumSQLBuildType.UpdateTable
                SQLFieldBuildEnd = " SET " & .strBuildValue
        End Select
        .blnActive = False
        .strBuildField = ""
        .strBuildValue = ""
    End With
    Err.Clear
End Function

Private Sub SQLFieldBuildValue(intBuildIndex As Integer, strField As String, strValue As String, Optional strIndex As String = "")
'
'   Emit an CreateTable/Insert/Update field/value
'
    On Error Resume Next
    With m_recBuild(intBuildIndex)
        Select Case .intBuildType
            Case Is = enumSQLBuildType.CreateTable
                If strIndex <> "" Then
                    .strBuildField = .strBuildField & IIf(.strBuildField = "", "", "|") & strIndex
                End If
                .strBuildValue = .strBuildValue & IIf(.strBuildValue = "", "", ", ") & strField & " " & strValue
            Case Is = enumSQLBuildType.InsertTable
                .strBuildField = .strBuildField & IIf(.strBuildField = "", "", ", ") & strField
                .strBuildValue = .strBuildValue & IIf(.strBuildValue = "", "", ", ") & strValue
            Case Is = enumSQLBuildType.UpdateTable
                .strBuildValue = .strBuildValue & IIf(.strBuildValue = "", "", ", ") & strField & " = " & strValue
        End Select
    End With
    Err.Clear
End Sub

Private Sub SQLFieldBuildBoolean(intBuildIndex As Integer, strField As String, blnValue As Boolean)
'
'   Emit a boolean value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, IIf(blnValue, "true", "false"))
End Sub

Private Sub SQLFieldBuildCreateTable(intBuildIndex As Integer, strField As String, strFieldType As String, lngLength As Long, blnAutoKey As Boolean, blnPrimary As Boolean, blnIndex As Boolean, blnUnique As Boolean)
'
'   Emit a file type for create table statement - customized for the database system
'
    Dim strIndex As String
    Dim strType As String

    On Error Resume Next
    Select Case LCase$(strFieldType)
        Case Is = "boolean"
            Select Case m_recBuild(intBuildIndex).strDatabaseSystem
                Case Is = enumSQLDatabaseAccess
                    strType = "YESNO"
                Case Else
                    strType = "BOOLEAN"
            End Select
        Case Is = "currency"
            Select Case m_recBuild(intBuildIndex).strDatabaseSystem
                Case Is = enumSQLDatabaseAccess
                    strType = "CURRENCY"
                Case Else
                    strType = "DOUBLE"
            End Select
        Case Is = "date"
            strType = "DATE"
        Case Is = "double"
            strType = "DOUBLE"
        Case Is = "integer"
            Select Case m_recBuild(intBuildIndex).strDatabaseSystem
                Case Is = enumSQLDatabaseAccess
                    strType = "INTEGER"
                Case Else
                    strType = "SHORT"
            End Select
        Case Is = "long"
            Select Case m_recBuild(intBuildIndex).strDatabaseSystem
                Case Is = enumSQLDatabaseAccess
                    If blnAutoKey Then
                        strType = "AUTOINCREMENT"
                    Else
                        strType = "LONG"
                    End If
                Case Else
                    strType = "INTEGER"
            End Select
        Case Is = "memo"
            strType = "MEMO"
        Case Is = "single"
            strType = "SINGLE"
        Case Is = "string"
            strType = "TEXT" & IIf(lngLength > 0, "(" & CStr(lngLength) & ")", "")
    End Select
    strIndex = ""
    If blnPrimary Then
        strType = strType & " PRIMARY KEY"
    ElseIf blnIndex Then
	strIndex = strField
        If blnUnique Then
            strIndex = "*" & strIndex
        End If
    End If
    Err.Clear
    Call SQLFieldBuildValue(intBuildIndex, strField, strType, strIndex)
End Sub

Private Sub SQLFieldBuildCurrency(intBuildIndex As Integer, strField As String, curValue As Currency)
'
'   Emit a currency value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, CStr(curValue))
End Sub

Private Sub SQLFieldBuildDate(intBuildIndex As Integer, strField As String, datValue As Date)
'
'   Emit a date value
'
    Dim strValue As String

    On Error Resume Next
    Select Case m_recBuild(intBuildIndex).strDatabaseSystem
        Case Is = enumSQLDatabaseOracle
            'Requires Oracle date/time field
            strValue = "to_date('" & Format$(datValue, "yyyy/mm/dd:hh:nn:ss am/pm") & "', 'yyyy/mm/dd:hh:mi:ss am')"
        Case Is = enumSQLDatabaseMySql
            'Requires MySql DATETIME field (i.e., not DATE, TIME, TIMESTAMP, or YEAR which have no analogs in the generator)
            strValue = "'" & Format$(datValue, "yyyy-mm-dd hh:nn:ss") & "'"
        Case Else
            If datValue = 0 Then
                strValue = "null"
            Else
                strValue = "'" & Format$(datValue, "mmm-dd-yyyy hh:nn:ss") & "'"
            End If
    End Select
    Err.Clear
    Call SQLFieldBuildValue(intBuildIndex, strField, strValue)
End Sub

Private Sub SQLFieldBuildDouble(intBuildIndex As Integer, strField As String, dblValue As Double)
'
'   Emit a double value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, CStr(dblValue))
End Sub

Private Sub SQLFieldBuildInteger(intBuildIndex As Integer, strField As String, intValue As Integer)
'
'   Emit an integer value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, CStr(intValue))
End Sub

Private Sub SQLFieldBuildLong(intBuildIndex As Integer, strField As String, lngValue As Long)
'
'   Emit a long value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, CStr(lngValue))
End Sub

Private Sub SQLFieldBuildSingle(intBuildIndex As Integer, strField As String, sngValue As Single)
'
'   Emit a single value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, CStr(sngValue))
End Sub

Private Sub SQLFieldBuildString(intBuildIndex As Integer, strField As String, strValue As String, lngLength As Long)
'
'   Emit a string value
'
    Call SQLFieldBuildValue(intBuildIndex, strField, "'" & SQLFieldEmitString(strValue, lngLength, m_recBuild(intBuildIndex).strEscapeQuote) & "'")
End Sub


Private Function SQLFieldEmitDate(datValue As Date, strDatabaseSystem As String) As String
'
'   Generate a date value string suitable for a query
'
    On Error Resume Next
    Select Case strDatabaseSystem
        Case Is = enumSQLDatabaseAccess, enumSQLDatabaseSqlServer
            SQLFieldEmitDate = "#" & Format$(datValue, "mm/dd/yyyy") & "#"
        Case Is = enumSQLDatabaseMySql
            SQLFieldEmitDate = "'" & Format$(datValue, "yyyy-mm-dd") & "'"
        Case Else
            SQLFieldEmitDate = "'" & Format$(datValue, "mm/dd/yyyy") & "'"
    End Select
End Function

Private Function SQLFieldEmitString(strString As String, lngLength As Long, strEscapeQuote As String) As String
'
'   Generate a string value suitable for database insertion, update or query purposes
'
    Dim lngX As Long
    Dim strData As String

    On Error Resume Next
    If lngLength > 0 And Len(strData) > lngLength Then
        strData = Left$(strData, lngLength)
    End If
    If strEscapeQuote = "" Then
        strData = Replace(strString, "'", "")
    Else
        strData = strString
        If strEscapeQuote <> "'" Then
            lngX = 1
            Do
                lngX = InStr(lngX, strData, strEscapeQuote)
                If lngX > 0 Then
                    strData = Left$(strData, lngX - 1) & strEscapeQuote & Mid$(strData, lngX)
                    lngX = lngX + 2
                End If
            Loop Until lngX = 0
        End If
        lngX = 1
        Do
            lngX = InStr(lngX, strData, "'")
            If lngX > 0 Then
                strData = Left$(strData, lngX - 1) & strEscapeQuote & Mid$(strData, lngX)
                lngX = lngX + 2
            End If
        Loop Until lngX = 0
    End If
    Err.Clear
    SQLFieldEmitString = strData
End Function


Private Function SQLFieldGetBoolean(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Boolean
'
'   Returns the value for a boolean field
'
    On Error Resume Next
    SQLFieldGetBoolean = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetBoolean", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetBoolean = False
    End If
End Function

Private Function SQLFieldGetCurrency(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Currency
'
'   Returns the value for a currency field
'
    On Error Resume Next
    SQLFieldGetCurrency = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetCurrency", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetCurrency = 0
    End If
End Function

Private Function SQLFieldGetDate(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Date
'
'   Returns the value for a date field
'
    On Error Resume Next
    SQLFieldGetDate = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetDate", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetDate = 0
    End If
End Function

Private Function SQLFieldGetDouble(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Double
'
'   Returns the value for a double field
'
    On Error Resume Next
    SQLFieldGetDouble = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetDouble", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetDouble = 0
    End If
End Function

Private Function SQLFieldGetInteger(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Integer
'
'   Returns the value for a integer field
'
    On Error Resume Next
    SQLFieldGetInteger = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetInteger", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetInteger = 0
    End If
End Function

Private Function SQLFieldGetLong(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Long
'
'   Returns the value for a long field
'
    On Error Resume Next
    SQLFieldGetLong = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetLong", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetLong = 0
    End If
End Function

Private Function SQLFieldGetSingle(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As Single
'
'   Returns the value for a single field
'
    On Error Resume Next
    SQLFieldGetSingle = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetSingle", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetSingle = 0
    End If
End Function

Private Function SQLFieldGetString(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, strTableNameInternal As String, strFieldName As String) As String
'
'   Returns the value for a string field
'
    On Error Resume Next
    SQLFieldGetString = objRecordset(strFieldName)
    If Err.Number <> 0 Then
        If Err.Number = enumSQLErrorNullValue Then
            Err.Clear
        Else
            Call SQLErrorProcess(recConnection, "SQLFieldGetString", strTableNameInternal, "Get " & strFieldName)
        End If
        SQLFieldGetString = ""
    End If
End Function


Private Function SQLGetCount(recConnection As typSQLConnection, strTableNameExternal As String, strWhereClause As String) As Long
'
'   Returns the count of records from a table
'
    Dim objGet As ADODB.RecordSet
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        strStatement = "SELECT COUNT(*) AS RecordCount FROM " & .strServerPrefix & strTableNameExternal
        If strWhereClause <> "" Then
            strStatement = strStatement & " WHERE " & strWhereClause
        End If
        Set objGet = New ADODB.RecordSet
        objGet.Open strStatement, .objConnection, adOpenForwardOnly, adLockReadOnly
    End With
    If Err.Number = 0 Then
        SQLGetCount = objGet("RecordCount")
        objGet.Close
    Else
        Err.Clear
    End If
    Set objGet = Nothing
End Function

Private Function SQLGetKeyValue(strKeyFieldName As String, recConnection As typSQLConnection, strTableNameExternal As String) As Long
'
'   Returns the new key value assigned after an insert on a table with an autonumber field
'
    Dim objGet As ADODB.RecordSet
    Dim strStatement As String

    On Error Resume Next
    With recConnection
        strStatement = "SELECT MAX(" & strKeyFieldName & ") AS MaxID FROM " & .strServerPrefix & strTableNameExternal
        Set objGet = New ADODB.RecordSet
        objGet.Open strStatement, .objConnection, adOpenForwardOnly, adLockReadOnly
    End With
    If Err.Number = 0 Then
        SQLGetKeyValue = objGet("MaxID")
        objGet.Close
    Else
        Err.Clear
    End If
    Set objGet = Nothing
End Function

Private Function SQLOpen(recConnection As typSQLConnection, strDSN As String, strConnectString As String, strServer As String, strDatabase As String, strDatabasePassword As String, strUserName As String, strPassword As String, strDatabaseSystem As String, strRecordSetParams As String, intConnectTimeout As Integer, intCommandTimeout As Integer, blnCursorClient As Boolean) As Boolean
'
'   Opens a SQL Connection
'
    Const enumDataSourceName As String = "DSN="
    Const enumDatabaseText As String = "<db>"
    Const enumDatabasePasswordText As String = "<dbpwd>"
    Const enumPasswordText As String = "<pwd>"
    Const enumServerText As String = "<srv>"
    Const enumUserNameText As String = "<uid>"

    Dim intX As Integer

    On Error Resume Next
    If m_strErrorLogFile = "" Then
        m_strErrorLogFile = App.Path & "\" & enumSQLErrorLogFile
    End If
    With recConnection
        m_blnErrorDisplay = True
        .intOperationType = enumSQLOperationType.ConnectOperation
        .intCommandTimeout = intCommandTimeout
        .intConnectTimeout = intConnectTimeout
        .lngCursorLocation = adUseServer
        .lngCursorDefault = adOpenKeyset
        .strServer = strServer
        .strDatabase = strDatabase
        .strDatabasePassword = strDatabasePassword
        .strUserName = strUserName
        .strPassword = strPassword
        If strDSN <> "" Then
            .blnDSN = True
            .strConnectString = enumDataSourceName & strDSN
        Else
            .blnDSN = False
            .strConnectString = Replace(LCase$(strConnectString), enumServerText, strServer, , , vbTextCompare)
            .strConnectString = Replace(LCase$(.strConnectString), enumDatabaseText, strDatabase, , , vbTextCompare)
            .strConnectString = Replace(LCase$(.strConnectString), enumUserNameText, strUserName, , , vbTextCompare)
            .strConnectErrorString = .strConnectString
            .strConnectString = Replace(LCase$(.strConnectString), enumDatabasePasswordText, strDatabasePassword, , , vbTextCompare)
            .strConnectString = Replace(LCase$(.strConnectString), enumPasswordText, strPassword, , , vbTextCompare)
        End If
        If blnCursorClient Then
            .lngCursorLocation = adUseClient
        End If
        If strDatabaseSystem = "" Then
            .strDatabaseSystem = enumSQLDatabaseSqlServer
        Else
            .strDatabaseSystem = LCase$(strDatabaseSystem)
        End If
        .strServerPrefix = ""
        .strRecordSetParams = strRecordSetParams
        .strEscapeQuote = ""
        Select Case .strDatabaseSystem
            Case Is = enumSQLDatabaseAccess
                .strEscapeQuote = "'"
                intX = InStrRev(strDatabase, "\")
                If intX > 0 Then
                    .strDatabase = Mid$(strDatabase, intX + 1)
                Else
                    .strDatabase = strDatabase
                End If
                .strFileName = strDatabase
            Case Is = enumSQLDatabaseMySql
                .strEscapeQuote = "\"
            Case Is = enumSQLDatabaseOracle
                .lngCursorDefault = adOpenStatic
                .strServerPrefix = strServer & "."
            Case Is = enumSQLDatabasePostgresQL
                .strEscapeQuote = "\"
        End Select
        Err.Clear
    End With
    SQLOpen = SQLConnect(recConnection, True)
End Function


'
'SQL Customer Table Public Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLClearCustomer(recBuffer As typTableCustomer)
    On Error Resume Next
    With recBuffer
        .lngCustomerID = 0
        .strName = ""
        .strAddress1 = ""
        .strAddress2 = ""
        .strCity = ""
        .strState = ""
        .strZip = ""
    End With
End Sub

Public Sub SQLCloseCustomer()
    Call SQLCloseRecordSet(g_objRecordsetCustomer)
End Sub

Public Function SQLDeleteFromCustomer(strWhereClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
    SQLDeleteFromCustomer = SQLDeleteFrom(g_recConnection, "Customer", "Customer", strWhereClause, blnCriticalWrite)
End Function

Public Sub SQLFetchCustomer(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, recBuffer As typTableCustomer)
    On Error Resume Next
    recConnection.intOperationType = enumSQLOperationType.ReadOperation
    With recBuffer
        .lngCustomerID = SQLFieldGetLong(recConnection, objRecordset, "Customer", "lngCustomerID")
        .strName = SQLFieldGetString(recConnection, objRecordset, "Customer", "strName")
        .strAddress1 = SQLFieldGetString(recConnection, objRecordset, "Customer", "strAddress1")
        .strAddress2 = SQLFieldGetString(recConnection, objRecordset, "Customer", "strAddress2")
        .strCity = SQLFieldGetString(recConnection, objRecordset, "Customer", "strCity")
        .strState = SQLFieldGetString(recConnection, objRecordset, "Customer", "strState")
        .strZip = SQLFieldGetString(recConnection, objRecordset, "Customer", "strZip")
    End With
End Sub

Public Function SQLInsertCustomer(recBuffer As typTableCustomer, Optional blnFetchKey As Boolean = False, Optional blnInsertAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildCustomer(g_recConnection, recBuffer, enumSQLBuildType.InsertTable, blnInsertAutoKeyField)
    If strSQLStatement <> "" Then
        If SQLInsert(g_recConnection, "Customer", "Customer", strSQLStatement, blnCriticalWrite) Then
            If blnFetchKey Then
                recBuffer.lngCustomerID = SQLGetKeyValue("lngCustomerID", g_recConnection, "Customer")
            End If
            SQLInsertCustomer = True
        End If
    End If
End Function

Public Function SQLNextCustomer() As Boolean
    On Error Resume Next
    If SQLNext(g_objRecordsetCustomer) Then
        Call SQLFetchCustomer(g_recConnection, g_objRecordsetCustomer, g_recCustomer)
        SQLNextCustomer = True
    End If
End Function

Public Function SQLRecordCountCustomer(strWhereClause As String) As Long
    SQLRecordCountCustomer = SQLGetCount(g_recConnection, "Customer", strWhereClause)
End Function

Public Function SQLSelectCustomer(strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    On Error Resume Next
    If SQLSelect(g_recConnection, g_objRecordsetCustomer, "Customer", "Customer", strWhereClause, lngTopCount, strRecordSetParams) Then
        Call SQLFetchCustomer(g_recConnection, g_objRecordsetCustomer, g_recCustomer)
        SQLSelectCustomer = True
    End If
End Function

Public Function SQLStatementBuildCustomer(recConnection As typSQLConnection, recBuffer As typTableCustomer, intBuildType As Integer, Optional blnAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As String
    Dim intBuildIndex As Integer

    On Error Resume Next
    intBuildIndex = SQLFieldBuildBegin(recConnection, intBuildType)
    If intBuildIndex >= 0 Then
        With recBuffer
            If blnAutoKeyField Then
                Call SQLFieldBuildLong(intBuildIndex, "lngCustomerID", .lngCustomerID)
            End If
            Call SQLFieldBuildString(intBuildIndex, "strName", .strName, 50)
            Call SQLFieldBuildString(intBuildIndex, "strAddress1", .strAddress1, 30)
            Call SQLFieldBuildString(intBuildIndex, "strAddress2", .strAddress2, 30)
            Call SQLFieldBuildString(intBuildIndex, "strCity", .strCity, 30)
            Call SQLFieldBuildString(intBuildIndex, "strState", .strState, 2)
            Call SQLFieldBuildString(intBuildIndex, "strZip", .strZip, 10)
        End With
        SQLStatementBuildCustomer = SQLFieldBuildEnd(intBuildIndex)
    End If
End Function

Public Function SQLUpdateIntoCustomer(recBuffer As typTableCustomer, strWhereClause As String, Optional blnUpdateAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildCustomer(g_recConnection, recBuffer, enumSQLBuildType.UpdateTable, blnUpdateAutoKeyField)
    If strSQLStatement <> "" Then
        SQLUpdateIntoCustomer = SQLUpdateInto(g_recConnection, "Customer", "Customer", strSQLStatement, strWhereClause, blnCriticalWrite)
    End If
End Function


'
'SQL Order Table Public Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLClearOrder(recBuffer As typTableOrder)
    On Error Resume Next
    With recBuffer
        .lngOrderID = 0
        .lngCustomerID = 0
        .strOrderNo = ""
        .datOrder = 0
        .strSalesmanID = ""
        .blnShipped = False
    End With
End Sub

Public Sub SQLCloseOrder()
    Call SQLCloseRecordSet(g_objRecordsetOrder)
End Sub

Public Function SQLDeleteFromOrder(strWhereClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
    SQLDeleteFromOrder = SQLDeleteFrom(g_recConnection, "Order", "Order", strWhereClause, blnCriticalWrite)
End Function

Public Sub SQLFetchOrder(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, recBuffer As typTableOrder)
    On Error Resume Next
    recConnection.intOperationType = enumSQLOperationType.ReadOperation
    With recBuffer
        .lngOrderID = SQLFieldGetLong(recConnection, objRecordset, "Order", "lngOrderID")
        .lngCustomerID = SQLFieldGetLong(recConnection, objRecordset, "Order", "lngCustomerID")
        .strOrderNo = SQLFieldGetString(recConnection, objRecordset, "Order", "strOrderNo")
        .datOrder = SQLFieldGetDate(recConnection, objRecordset, "Order", "datOrder")
        .strSalesmanID = SQLFieldGetString(recConnection, objRecordset, "Order", "strSalesmanID")
        .blnShipped = SQLFieldGetBoolean(recConnection, objRecordset, "Order", "blnShipped")
    End With
End Sub

Public Function SQLInsertOrder(recBuffer As typTableOrder, Optional blnFetchKey As Boolean = False, Optional blnInsertAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildOrder(g_recConnection, recBuffer, enumSQLBuildType.InsertTable, blnInsertAutoKeyField)
    If strSQLStatement <> "" Then
        If SQLInsert(g_recConnection, "Order", "Order", strSQLStatement, blnCriticalWrite) Then
            If blnFetchKey Then
                recBuffer.lngOrderID = SQLGetKeyValue("lngOrderID", g_recConnection, "Order")
            End If
            SQLInsertOrder = True
        End If
    End If
End Function

Public Function SQLNextOrder() As Boolean
    On Error Resume Next
    If SQLNext(g_objRecordsetOrder) Then
        Call SQLFetchOrder(g_recConnection, g_objRecordsetOrder, g_recOrder)
        SQLNextOrder = True
    End If
End Function

Public Function SQLRecordCountOrder(strWhereClause As String) As Long
    SQLRecordCountOrder = SQLGetCount(g_recConnection, "Order", strWhereClause)
End Function

Public Function SQLSelectOrder(strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    On Error Resume Next
    If SQLSelect(g_recConnection, g_objRecordsetOrder, "Order", "Order", strWhereClause, lngTopCount, strRecordSetParams) Then
        Call SQLFetchOrder(g_recConnection, g_objRecordsetOrder, g_recOrder)
        SQLSelectOrder = True
    End If
End Function

Public Function SQLStatementBuildOrder(recConnection As typSQLConnection, recBuffer As typTableOrder, intBuildType As Integer, Optional blnAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As String
    Dim intBuildIndex As Integer

    On Error Resume Next
    intBuildIndex = SQLFieldBuildBegin(recConnection, intBuildType)
    If intBuildIndex >= 0 Then
        With recBuffer
            If blnAutoKeyField Then
                Call SQLFieldBuildLong(intBuildIndex, "lngOrderID", .lngOrderID)
            End If
            Call SQLFieldBuildLong(intBuildIndex, "lngCustomerID", .lngCustomerID)
            Call SQLFieldBuildString(intBuildIndex, "strOrderNo", .strOrderNo, 20)
            Call SQLFieldBuildDate(intBuildIndex, "datOrder", .datOrder)
            Call SQLFieldBuildString(intBuildIndex, "strSalesmanID", .strSalesmanID, 10)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnShipped", .blnShipped)
        End With
        SQLStatementBuildOrder = SQLFieldBuildEnd(intBuildIndex)
    End If
End Function

Public Function SQLUpdateIntoOrder(recBuffer As typTableOrder, strWhereClause As String, Optional blnUpdateAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildOrder(g_recConnection, recBuffer, enumSQLBuildType.UpdateTable, blnUpdateAutoKeyField)
    If strSQLStatement <> "" Then
        SQLUpdateIntoOrder = SQLUpdateInto(g_recConnection, "Order", "Order", strSQLStatement, strWhereClause, blnCriticalWrite)
    End If
End Function


'
'SQL OrderDetail Table Public Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLClearOrderDetail(recBuffer As typTableOrderDetail)
    On Error Resume Next
    With recBuffer
        .lngOrderDetailID = 0
        .lngOrderID = 0
        .lngCustomerID = 0
        .strItemID = ""
        .sngQuantity = 0
    End With
End Sub

Public Sub SQLCloseOrderDetail()
    Call SQLCloseRecordSet(g_objRecordsetOrderDetail)
End Sub

Public Function SQLDeleteFromOrderDetail(strWhereClause As String, Optional blnCriticalWrite As Boolean = False) As Boolean
    SQLDeleteFromOrderDetail = SQLDeleteFrom(g_recConnection, "OrderDetail", "OrderDetail", strWhereClause, blnCriticalWrite)
End Function

Public Sub SQLFetchOrderDetail(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, recBuffer As typTableOrderDetail)
    On Error Resume Next
    recConnection.intOperationType = enumSQLOperationType.ReadOperation
    With recBuffer
        .lngOrderDetailID = SQLFieldGetLong(recConnection, objRecordset, "OrderDetail", "lngOrderDetailID")
        .lngOrderID = SQLFieldGetLong(recConnection, objRecordset, "OrderDetail", "lngOrderID")
        .lngCustomerID = SQLFieldGetLong(recConnection, objRecordset, "OrderDetail", "lngCustomerID")
        .strItemID = SQLFieldGetString(recConnection, objRecordset, "OrderDetail", "strItemID")
        .sngQuantity = SQLFieldGetSingle(recConnection, objRecordset, "OrderDetail", "sngQuantity")
    End With
End Sub

Public Function SQLInsertOrderDetail(recBuffer As typTableOrderDetail, Optional blnFetchKey As Boolean = False, Optional blnInsertAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildOrderDetail(g_recConnection, recBuffer, enumSQLBuildType.InsertTable, blnInsertAutoKeyField)
    If strSQLStatement <> "" Then
        If SQLInsert(g_recConnection, "OrderDetail", "OrderDetail", strSQLStatement, blnCriticalWrite) Then
            If blnFetchKey Then
                recBuffer.lngOrderDetailID = SQLGetKeyValue("lngOrderDetailID", g_recConnection, "OrderDetail")
            End If
            SQLInsertOrderDetail = True
        End If
    End If
End Function

Public Function SQLNextOrderDetail() As Boolean
    On Error Resume Next
    If SQLNext(g_objRecordsetOrderDetail) Then
        Call SQLFetchOrderDetail(g_recConnection, g_objRecordsetOrderDetail, g_recOrderDetail)
        SQLNextOrderDetail = True
    End If
End Function

Public Function SQLRecordCountOrderDetail(strWhereClause As String) As Long
    SQLRecordCountOrderDetail = SQLGetCount(g_recConnection, "OrderDetail", strWhereClause)
End Function

Public Function SQLSelectOrderDetail(strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    On Error Resume Next
    If SQLSelect(g_recConnection, g_objRecordsetOrderDetail, "OrderDetail", "OrderDetail", strWhereClause, lngTopCount, strRecordSetParams) Then
        Call SQLFetchOrderDetail(g_recConnection, g_objRecordsetOrderDetail, g_recOrderDetail)
        SQLSelectOrderDetail = True
    End If
End Function

Public Function SQLStatementBuildOrderDetail(recConnection As typSQLConnection, recBuffer As typTableOrderDetail, intBuildType As Integer, Optional blnAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As String
    Dim intBuildIndex As Integer

    On Error Resume Next
    intBuildIndex = SQLFieldBuildBegin(recConnection, intBuildType)
    If intBuildIndex >= 0 Then
        With recBuffer
            If blnAutoKeyField Then
                Call SQLFieldBuildLong(intBuildIndex, "lngOrderDetailID", .lngOrderDetailID)
            End If
            Call SQLFieldBuildLong(intBuildIndex, "lngOrderID", .lngOrderID)
            Call SQLFieldBuildLong(intBuildIndex, "lngCustomerID", .lngCustomerID)
            Call SQLFieldBuildString(intBuildIndex, "strItemID", .strItemID, 20)
            Call SQLFieldBuildSingle(intBuildIndex, "sngQuantity", .sngQuantity)
        End With
        SQLStatementBuildOrderDetail = SQLFieldBuildEnd(intBuildIndex)
    End If
End Function

Public Function SQLUpdateIntoOrderDetail(recBuffer As typTableOrderDetail, strWhereClause As String, Optional blnUpdateAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildOrderDetail(g_recConnection, recBuffer, enumSQLBuildType.UpdateTable, blnUpdateAutoKeyField)
    If strSQLStatement <> "" Then
        SQLUpdateIntoOrderDetail = SQLUpdateInto(g_recConnection, "OrderDetail", "OrderDetail", strSQLStatement, strWhereClause, blnCriticalWrite)
    End If
End Function


'
'SQL System Table Public Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLClearSystem(recBuffer As typTableSystem)
    On Error Resume Next
    With recBuffer
        .blnOption1 = False
        .blnOption2 = False
        .blnOption3 = False
        .blnOption4 = False
        .blnOption5 = False
        .blnOption6 = False
    End With
End Sub

Public Sub SQLCloseSystem()
    Call SQLCloseRecordSet(g_objRecordsetSystem)
End Sub

Public Sub SQLFetchSystem(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, recBuffer As typTableSystem)
    On Error Resume Next
    recConnection.intOperationType = enumSQLOperationType.ReadOperation
    With recBuffer
        .blnOption1 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption1")
        .blnOption2 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption2")
        .blnOption3 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption3")
        .blnOption4 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption4")
        .blnOption5 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption5")
        .blnOption6 = SQLFieldGetBoolean(recConnection, objRecordset, "System", "blnOption6")
    End With
End Sub

Public Function SQLNextSystem() As Boolean
    On Error Resume Next
    If SQLNext(g_objRecordsetSystem) Then
        Call SQLFetchSystem(g_recConnection, g_objRecordsetSystem, g_recSystem)
        SQLNextSystem = True
    End If
End Function

Public Function SQLSelectSystem(strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    On Error Resume Next
    If SQLSelect(g_recConnection, g_objRecordsetSystem, "System", "System", strWhereClause, lngTopCount, strRecordSetParams) Then
        Call SQLFetchSystem(g_recConnection, g_objRecordsetSystem, g_recSystem)
        SQLSelectSystem = True
    End If
End Function

Public Function SQLStatementBuildSystem(recConnection As typSQLConnection, recBuffer As typTableSystem, intBuildType As Integer, Optional blnAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As String
    Dim intBuildIndex As Integer

    On Error Resume Next
    intBuildIndex = SQLFieldBuildBegin(recConnection, intBuildType)
    If intBuildIndex >= 0 Then
        With recBuffer
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption1", .blnOption1)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption2", .blnOption2)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption3", .blnOption3)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption4", .blnOption4)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption5", .blnOption5)
            Call SQLFieldBuildBoolean(intBuildIndex, "blnOption6", .blnOption6)
        End With
        SQLStatementBuildSystem = SQLFieldBuildEnd(intBuildIndex)
    End If
End Function

Public Function SQLUpdateIntoSystem(recBuffer As typTableSystem, strWhereClause As String, Optional blnUpdateAutoKeyField As Boolean = False, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQLStatement As String

    On Error Resume Next
    strSQLStatement = SQLStatementBuildSystem(g_recConnection, recBuffer, enumSQLBuildType.UpdateTable, blnUpdateAutoKeyField)
    If strSQLStatement <> "" Then
        SQLUpdateIntoSystem = SQLUpdateInto(g_recConnection, "System", "System", strSQLStatement, strWhereClause, blnCriticalWrite)
    End If
End Function


'
'SQL Version Table Public Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Sub SQLCloseVersion()
    Call SQLCloseRecordSet(g_objRecordsetVersion)
End Sub

Public Sub SQLFetchVersion(recConnection As typSQLConnection, objRecordset As ADODB.Recordset, recBuffer As typTableVersion)
    On Error Resume Next
    recConnection.intOperationType = enumSQLOperationType.ReadOperation
    With recBuffer
        .strSystemName = SQLFieldGetString(recConnection, objRecordset, "Version", "strSystemName")
        .strVersion = SQLFieldGetString(recConnection, objRecordset, "Version", "strVersion")
        .intSchemaVersion = SQLFieldGetInteger(recConnection, objRecordset, "Version", "intSchemaVersion")
    End With
End Sub

Public Function SQLNextVersion() As Boolean
    On Error Resume Next
    If SQLNext(g_objRecordsetVersion) Then
        Call SQLFetchVersion(g_recConnection, g_objRecordsetVersion, g_recVersion)
        SQLNextVersion = True
    End If
End Function

Public Function SQLSelectVersion(strWhereClause As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    On Error Resume Next
    If SQLSelect(g_recConnection, g_objRecordsetVersion, "Version", "Version", strWhereClause, lngTopCount, strRecordSetParams) Then
        Call SQLFetchVersion(g_recConnection, g_objRecordsetVersion, g_recVersion)
        SQLSelectVersion = True
    End If
End Function

'
'SQL Query Routines   Automatically Generated  -  DO NOT MODIFY
'
Public Function SQLQueryDeleteCustomer_ID(lngCustomerID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1)
    SQLQueryDeleteCustomer_ID = SQLDeleteFrom(g_recConnection, "Customer", "Customer", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteCustomer_Name(strName_1 As String, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " strName = " & "'" & SQLFieldEmitString(strName_1, 50, g_recConnection.strEscapeQuote) & "'"
    SQLQueryDeleteCustomer_Name = SQLDeleteFrom(g_recConnection, "Customer", "Customer", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteOrder_ID(lngOrderID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderID = " & CStr(lngOrderID_1)
    SQLQueryDeleteOrder_ID = SQLDeleteFrom(g_recConnection, "Order", "Order", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteOrder_CustID(lngCustomerID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1)
    SQLQueryDeleteOrder_CustID = SQLDeleteFrom(g_recConnection, "Order", "Order", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteOrderDetail_ID(lngOrderDetailID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderDetailID = " & CStr(lngOrderDetailID_1)
    SQLQueryDeleteOrderDetail_ID = SQLDeleteFrom(g_recConnection, "OrderDetail", "OrderDetail", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteOrderDetail_CustID(lngCustomerID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1)
    SQLQueryDeleteOrderDetail_CustID = SQLDeleteFrom(g_recConnection, "OrderDetail", "OrderDetail", strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryDeleteOrderDetail_OrderID(lngOrderID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderID = " & CStr(lngOrderID_1)
    SQLQueryDeleteOrderDetail_OrderID = SQLDeleteFrom(g_recConnection, "OrderDetail", "OrderDetail", strSQL, blnCriticalWrite)
End Function

Public Function SQLQuerySelectCustomer_ID(lngCustomerID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1)
    SQLQuerySelectCustomer_ID = SQLSelectCustomer(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectCustomer_Name(strName_1 As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " strName LIKE " & "'" & SQLFieldEmitString(strName_1, 50, g_recConnection.strEscapeQuote) & "'" & _
        " ORDER BY strName"
    SQLQuerySelectCustomer_Name = SQLSelectCustomer(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrder_ID(lngOrderID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderID = " & CStr(lngOrderID_1)
    SQLQuerySelectOrder_ID = SQLSelectOrder(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrder_CustID(lngCustomerID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1) & _
        " ORDER BY strOrderNo"
    SQLQuerySelectOrder_CustID = SQLSelectOrder(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrder_Date(datOrder_1 As Date, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " datOrder > " & SQLFieldEmitDate(datOrder_1, g_recConnection.strDatabaseSystem) & _
        " ORDER BY datOrder DESC, strOrderNo"
    SQLQuerySelectOrder_Date = SQLSelectOrder(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrder_Salesman(strSalesmanID_1 As String, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " strSalesmanID = " & "'" & SQLFieldEmitString(strSalesmanID_1, 10, g_recConnection.strEscapeQuote) & "'" & _
        " ORDER BY lngCustomerID, strOrderNo"
    SQLQuerySelectOrder_Salesman = SQLSelectOrder(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrderDetail_ID(lngOrderDetailID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderDetailID = " & CStr(lngOrderDetailID_1)
    SQLQuerySelectOrderDetail_ID = SQLSelectOrderDetail(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrderDetail_CustID(lngCustomerID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1) & _
        " ORDER BY lngOrderID, strItemID"
    SQLQuerySelectOrderDetail_CustID = SQLSelectOrderDetail(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySelectOrderDetail_OrderID(lngOrderID_1 As Long, Optional lngTopCount As Long = 0, Optional strRecordSetParams As String = "") As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderID = " & CStr(lngOrderID_1) & _
        " ORDER BY strItemID"
    SQLQuerySelectOrderDetail_OrderID = SQLSelectOrderDetail(strSQL, lngTopCount, strRecordSetParams)
End Function

Public Function SQLQuerySetOrder_Shipped(lngOrderID_1 As Long, blnShipped_2 As Boolean, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " UPDATE Order SET" & _
        " blnShipped = " & IIf(blnShipped_2, "true", "false") & _
        " WHERE" & _
        " lngOrderID = " & CStr(lngOrderID_1)
    SQLQuerySetOrder_Shipped = SQLCommand(strSQL, , , IIf(blnCriticalWrite, enumSQLOperationType.CriticalWriteOperation, enumSQLOperationType.WriteOperation))
End Function

Public Function SQLQueryUpdateCustomer_ID(recBuffer As typTableCustomer, lngCustomerID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngCustomerID = " & CStr(lngCustomerID_1)
    SQLQueryUpdateCustomer_ID = SQLUpdateIntoCustomer(recBuffer, strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryUpdateOrder_ID(recBuffer As typTableOrder, lngOrderID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderID = " & CStr(lngOrderID_1)
    SQLQueryUpdateOrder_ID = SQLUpdateIntoOrder(recBuffer, strSQL, blnCriticalWrite)
End Function

Public Function SQLQueryUpdateOrderDetail_ID(recBuffer As typTableOrderDetail, lngOrderDetailID_1 As Long, Optional blnCriticalWrite As Boolean = False) As Boolean
    Dim strSQL As String

    On Error Resume Next
    strSQL = _
        " lngOrderDetailID = " & CStr(lngOrderDetailID_1)
    SQLQueryUpdateOrderDetail_ID = SQLUpdateIntoOrderDetail(recBuffer, strSQL, blnCriticalWrite)
End Function

'
'End of Automatically Generated Code
'
