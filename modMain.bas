Attribute VB_Name = "modMain"
Option Explicit
Option Base 0
Option Compare Text

'
'SQL Code Generator
'
'This program requests an input file, which must be an XML schema file, and
'which must contain a schema definition of one or more database tables.
'The schema will be read into the tables of this program, and used, along
'with a template code file, to generate a VB code module with routines that
'are capable of performing SQL operations on the above tables.
'
Private Enum enumQueryLine
    OrderByField = 1
    QueryField = 2
    QueryFieldFixed = 3
    SetField = 4
    SetFieldFixed = 5
    Conjunction = 6
    Parenthesis = 7
End Enum

Private Enum enumQueryType
    DeleteQuery = 1
    SelectQuery = 2
    SetQuery = 3
    UpdateQuery = 4
End Enum

Private Type typConnection
    strName As String
End Type
Public g_recConnection() As typConnection

Private Type typTable
    strNameExternal As String
    strNameInternal As String
    strConnection As String
    intFieldCount As Integer
    intFieldPtr As Integer
    intBuffers As Integer
    blnCreateTable As Boolean
    blnNoAdd As Boolean
    blnNoModify As Boolean
    blnNoDelete As Boolean
End Type
Public g_recTable() As typTable

Private Type typField
    strNameExternal As String
    strNameInternal As String
    strType As String
    strTypeADO As String
    blnAutoKey As Boolean
    blnIndex As Boolean
    blnPrimary As Boolean
    blnUnique As Boolean
    lngLength As Long
End Type
Public g_recField() As typField

Private Type typQuery
    intQueryType As Integer
    strQueryName As String
    strQueryConnection As String
    strQueryTableExternal As String
    strQueryTableInternal As String
    intQueryTableBuffer As Integer
    intQueryLineCount As Integer
    intQueryLinePtr As Integer
End Type
Public g_recQuery() As typQuery

Private Type typQueryLine
    intQueryLineType As Integer
    lngFieldLength As Long
    strFieldName As String
    strFieldType As String
    strFieldValue As String
    strFormat As String
    strOperator As String
End Type
Public g_recQueryLine() As typQueryLine

Public Const enumInsertConnectionName As String = "connectionname"
Public Const enumInsertFieldAdoType As String = "fieldadotype"
Public Const enumInsertFieldAutoKey As String = "fieldautokey"
Public Const enumInsertFieldClear As String = "fieldclear"
Public Const enumInsertFieldComma As String = "fieldcomma"
Public Const enumInsertFieldIndex As String = "fieldindex"
Public Const enumInsertFieldLength As String = "fieldlength"
Public Const enumInsertFieldNameExternal As String = "fieldnameexternal"
Public Const enumInsertFieldNameInternal As String = "fieldnameinternal"
Public Const enumInsertFieldPrimary As String = "fieldprimary"
Public Const enumInsertFieldType As String = "fieldtype"
Public Const enumInsertFieldTypeActual As String = "fieldtypeactual"
Public Const enumInsertFieldUnique As String = "fieldunique"
Public Const enumInsertOutputName As String = "outputname"
Public Const enumInsertQueryConnection As String = "queryconnection"
Public Const enumInsertQueryBufferID As String = "querybufferid"
Public Const enumInsertQueryName As String = "queryname"
Public Const enumInsertQueryParam As String = "queryparam"
Public Const enumInsertQueryString As String = "querystring"
Public Const enumInsertQueryTableNameExternal As String = "querytablenameexternal"
Public Const enumInsertQueryTableNameInternal As String = "querytablenameinternal"
Public Const enumInsertTableBufferID As String = "tablebufferid"
Public Const enumInsertTableConnection As String = "tableconnection"
Public Const enumInsertTableNameExternal As String = "tablenameexternal"
Public Const enumInsertTableNameExternalLC As String = "tablenameexternallc"
Public Const enumInsertTableNameInternal As String = "tablenameinternal"
Public Const enumInsertTableNameInternalLC As String = "tablenameinternallc"

Public Const enumOptionAddnew As String = "addnew"
Public Const enumOptionClear As String = "clear"
Public Const enumOptionCompactRepair As String = "compactrepair"
Public Const enumOptionDelete As String = "delete"
Public Const enumOptionDeleteFrom As String = "deletefrom"
Public Const enumOptionInsert As String = "insert"
Public Const enumOptionParameter As String = "parameter"
Public Const enumOptionRecordCount As String = "recordcount"
Public Const enumOptionSearch As String = "search"
Public Const enumOptionTransaction As String = "transaction"
Public Const enumOptionUpdate As String = "update"
Public Const enumOptionUpdateInto As String = "updateinto"
Public Const enumOptionVB6 As String = "vb6"
Public Const enumOptionVBA As String = "vba"
Public Const enumOptionVBNet As String = "vbnet"

Public Const enumSchemaConnection As String = "<connection "
Public Const enumSchemaConnectionEnd As String = "</connection>"
Public Const enumSchemaTable As String = "<table "
Public Const enumSchemaTableEnd As String = "</table>"
Public Const enumSchemaOption As String = "<option "
Public Const enumSchemaOutput As String = "<output "
Public Const enumSchemaQueryConnection As String = "<queryconnection "
Public Const enumSchemaQueryConnectionEnd As String = "</queryconnection>"
Public Const enumSchemaQueryEnd As String = "</query>"
Public Const enumSchemaQueryDelete As String = "<querydelete "
Public Const enumSchemaQuerySelect As String = "<queryselect "
Public Const enumSchemaQuerySet As String = "<queryset "
Public Const enumSchemaQueryUpdate As String = "<queryupdate "
Public Const enumSchemaXML As String = "<?xml "

Public Const enumTemplateBuffer As String = "<buffer>"
Public Const enumTemplateBufferEnd As String = "</buffer>"
Public Const enumTemplateConnection As String = "<connection>"
Public Const enumTemplateConnectionEnd As String = "</connection>"
Public Const enumTemplateField As String = "<field>"
Public Const enumTemplateFieldEnd As String = "</field>"
Public Const enumTemplateFieldAutoKey As String = "<fieldautokey>"
Public Const enumTemplateFieldAutoKeyEnd As String = "</fieldautokey>"
Public Const enumTemplateFieldNotAutoKey As String = "<fieldnotautokey>"
Public Const enumTemplateFieldNotAutoKeyEnd As String = "</fieldnotautokey>"
Public Const enumTemplateFieldType As String = "<fieldtype>"
Public Const enumTemplateFieldTypeEnd As String = "</fieldtype>"
Public Const enumTemplateLogic As String = "<logic>"
Public Const enumTemplateLogicEnd As String = "</logic>"
Public Const enumTemplateOption As String = "<option>"
Public Const enumTemplateOptionEnd As String = "</option>"
Public Const enumTemplateQueryDelete As String = "<querydelete>"
Public Const enumTemplateQueryDeleteEnd As String = "</querydelete>"
Public Const enumTemplateQuerySelect As String = "<queryselect>"
Public Const enumTemplateQuerySelectEnd As String = "</queryselect>"
Public Const enumTemplateQuerySet As String = "<queryset>"
Public Const enumTemplateQuerySetEnd As String = "</queryset>"
Public Const enumTemplateQueryUpdate As String = "<queryupdate>"
Public Const enumTemplateQueryUpdateEnd As String = "</queryupdate>"
Public Const enumTemplateTable As String = "<table>"
Public Const enumTemplateTableEnd As String = "</table>"

Public g_strModuleFile As String
Public g_strTemplateFile As String
Public g_strXMLFile As String

Private m_blnComment As Boolean
Private m_blnProcessLine As Boolean
Private m_blnVBA As Boolean
Private m_blnVBNet As Boolean
Private m_intConnectionCount As Integer
Private m_intConnectionIndex As Integer
Private m_intFieldCount As Integer
Private m_intFieldIndex As Integer
Private m_intFileNoXML As Integer
Private m_intFileNoModule As Integer
Private m_intFileNoTemplate As Integer
Private m_intOptionNested As Integer
Private m_intQueryCount As Integer
Private m_intQueryIndex As Integer
Private m_intQueryLineCount As Integer
Private m_intTableBuffer As Integer
Private m_intTableCount As Integer
Private m_intTableIndex As Integer
Private m_lngCount As Long
Private m_strArray() As String
Private m_strCurrentConnection As String
Private m_strCurrentQueryConnection As String
Private m_strLine As String
Private m_strLines() As String
Private m_strOptionArray() As String

Public Sub Main()
'
'   Initialization
'
    Call UtilityInitialize
    frmGenerator.Show
End Sub

Public Sub Process()
'
'   Main processing routine
'
    Dim intLineIndex As Integer
    Dim intX As Integer
    Dim strFileName As String
    
    On Error Resume Next
    Err.Clear
    m_intFileNoXML = FreeFile
    Open g_strXMLFile For Input As #m_intFileNoXML
    If Err <> 0 Then
        MsgBox "Unable to open the XML file '" & g_strXMLFile & "'", vbCritical
        End
    End If
    m_intConnectionCount = 0
    m_intTableCount = 0
    m_intFieldCount = 0
    m_strCurrentConnection = ""
    ReDim m_strOptionArray(0)
    
    'Read in the XML schema file
    Do Until EOF(m_intFileNoXML)
        If ProcessSchemaLine Then
            If InStr(LCase$(m_strLine), enumSchemaTable) <> 0 Then
                Call ProcessSchemaTable
            ElseIf InStr(LCase$(m_strLine), enumSchemaQueryDelete) <> 0 Then
                Call ProcessSchemaQuery(enumQueryType.DeleteQuery)
            ElseIf InStr(LCase$(m_strLine), enumSchemaQuerySelect) <> 0 Then
                Call ProcessSchemaQuery(enumQueryType.SelectQuery)
            ElseIf InStr(LCase$(m_strLine), enumSchemaQuerySet) <> 0 Then
                Call ProcessSchemaQuery(enumQueryType.SetQuery)
            ElseIf InStr(LCase$(m_strLine), enumSchemaQueryUpdate) <> 0 Then
                Call ProcessSchemaQuery(enumQueryType.UpdateQuery)
            ElseIf InStr(LCase$(m_strLine), enumSchemaQueryConnection) <> 0 Then
                Call ProcessSchemaQueryConnection
            ElseIf InStr(LCase$(m_strLine), enumSchemaQueryConnectionEnd) <> 0 Then
                m_strCurrentQueryConnection = ""
            ElseIf InStr(LCase$(m_strLine), enumSchemaConnection) <> 0 Then
                Call ProcessSchemaConnection
            ElseIf InStr(LCase$(m_strLine), enumSchemaConnectionEnd) <> 0 Then
                m_strCurrentConnection = ""
            ElseIf InStr(LCase$(m_strLine), enumSchemaOption) <> 0 Then
                Call ProcessSchemaOption
            ElseIf InStr(LCase$(m_strLine), enumSchemaOutput) <> 0 Then
                Call ProcessSchemaOutput
            End If
        End If
    Loop
    Close #m_intFileNoXML
    If g_strModuleFile = "" Then
        MsgBox "No '" & enumSchemaOutput & "' statement was included in the schema file to specify the output module name", vbCritical
        End
    End If
    If m_blnVBNet Then
        g_strTemplateFile = SetSlash(App.Path) & "SQLCodeVBNet.txt"
    Else
        g_strTemplateFile = SetSlash(App.Path) & "SQLCodeVB6.txt"
    End If
    If Not FileExists(g_strTemplateFile) Then
        MsgBox "The template file '" & g_strTemplateFile & "' does not exist", vbCritical
        End
    End If
    
    'Load the template file into memory
    m_intFileNoTemplate = FreeFile
    Open g_strTemplateFile For Input As #m_intFileNoTemplate
    If Err.Number <> 0 Then
        MsgBox "Unable to open the template file '" & g_strTemplateFile & "'", vbCritical
        End
    End If
    If m_intConnectionCount = 0 Then
        ReDim Preserve g_recConnection(0)
        g_recConnection(0).strName = ""
        m_intConnectionCount = 1
    End If
    m_blnProcessLine = True
    m_intOptionNested = 0
    ReDim m_strLines(0)
    intLineIndex = 0
    Do Until EOF(m_intFileNoTemplate)
        Line Input #m_intFileNoTemplate, m_strLine
        If InStr(LCase$(m_strLine), enumTemplateOption) <> 0 Then
            If m_blnProcessLine Then
                m_blnProcessLine = ProcessOutputOption()
            Else
                m_intOptionNested = m_intOptionNested + 1
            End If
        ElseIf InStr(LCase$(m_strLine), enumTemplateOptionEnd) <> 0 Then
            If m_intOptionNested > 0 Then
                m_intOptionNested = m_intOptionNested - 1
            Else
                m_blnProcessLine = True
            End If
        ElseIf m_blnProcessLine Then
            intLineIndex = intLineIndex + 1
            ReDim Preserve m_strLines(intLineIndex)
            m_strLines(intLineIndex - 1) = m_strLine
        End If
    Loop
    Close #m_intFileNoTemplate
    intLineIndex = 0
    
    'Output the module file
    strFileName = SetSlash(GetFilePath(g_strXMLFile)) & g_strModuleFile
    m_intFileNoModule = FreeFile
    Open strFileName For Output As #m_intFileNoModule
    If Err <> 0 Then
        MsgBox "Unable to open the output module file '" & g_strModuleFile & "'", vbCritical
        End
    End If
    Do Until intLineIndex >= UBound(m_strLines)
        m_strLine = m_strLines(intLineIndex)
        If InStr(LCase$(m_strLine), enumTemplateTable) <> 0 Then
            Call ProcessOutputTable(intLineIndex)
            Call ProcessOutputReposition(intLineIndex, enumTemplateTableEnd)
        ElseIf InStr(LCase$(m_strLine), enumTemplateConnection) <> 0 Then
            Call ProcessOutputConnection(intLineIndex)
            Call ProcessOutputReposition(intLineIndex, enumTemplateConnectionEnd)
        ElseIf InStr(LCase$(m_strLine), enumTemplateQueryDelete) <> 0 Then
            Call ProcessOutputQuery(intLineIndex, enumQueryType.DeleteQuery)
            Call ProcessOutputReposition(intLineIndex, enumTemplateQueryDeleteEnd)
        ElseIf InStr(LCase$(m_strLine), enumTemplateQuerySelect) <> 0 Then
            Call ProcessOutputQuery(intLineIndex, enumQueryType.SelectQuery)
            Call ProcessOutputReposition(intLineIndex, enumTemplateQuerySelectEnd)
        ElseIf InStr(LCase$(m_strLine), enumTemplateQuerySet) <> 0 Then
            Call ProcessOutputQuery(intLineIndex, enumQueryType.SetQuery)
            Call ProcessOutputReposition(intLineIndex, enumTemplateQuerySetEnd)
        ElseIf InStr(LCase$(m_strLine), enumTemplateQueryUpdate) <> 0 Then
            Call ProcessOutputQuery(intLineIndex, enumQueryType.UpdateQuery)
            Call ProcessOutputReposition(intLineIndex, enumTemplateQueryUpdateEnd)
        Else
            Call ProcessOutputInsert
            Print #m_intFileNoModule, m_strLine
        End If
        intLineIndex = intLineIndex + 1
    Loop
    Close #m_intFileNoModule
    Err.Clear
    End
End Sub

Private Sub ProcessOutputConnection(intLineIndexCurrent As Integer)
'
'   Emits code for a Connection loop
'
    Dim intLineIndex As Integer
    Dim intLineIndexStart As Integer
    
    For m_intConnectionIndex = 0 To m_intConnectionCount - 1
        For intLineIndex = intLineIndexCurrent + 1 To UBound(m_strLines) - 1
            m_strLine = m_strLines(intLineIndex)
            If InStr(LCase$(m_strLine), enumTemplateConnection) <> 0 Then
                intLineIndexStart = intLineIndex
                For m_intTableIndex = 0 To m_intTableCount - 1
                    If g_recTable(m_intTableIndex).strConnection = g_recConnection(m_intConnectionIndex).strName Then
                        intLineIndex = intLineIndexStart
                        Do
                            intLineIndex = intLineIndex + 1
                            If intLineIndex >= UBound(m_strLines) Then
                                MsgBox "A '" & enumTemplateConnectionEnd & "' line is missing from the template file", vbCritical
                                End
                            End If
                            m_strLine = m_strLines(intLineIndex)
                            If InStr(LCase$(m_strLine), enumTemplateConnectionEnd) <> 0 Then
                                Exit Do
                            ElseIf InStr(LCase$(m_strLine), enumTemplateField) <> 0 Then
                                Call ProcessOutputField(intLineIndex)
                                Call ProcessOutputReposition(intLineIndex, enumTemplateFieldEnd)
                            Else
                                Call ProcessOutputInsert
                                Print #m_intFileNoModule, m_strLine
                            End If
                        Loop
                    End If
                Next
            ElseIf InStr(LCase$(m_strLine), enumTemplateConnectionEnd) <> 0 Then
                Exit For
            Else
                Call ProcessOutputInsert
                Print #m_intFileNoModule, m_strLine
            End If
        Next
    Next
End Sub

Private Sub ProcessOutputField(intLineIndexCurrent As Integer)
'
'   Emits code for a field loop
'
    Dim intX As Integer
    Dim intLineIndex As Integer
    
    For m_intFieldIndex = g_recTable(m_intTableIndex).intFieldPtr To g_recTable(m_intTableIndex).intFieldPtr + g_recTable(m_intTableIndex).intFieldCount - 1
        intLineIndex = intLineIndexCurrent
        Do
            intLineIndex = intLineIndex + 1
            If intLineIndex >= UBound(m_strLines) Then
                MsgBox "A '" & enumTemplateFieldEnd & "' line is missing from the template file", vbCritical
                End
            End If
            m_strLine = m_strLines(intLineIndex)
            If InStr(LCase$(m_strLine), enumTemplateFieldEnd) <> 0 Then
                Exit Do
            End If
            If InStr(LCase$(m_strLine), enumTemplateFieldAutoKey) <> 0 Then
                If Not g_recField(m_intFieldIndex).blnAutoKey Then
                    Call ProcessOutputReposition(intLineIndex, enumTemplateFieldAutoKeyEnd)
                End If
            ElseIf InStr(LCase$(m_strLine), enumTemplateFieldNotAutoKey) <> 0 Then
                If g_recField(m_intFieldIndex).blnAutoKey Then
                    Call ProcessOutputReposition(intLineIndex, enumTemplateFieldNotAutoKeyEnd)
                End If
            ElseIf InStr(LCase$(m_strLine), enumTemplateFieldType) <> 0 Then
                intX = InStr(LCase$(m_strLine), enumTemplateFieldType)
                If Not ProcessOutputFieldType(Mid$(m_strLine, intX + Len(enumTemplateFieldType))) Then
                    Call ProcessOutputReposition(intLineIndex, enumTemplateFieldTypeEnd)
                End If
            ElseIf InStr(LCase$(m_strLine), enumTemplateFieldAutoKeyEnd) = 0 _
                And InStr(LCase$(m_strLine), enumTemplateFieldNotAutoKeyEnd) = 0 _
                And InStr(LCase$(m_strLine), enumTemplateFieldTypeEnd) = 0 Then
                Call ProcessOutputInsert
                Print #m_intFileNoModule, m_strLine
            End If
        Loop
    Next
End Sub

Private Function ProcessOutputFieldType(strLine As String) As Boolean
'
'   Check a field type to see if code should be emitted for it
'
    Dim blnNot As Boolean
    Dim blnResult As Boolean
    Dim intX As Integer
    
    blnNot = False
    blnResult = False
    m_strArray = Split(strLine, " ")
    For intX = 0 To UBound(m_strArray)
        If LCase$(m_strArray(intX)) = "not" Then
            blnNot = True
        Else
            If LCase$(g_recField(m_intFieldIndex).strType) = LCase$(m_strArray(intX)) Then
                blnResult = True
                Exit For
            End If
        End If
    Next
    If blnNot Then
        blnResult = Not blnResult
    End If
    ProcessOutputFieldType = blnResult
End Function

Private Sub ProcessOutputInsert()
'
'   Checks for insert parameters and replaces them with values
'
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim intLength As Integer
    Dim intParam As Integer
    Dim intX As Integer
    Dim strOrderBy As String
    Dim strSet As String
    Dim strSetWhereClause As String
    Dim strString As String
    Dim strValue As String
    
    Do
        intBegin = InStr(m_strLine, "[{")
        If intBegin = 0 Then
            Exit Do
        End If
        intEnd = InStr(intBegin, m_strLine, "}]")
        If intEnd = 0 Then
            MsgBox "An '}]' tag is missing from a line in the template file", vbCritical
            End
        End If
        intLength = intEnd - intBegin + 2
        Select Case LCase$(Mid$(m_strLine, intBegin + 2, intLength - 4))
            Case Is = enumInsertConnectionName
                strValue = g_recConnection(m_intConnectionIndex).strName
            Case Is = enumInsertFieldAdoType
                strValue = g_recField(m_intFieldIndex).strTypeADO
            Case Is = enumInsertFieldAutoKey
                strValue = IIf(g_recField(m_intFieldIndex).blnAutoKey, "True", "False")
            Case Is = enumInsertFieldClear
                Select Case LCase$(g_recField(m_intFieldIndex).strType)
                    Case Is = "boolean"
                        strValue = "False"
                    Case Is = "memo", "string"
                        strValue = """" & """"
                    Case Else
                        strValue = "0"
                End Select
            Case Is = enumInsertFieldComma
                If m_intFieldIndex <= g_recTable(m_intTableIndex).intFieldPtr Then
                    strValue = ""
                Else
                    strValue = ", "
                End If
            Case Is = enumInsertFieldIndex
                strValue = IIf(g_recField(m_intFieldIndex).blnIndex, "True", "False")
            Case Is = enumInsertFieldLength
                strValue = CStr(g_recField(m_intFieldIndex).lngLength)
            Case Is = enumInsertFieldNameExternal
                strValue = g_recField(m_intFieldIndex).strNameExternal
            Case Is = enumInsertFieldNameInternal
                strValue = g_recField(m_intFieldIndex).strNameInternal
            Case Is = enumInsertFieldPrimary
                strValue = IIf(g_recField(m_intFieldIndex).blnPrimary, "True", "False")
            Case Is = enumInsertFieldType
                strValue = IIf(LCase$(g_recField(m_intFieldIndex).strType) = "memo", "String", g_recField(m_intFieldIndex).strType)
            Case Is = enumInsertFieldTypeActual
                strValue = g_recField(m_intFieldIndex).strType
            Case Is = enumInsertFieldUnique
                strValue = IIf(g_recField(m_intFieldIndex).blnUnique, "True", "False")
            Case Is = enumInsertOutputName
                strValue = GetFileName(g_strModuleFile, True)
            Case Is = enumInsertQueryConnection
                strValue = g_recQuery(m_intQueryIndex).strQueryConnection
            Case Is = enumInsertQueryBufferID
                strValue = IIf(g_recQuery(m_intQueryIndex).intQueryTableBuffer <= 1, "", "_" & Format$(g_recQuery(m_intQueryIndex).intQueryTableBuffer))
            Case Is = enumInsertQueryName
                strValue = g_recQuery(m_intQueryIndex).strQueryName
            Case Is = enumInsertQueryParam
                intParam = 0
                strValue = ""
                For intX = g_recQuery(m_intQueryIndex).intQueryLinePtr To g_recQuery(m_intQueryIndex).intQueryLinePtr + g_recQuery(m_intQueryIndex).intQueryLineCount - 1
                    If g_recQueryLine(intX).intQueryLineType = enumQueryLine.QueryField _
                        Or g_recQueryLine(intX).intQueryLineType = enumQueryLine.SetField Then
                        intParam = intParam + 1
                        strValue = strValue & _
                            IIf(strValue = "", "", ", ") & _
                            g_recQueryLine(intX).strFieldName & "_" & Format$(intParam) & " As " & IIf(LCase$(g_recQueryLine(intX).strFieldType) = "memo", "String", g_recQueryLine(intX).strFieldType)
                    End If
                Next
                If strValue <> "" Then
                    strValue = strValue & ", "
                End If
            Case Is = enumInsertQueryString
                intParam = 0
                strOrderBy = ""
                strSet = ""
                strValue = ""
                For intX = g_recQuery(m_intQueryIndex).intQueryLinePtr To g_recQuery(m_intQueryIndex).intQueryLinePtr + g_recQuery(m_intQueryIndex).intQueryLineCount - 1
                    Select Case g_recQueryLine(intX).intQueryLineType
                        Case Is = enumQueryLine.Conjunction, enumQueryLine.Parenthesis
                            strValue = strValue & _
                                IIf(strValue = "", "", " & _" & vbCrLf & Space$(intBegin - 1)) & _
                                """ " & _
                                g_recQueryLine(intX).strFieldName & _
                                """"
                        Case Is = enumQueryLine.QueryFieldFixed
                            strValue = strValue & _
                                IIf(strValue = "", "", " & _" & vbCrLf & Space$(intBegin - 1)) & _
                                """ " & _
                                g_recQueryLine(intX).strFieldName & _
                                " " & _
                                g_recQueryLine(intX).strOperator & _
                                " " & _
                                g_recQueryLine(intX).strFieldValue & _
                                """"
                        Case Is = enumQueryLine.SetFieldFixed
                            strSet = strSet & _
                                IIf(strSet = "", "", " & "","" & _" & vbCrLf & Space$(intBegin - 1)) & _
                                """ " & _
                                g_recQueryLine(intX).strFieldName & _
                                " " & _
                                g_recQueryLine(intX).strOperator & _
                                " " & _
                                g_recQueryLine(intX).strFieldValue & _
                                """"
                        Case Is = enumQueryLine.OrderByField
                            strOrderBy = strOrderBy & _
                                IIf(strOrderBy = "", "", ", ") & _
                                g_recQueryLine(intX).strFieldName & _
                                IIf(g_recQueryLine(intX).strFieldValue = "", "", " " & g_recQueryLine(intX).strFieldValue)
                        Case Is = enumQueryLine.QueryField, enumQueryLine.SetField
                            intParam = intParam + 1
                            strString = """ " & _
                                g_recQueryLine(intX).strFieldName & _
                                " " & _
                                g_recQueryLine(intX).strOperator & _
                                " """ & " & "
                            Select Case g_recQueryLine(intX).strFieldType
                                Case Is = "Boolean"
                                    strString = strString & _
                                        "IIf(" & g_recQueryLine(intX).strFieldName & _
                                        "_" & CStr(intParam) & _
                                        ", " & """true""" & ", " & """false""" & ")"
                                Case Is = "Date"
                                    strString = strString & _
                                        "SQLFieldEmitDate(" & _
                                        g_recQueryLine(intX).strFieldName & _
                                        "_" & CStr(intParam) & _
                                        ", g_recConnection" & g_recQuery(m_intQueryIndex).strQueryConnection & ".strDatabaseSystem" & _
                                        ")"
                                Case Is = "Integer", "Long", "Double", "Single", "Currency"
                                    If g_recQueryLine(intX).strFormat = "" Then
                                        strString = strString & _
                                            "CStr(" & _
                                            g_recQueryLine(intX).strFieldName & _
                                            "_" & Format$(intParam) & _
                                            ")"
                                    Else
                                        strString = strString & _
                                            "Format$(" & _
                                            g_recQueryLine(intX).strFieldName & _
                                            "_" & Format$(intParam) & _
                                            "," & g_recQueryLine(intX).strFormat & _
                                            ")"
                                    End If
                                Case Is = "String", "Memo"
                                    strString = strString & _
                                        """'""" & _
                                        " & SQLFieldEmitString(" & _
                                        g_recQueryLine(intX).strFieldName & _
                                        "_" & CStr(intParam) & _
                                        ", " & CStr(g_recQueryLine(intX).lngFieldLength) & _
                                        ", " & "g_recConnection" & g_recQuery(m_intQueryIndex).strQueryConnection & ".strEscapeQuote" & _
                                        ") & " & _
                                        """'"""
                            End Select
                            If g_recQueryLine(intX).intQueryLineType = enumQueryLine.SetField Then
                                strSet = strSet & _
                                    IIf(strSet = "", "", " & "","" & _" & vbCrLf & Space$(intBegin - 1)) & _
                                    strString
                            Else
                                strValue = strValue & _
                                    IIf(strValue = "", "", " & _" & vbCrLf & Space$(intBegin - 1)) & _
                                    strString
                            End If
                    End Select
                Next
                If strOrderBy <> "" Then
                    strValue = strValue & _
                        IIf(strValue = "", "", " & _" & vbCrLf & Space$(intBegin - 1)) & _
                        """" & _
                        " ORDER BY " & strOrderBy & _
                        """"
                ElseIf strSet <> "" Then
                    strSetWhereClause = strValue
                    strValue = """" & _
                        " UPDATE " & _
                        g_recQuery(m_intQueryIndex).strQueryTableExternal & _
                        " SET" & """" & " & _" & vbCrLf & Space$(intBegin - 1) & _
                        strSet
                        If strSetWhereClause <> "" Then
                            strValue = strValue & _
                                " & _" & vbCrLf & _
                                Space$(intBegin - 1) & """" & _
                                " WHERE" & _
                                """" & " & _" & vbCrLf & Space$(intBegin - 1) & _
                                strSetWhereClause
                        End If
                End If
            Case Is = enumInsertQueryTableNameExternal
                strValue = g_recQuery(m_intQueryIndex).strQueryTableExternal
            Case Is = enumInsertQueryTableNameInternal
                strValue = g_recQuery(m_intQueryIndex).strQueryTableInternal
            Case Is = enumInsertTableBufferID
                strValue = IIf(m_intTableBuffer <= 1, "", "_" & Format$(m_intTableBuffer))
            Case Is = enumInsertTableConnection
                strValue = g_recTable(m_intTableIndex).strConnection
            Case Is = enumInsertTableNameExternal
                strValue = g_recTable(m_intTableIndex).strNameExternal
            Case Is = enumInsertTableNameExternalLC
                strValue = LCase$(g_recTable(m_intTableIndex).strNameExternal)
            Case Is = enumInsertTableNameInternal
                strValue = g_recTable(m_intTableIndex).strNameInternal
            Case Is = enumInsertTableNameInternalLC
                strValue = LCase$(g_recTable(m_intTableIndex).strNameInternal)
            Case Else
                MsgBox "An invalid insert tag was detected in the template file - '" & Mid$(m_strLine, intBegin, intLength) & "'", vbCritical
                End
        End Select
        m_strLine = Left$(m_strLine, intBegin - 1) & strValue & Mid$(m_strLine, intBegin + intLength)
    Loop
End Sub

Private Function ProcessOutputOption() As Boolean
'
'   Determine if code will be included
'
    Dim blnNot As Boolean
    Dim blnResult As Boolean
    Dim intX As Integer
    Dim intY As Integer
        
    blnResult = False
    blnNot = False
    intX = InStr(LCase$(m_strLine), enumTemplateOption)
    If intX <> 0 Then
        m_strArray = Split(LCase$(Mid$(m_strLine, intX + Len(enumSchemaOption))), " ")
        For intY = 0 To UBound(m_strArray)
            If m_strArray(intY) = "not" Then
                blnNot = True
            Else
                For intX = 0 To UBound(m_strOptionArray) - 1
                    If LCase$(m_strOptionArray(intX)) = m_strArray(intY) Then
                        If Not blnNot Then
                            blnResult = True
                        End If
                        Exit For
                    End If
                Next
                If blnNot And intX >= UBound(m_strOptionArray) Then
                    blnResult = Not blnResult
                End If
                blnNot = False
            End If
        Next
    End If
    ProcessOutputOption = blnResult
End Function

Private Sub ProcessOutputReposition(intLineIndex As Integer, strEndKey As String)
'
'   Move through the memory version of the template file contents
'
    Do Until intLineIndex >= UBound(m_strLines)
        If InStr(LCase$(m_strLines(intLineIndex)), strEndKey) <> 0 Then
            Exit Do
        End If
        intLineIndex = intLineIndex + 1
    Loop
    If intLineIndex >= UBound(m_strLines) Then
        MsgBox "A '" & strEndKey & "' line was not found in the template file", vbCritical
        End
    End If
End Sub

Private Sub ProcessOutputQuery(intLineIndexCurrent As Integer, intQueryType As Integer)
'
'   Emits code for a series of queries
'
    Dim intLineIndex As Integer
    Dim intLineIndexStart As Integer
    Dim intX As Integer
    Dim strIO As String
    
    For m_intQueryIndex = 0 To m_intQueryCount - 1
        If g_recQuery(m_intQueryIndex).intQueryType = intQueryType Then
            For intLineIndex = intLineIndexCurrent + 1 To UBound(m_strLines) - 1
                m_strLine = m_strLines(intLineIndex)
                If InStr(LCase$(m_strLine), enumTemplateQueryDeleteEnd) <> 0 _
                    Or InStr(LCase$(m_strLine), enumTemplateQuerySelectEnd) <> 0 _
                    Or InStr(LCase$(m_strLine), enumTemplateQuerySetEnd) <> 0 _
                    Or InStr(LCase$(m_strLine), enumTemplateQueryUpdateEnd) <> 0 Then
                    Exit For
                End If
                Call ProcessOutputInsert
                Print #m_intFileNoModule, m_strLine
            Next
        End If
    Next
End Sub

Private Sub ProcessOutputTable(intLineIndexCurrent As Integer)
'
'   Emits code for a table loop
'
    Dim intLineIndex As Integer
    Dim intLineIndexStart As Integer
    Dim intX As Integer
    Dim strIO As String
    Dim strLine As String
    
    For m_intTableIndex = 0 To m_intTableCount - 1
        For intLineIndex = intLineIndexCurrent + 1 To UBound(m_strLines) - 1
            m_strLine = m_strLines(intLineIndex)
            If InStr(LCase$(m_strLine), enumTemplateLogic) <> 0 Then
                intX = InStr(LCase$(m_strLine), enumTemplateLogic)
                strIO = LCase$(Trim$(Mid$(m_strLine, intX + Len(enumTemplateLogic))))
                If (strIO = "add" And g_recTable(m_intTableIndex).blnNoAdd) _
                    Or (strIO = "modify" And g_recTable(m_intTableIndex).blnNoModify) _
                    Or (strIO = "addmodify" And g_recTable(m_intTableIndex).blnNoAdd And g_recTable(m_intTableIndex).blnNoModify) _
                    Or (strIO = "delete" And g_recTable(m_intTableIndex).blnNoDelete) _
                    Or (strIO = "create" And Not g_recTable(m_intTableIndex).blnCreateTable) Then
                    Do
                        intLineIndex = intLineIndex + 1
                        If intLineIndex >= UBound(m_strLines) Then
                            MsgBox "A '" & enumTemplateLogicEnd & "' line is missing from the template file", vbCritical
                            End
                        End If
                        m_strLine = m_strLines(intLineIndex)
                        If InStr(LCase$(m_strLine), enumTemplateLogicEnd) <> 0 Then
                            Exit Do
                        End If
                    Loop
                End If
            ElseIf InStr(LCase$(m_strLine), enumTemplateBuffer) <> 0 Then
                intLineIndexStart = intLineIndex
                For m_intTableBuffer = 1 To g_recTable(m_intTableIndex).intBuffers
                    intLineIndex = intLineIndexStart
                    Do
                        intLineIndex = intLineIndex + 1
                        If intLineIndex >= UBound(m_strLines) Then
                            MsgBox "A '" & enumTemplateBufferEnd & "' line is missing from the template file", vbCritical
                            End
                        End If
                        m_strLine = m_strLines(intLineIndex)
                        If InStr(LCase$(m_strLine), enumTemplateBufferEnd) <> 0 Then
                            Exit Do
                        End If
                        If InStr(LCase$(m_strLine), enumTemplateField) <> 0 Then
                            Call ProcessOutputField(intLineIndex)
                            Call ProcessOutputReposition(intLineIndex, enumTemplateFieldEnd)
                        Else
                            Call ProcessOutputInsert
                            Print #m_intFileNoModule, m_strLine
                        End If
                    Loop
                Next
            ElseIf InStr(LCase$(m_strLine), enumTemplateField) <> 0 Then
                Call ProcessOutputField(intLineIndex)
                Call ProcessOutputReposition(intLineIndex, enumTemplateFieldEnd)
            ElseIf InStr(LCase$(m_strLine), enumTemplateLogicEnd) <> 0 Then
            ElseIf InStr(LCase$(m_strLine), enumTemplateTableEnd) <> 0 Then
                Exit For
            Else
                Call ProcessOutputInsert
                Print #m_intFileNoModule, m_strLine
            End If
        Next
    Next
End Sub

Private Sub ProcessSchemaConnection()
    '
    '   Scans and parses a Connection definition
    '
    Dim intX As Integer
    Dim strLine As String
        
    intX = InStr(LCase$(m_strLine), enumSchemaConnection)
    If intX > 0 Then
        strLine = Trim$(Mid$(m_strLine, intX + Len(enumSchemaConnection)))
        If Right$(strLine, 1) = ">" Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        If strLine = "" Then
            Call ProcessSchemaError("Connection line contains no connection name value")
        End If
        m_strArray = Split(strLine, " ")
        m_strCurrentConnection = m_strArray(0)
    Else
        Exit Sub
    End If
    For intX = 1 To Len(m_strCurrentConnection)
        If InStr("abcdefghijklmnopqrstuvwxyz0123456789_", LCase$(Mid$(m_strCurrentConnection, intX, 1))) = 0 Then
            Mid$(m_strCurrentConnection, intX, 1) = "_"
        End If
    Next
    ReDim Preserve g_recConnection(m_intConnectionCount + 1)
    g_recConnection(m_intConnectionCount).strName = m_strCurrentConnection
    m_intConnectionCount = m_intConnectionCount + 1
End Sub

Private Sub ProcessSchemaError(strError As String)
'
'   Display an error message and terminate
'
    MsgBox strError & " at XML schema line# " & CStr(m_lngCount) & vbCrLf & vbCrLf & _
        "The schema line was:" & vbCrLf & m_strLine, vbCritical
    End
End Sub

Private Function ProcessSchemaLine() As Boolean
'
'   Read and pre-process a line read from the schema file
'
    Dim blnProcess As Boolean
    Dim intX As Integer
    
    Line Input #m_intFileNoXML, m_strLine
    m_strLine = Replace(m_strLine, vbTab, " ")
    m_lngCount = m_lngCount + 1
    If InStr(m_strLine, "<!--") <> 0 Then
        m_blnComment = True
    End If
    For intX = 1 To Len(m_strLine)
        If Mid$(m_strLine, intX, 1) <> " " Then
            Exit For
        End If
    Next
    blnProcess = Not m_blnComment
    If blnProcess Then
        If intX > Len(m_strLine) Then
            blnProcess = False
        ElseIf InStr("';", Mid$(m_strLine, intX, 1)) > 0 Then
            blnProcess = False
        End If
    End If
    If m_blnComment Then
        If InStr(m_strLine, "-->") <> 0 Then
            m_blnComment = False
        End If
    End If
    ProcessSchemaLine = blnProcess
End Function

Private Function ProcessSchemaField() As Boolean
'
'   Process the field data type value
'
    Dim blnResult As Boolean
    
    blnResult = True
    Select Case LCase$(g_recField(m_intFieldCount).strType)
        Case Is = "boolean"
            g_recField(m_intFieldCount).lngLength = 2
            g_recField(m_intFieldCount).strTypeADO = "adTinyInt"
        Case Is = "byte"
            g_recField(m_intFieldCount).lngLength = 2
            g_recField(m_intFieldCount).strTypeADO = "adByte"
        Case Is = "currency"
            g_recField(m_intFieldCount).lngLength = 16
            g_recField(m_intFieldCount).strTypeADO = "adLongInt"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).strType = "Long"
            End If
        Case Is = "date"
            g_recField(m_intFieldCount).lngLength = 16
            g_recField(m_intFieldCount).strTypeADO = "adDBTimeStamp"
        Case Is = "double"
            g_recField(m_intFieldCount).lngLength = 16
            g_recField(m_intFieldCount).strTypeADO = "adDouble"
        Case Is = "integer"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).lngLength = 8
                g_recField(m_intFieldCount).strTypeADO = "adInteger"
            Else
                g_recField(m_intFieldCount).lngLength = 4
                g_recField(m_intFieldCount).strTypeADO = "adSmallInt"
            End If
        Case Is = "int8"
            g_recField(m_intFieldCount).lngLength = 2
            g_recField(m_intFieldCount).strTypeADO = "adTinyInt"
            g_recField(m_intFieldCount).strType = "Byte"
        Case Is = "int16"
            g_recField(m_intFieldCount).lngLength = 4
            g_recField(m_intFieldCount).strTypeADO = "adSmallInt"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).strType = "Short"
            Else
                g_recField(m_intFieldCount).strType = "Integer"
            End If
        Case Is = "int32"
            g_recField(m_intFieldCount).lngLength = 8
            g_recField(m_intFieldCount).strTypeADO = "adInteger"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).strType = "Integer"
            Else
                g_recField(m_intFieldCount).strType = "Long"
            End If
        Case Is = "int64"
            g_recField(m_intFieldCount).lngLength = 16
            g_recField(m_intFieldCount).strTypeADO = "adLongInt"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).strType = "Long"
            Else
                g_recField(m_intFieldCount).strType = "Currency"
            End If
        Case Is = "long"
            If m_blnVBNet Then
                g_recField(m_intFieldCount).lngLength = 16
                g_recField(m_intFieldCount).strTypeADO = "adLongInt"
            Else
                g_recField(m_intFieldCount).lngLength = 8
                g_recField(m_intFieldCount).strTypeADO = "adInteger"
            End If
        Case Is = "memo"
            g_recField(m_intFieldCount).strTypeADO = "adLongVarChar"
        Case Is = "short"
            g_recField(m_intFieldCount).lngLength = 4
            g_recField(m_intFieldCount).strTypeADO = "adSmallInt"
            If Not m_blnVBNet Then
                g_recField(m_intFieldCount).strType = "Integer"
            End If
        Case Is = "single"
            g_recField(m_intFieldCount).lngLength = 8
            g_recField(m_intFieldCount).strTypeADO = "adSingle"
        Case Is = "string"
            g_recField(m_intFieldCount).strTypeADO = "adVarChar"
        Case Else
            blnResult = False
    End Select
    ProcessSchemaField = blnResult
End Function

Private Sub ProcessSchemaOption()
'
'   Scans and parses the Option line
'
    Dim intX As Integer
    Dim strLine As String
    Dim strOption As String
        
    intX = InStr(LCase$(m_strLine), enumSchemaOption)
    If intX <> 0 Then
        strLine = LCase$(Trim$(Mid$(m_strLine, intX + Len(enumSchemaOption))))
        If Right$(strLine, 1) = ">" Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        m_strArray = Split(strLine, " ")
        For intX = 0 To UBound(m_strArray)
            strOption = LCase$(m_strArray(intX))
            If strOption <> "" Then
                If strOption <> enumOptionAddnew _
                    And strOption <> enumOptionClear _
                    And strOption <> enumOptionCompactRepair _
                    And strOption <> enumOptionDelete _
                    And strOption <> enumOptionDeleteFrom _
                    And strOption <> enumOptionInsert _
                    And strOption <> enumOptionParameter _
                    And strOption <> enumOptionRecordCount _
                    And strOption <> enumOptionSearch _
                    And strOption <> enumOptionTransaction _
                    And strOption <> enumOptionUpdate _
                    And strOption <> enumOptionUpdateInto _
                    And strOption <> enumOptionVB6 _
                    And strOption <> enumOptionVBA _
                    And strOption <> enumOptionVBNet Then
                    Call ProcessSchemaError("An invalid keyword was included in the Option statement - " & m_strArray(intX))
                End If
                If strOption = enumOptionVB6 Then
                    m_blnVBNet = False
                ElseIf strOption = enumOptionVBNet Then
                    m_blnVBNet = True
                Else
                    ReDim Preserve m_strOptionArray(UBound(m_strOptionArray) + 1)
                    m_strOptionArray(UBound(m_strOptionArray) - 1) = strOption
                End If
            End If
        Next
    End If
End Sub

Private Sub ProcessSchemaOutput()
'
'   Scans and parses the Output line
'
    Dim intX As Integer
    Dim strLine As String
        
    intX = InStr(LCase$(m_strLine), enumSchemaOutput)
    If intX <> 0 Then
        strLine = Trim$(Mid$(m_strLine, intX + Len(enumSchemaOption)))
        If Right$(strLine, 1) = ">" Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        g_strModuleFile = strLine
    End If
End Sub

Private Sub ProcessSchemaQuery(intQueryType As Integer)
'
'   Scans and processes a query definition
'
    Dim blnOrderBy As Boolean
    Dim blnSet As Boolean
    Dim intFieldIndex As Integer
    Dim intTableIndex As Integer
    Dim intX As Integer
    Dim strLine As String
    Dim strQueryType As String
        
    Select Case intQueryType
        Case Is = enumQueryType.DeleteQuery
            strQueryType = enumSchemaQueryDelete
        Case Is = enumQueryType.SelectQuery
            strQueryType = enumSchemaQuerySelect
        Case Is = enumQueryType.SetQuery
            strQueryType = enumSchemaQuerySet
        Case Is = enumQueryType.UpdateQuery
            strQueryType = enumSchemaQueryUpdate
    End Select
    intX = InStr(LCase$(m_strLine), strQueryType)
    If intX <> 0 Then
        strLine = Trim$(Mid$(m_strLine, intX + Len(strQueryType)))
        If Right$(strLine, 1) = ">" Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        If strLine = "" Then
            Call ProcessSchemaError("A query line was read from the XML schema file with no following table name")
        End If
        ReDim Preserve g_recQuery(m_intQueryCount + 1)
        m_strArray = Split(strLine, " ")
        With g_recQuery(m_intQueryCount)
            If UBound(m_strArray) <= 0 Then
                Call ProcessSchemaError("A query line was read with no following query name")
            End If
            .intQueryType = intQueryType
            .strQueryConnection = m_strCurrentQueryConnection
            .strQueryTableExternal = m_strArray(0)
            intX = 0
            Do While intX < UBound(m_strArray)
                intX = intX + 1
                If m_strArray(intX) <> "" Then
                    .strQueryName = m_strArray(intX)
                    Exit Do
                End If
            Loop
            Do While intX < UBound(m_strArray)
                intX = intX + 1
                If m_strArray(intX) <> "" Then
                    If IsNumeric(m_strArray(intX)) Then
                        .intQueryTableBuffer = CInt(m_strArray(intX))
                        Exit Do
                    End If
                End If
            Loop
            For intTableIndex = 0 To m_intTableCount - 1
                If LCase$(.strQueryTableExternal) = LCase$(g_recTable(intTableIndex).strNameExternal) _
                    And LCase$(.strQueryConnection) = LCase$(g_recTable(intTableIndex).strConnection) Then
                    Exit For
                End If
            Next
            If intTableIndex >= m_intTableCount Then
                Call ProcessSchemaError("A query line specified a non-existant table - " & .strQueryName & "/" & .strQueryTableExternal)
            End If
            If .intQueryTableBuffer < 2 Then
                .intQueryTableBuffer = 1
            End If
            If .intQueryTableBuffer > g_recTable(intTableIndex).intBuffers Then
                Call ProcessSchemaError("A query line specified an invalid table buffer - " & .strQueryName & "/" & .strQueryTableExternal)
            End If
            .strQueryTableInternal = g_recTable(intTableIndex).strNameInternal
            .intQueryLineCount = 0
            .intQueryLinePtr = m_intQueryLineCount
            Do Until EOF(m_intFileNoXML)
                If ProcessSchemaLine Then
                    If InStr(LCase$(m_strLine), enumSchemaQueryEnd) Then
                        Exit Do
                    End If
                    strLine = Trim$(m_strLine)
                    Call ParseStringToArray(strLine, m_strArray, " ")
                    ReDim Preserve g_recQueryLine(m_intQueryLineCount + 1)
                    With g_recQueryLine(m_intQueryLineCount)
                        If m_strArray(0) = "(" Or m_strArray(0) = ")" Then
                            .intQueryLineType = enumQueryLine.Parenthesis
                            .strFieldName = m_strArray(0)
                        ElseIf LCase$(m_strArray(0)) = "and" Or LCase$(m_strArray(0)) = "or" Then
                            .intQueryLineType = enumQueryLine.Conjunction
                            .strFieldName = UCase$(m_strArray(0))
                        ElseIf LCase$(m_strArray(0)) = "orderby" Or LCase$(m_strArray(0)) = "order" Then
                            If g_recQuery(m_intQueryCount).intQueryType <> enumQueryType.SelectQuery Then
                                Call ProcessSchemaError("ORDER BY clause was specified on a non-select query - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                            Else
                                .strFieldName = ""
                                blnOrderBy = True
                            End If
                        ElseIf LCase$(m_strArray(0)) = "set" Then
                            If g_recQuery(m_intQueryCount).intQueryType <> enumQueryType.SetQuery Then
                                Call ProcessSchemaError("SET clause was specified on a non-set query - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                            Else
                                .strFieldName = ""
                                blnSet = True
                            End If
                        Else
                            .strFieldName = m_strArray(0)
                            If blnOrderBy Then
                                .intQueryLineType = enumQueryLine.OrderByField
                                If UBound(m_strArray) >= 1 Then
                                    .strFieldValue = UCase$(m_strArray(1))
                                End If
                            Else
                                If UBound(m_strArray) < 2 Then
                                    Call ProcessSchemaError("Invalid query line configuration in query - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                                End If
                                .strOperator = m_strArray(1)
                                If .strOperator <> "=" _
                                    And .strOperator <> "<>" _
                                    And .strOperator <> "!=" _
                                    And .strOperator <> ">" _
                                    And .strOperator <> "<" _
                                    And .strOperator <> ">=" _
                                    And .strOperator <> "<=" _
                                    And LCase$(.strOperator) <> "like" Then
                                    Call ProcessSchemaError("An invalid query line operator was specified on - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                                End If
                                If m_strArray(2) = "%" Then
                                    .intQueryLineType = IIf(blnSet, enumQueryLine.SetField, enumQueryLine.QueryField)
                                    If UBound(m_strArray) >= 3 Then
                                        For intX = 3 To UBound(m_strArray)
                                            .strFormat = .strFormat & m_strArray(intX)
                                        Next
                                    End If
                                ElseIf m_strArray(2) = "" Then
                                    Call ProcessSchemaError("An invalid query line operand was specified on - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                                Else
                                    .intQueryLineType = IIf(blnSet, enumQueryLine.SetFieldFixed, enumQueryLine.QueryFieldFixed)
                                    .strFieldValue = m_strArray(2)
                                End If
                            End If
                            For intFieldIndex = g_recTable(intTableIndex).intFieldPtr To g_recTable(intTableIndex).intFieldPtr + g_recTable(intTableIndex).intFieldCount - 1
                                If LCase$(.strFieldName) = LCase$(g_recField(intFieldIndex).strNameExternal) Then
                                    .lngFieldLength = g_recField(intFieldIndex).lngLength
                                    .strFieldType = g_recField(intFieldIndex).strType
                                    Exit For
                                End If
                            Next
                            If .strFieldType = "" Then
                                Call ProcessSchemaError("An invalid query field for a table was specified - " & g_recQuery(m_intQueryCount).strQueryName & "/" & g_recQuery(m_intQueryCount).strQueryTableExternal)
                            End If
                        End If
                    End With
                    If .intQueryLineCount = 0 Then
                        .intQueryLinePtr = m_intQueryLineCount
                    End If
                    If g_recQueryLine(m_intQueryLineCount).strFieldName <> "" Then
                        .intQueryLineCount = .intQueryLineCount + 1
                        m_intQueryLineCount = m_intQueryLineCount + 1
                    End If
                End If
            Loop
        End With
        m_intQueryCount = m_intQueryCount + 1
    End If
End Sub

Private Sub ProcessSchemaQueryConnection()
    '
    '   Assigns a series of queries to a connection
    '
    Dim intX As Integer
    Dim strLine As String
        
    intX = InStr(LCase$(m_strLine), enumSchemaQueryConnection)
    If intX > 0 Then
        strLine = Trim$(Mid$(m_strLine, intX + Len(enumSchemaQueryConnection)))
        If Right$(strLine, 1) = ">" Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        If strLine = "" Then
            Call ProcessSchemaError("Query connection line contains no connection name value")
        End If
        Call ParseStringToArray(strLine, m_strArray, " ")
        m_strCurrentQueryConnection = m_strArray(0)
    Else
        Exit Sub
    End If
    For intX = 1 To Len(m_strCurrentQueryConnection)
        If InStr("abcdefghijklmnopqrstuvwxyz0123456789_", LCase$(Mid$(m_strCurrentConnection, intX, 1))) = 0 Then
            Mid$(m_strCurrentQueryConnection, intX, 1) = "_"
        End If
    Next
    For intX = 0 To m_intConnectionCount - 1
        If LCase$(m_strCurrentQueryConnection) = LCase$(g_recConnection(intX).strName) Then
            Exit For
        End If
    Next
    If intX >= m_intConnectionCount Then
        Call ProcessSchemaError("Query connection name was not previously defined - " & m_strCurrentQueryConnection)
    End If
End Sub

Private Sub ProcessSchemaTable()
    '
    '   Scans and parses a table definition
    '
    Dim blnFound As Boolean
    Dim intX As Integer
    Dim strValue As String
    Dim strLength As String
    Dim strLine As String
    
    intX = InStr(LCase$(m_strLine), enumSchemaTable)
    ReDim Preserve g_recTable(m_intTableCount + 1)
    strLine = Trim$(Mid$(m_strLine, intX + Len(enumSchemaTable)))
    If Right$(strLine, 1) = ">" Then
        strLine = Left$(strLine, Len(strLine) - 1)
    End If
    m_strArray = Split(strLine, " ")
    If UBound(m_strArray) < 0 Then
        Call ProcessSchemaError("A '" & enumSchemaTable & "' line was read with no following table name")
    End If
    With g_recTable(m_intTableCount)
        .strConnection = m_strCurrentConnection
        .strNameExternal = m_strArray(0)
        .strNameInternal = IIf(m_strCurrentConnection = "", "", m_strCurrentConnection & "_") & m_strArray(0)
        For intX = 1 To Len(.strNameInternal)
            If InStr("abcdefghijklmnopqrstuvwxyz0123456789_", LCase$(Mid$(.strNameInternal, intX, 1))) = 0 Then
                Mid$(.strNameInternal, intX, 1) = "_"
            End If
        Next
        .intBuffers = 0
        .blnCreateTable = False
        .blnNoAdd = False
        .blnNoModify = False
        .blnNoDelete = False
        
        If UBound(m_strArray) > 0 Then
            For intX = 1 To UBound(m_strArray)
                strValue = LCase$(m_strArray(intX))
                Select Case strValue
                    Case Is = "create", "createtable"
                        .blnCreateTable = True
                    Case Is = "nowrite"
                        .blnNoAdd = True
                        .blnNoModify = True
                        .blnNoDelete = True
                    Case Is = "noadd"
                        .blnNoAdd = True
                    Case Is = "nodelete"
                        .blnNoDelete = True
                    Case Is = "nomodify"
                        .blnNoModify = True
                    Case Else
                        If Val(m_strArray(intX)) > 1 Then
                            .intBuffers = Val(m_strArray(intX))
                        End If
                End Select
            Next
        End If
        If .intBuffers < 1 Then
            .intBuffers = 1
        End If
        .intFieldCount = 0
        .intFieldPtr = m_intFieldCount
        blnFound = False
        Do Until EOF(m_intFileNoXML)
            If ProcessSchemaLine Then
                ReDim Preserve g_recField(m_intFieldCount + 1)
                If InStr(m_strLine, enumSchemaTableEnd) > 0 Then
                    Exit Do
                End If
                For intX = 1 To Len(m_strLine)
                    If InStr("abcdefghijklmnopqrstuvwxyz", LCase$(Mid$(m_strLine, intX, 1))) <> 0 Then
                        Exit For
                    End If
                Next
                strLine = Mid$(m_strLine, intX)
                If strLine <> "" Then
                    With g_recField(m_intFieldCount)
                        Call ParseStringToArray(strLine, m_strArray, " ")
                        If UBound(m_strArray) < 2 Then
                            Call ProcessSchemaError("Incorrect line formatting")
                        End If
                        .blnAutoKey = False
                        .blnIndex = False
                        .blnPrimary = False
                        .blnUnique = False
                        .lngLength = 0
                        .strNameExternal = m_strArray(0)
                        .strNameInternal = m_strArray(0)
                        For intX = 1 To Len(.strNameInternal)
                            If InStr("abcdefghijklmnopqrstuvwxyz0123456789_", LCase$(Mid$(.strNameInternal, intX, 1))) = 0 Then
                                Mid$(.strNameInternal, intX, 1) = "_"
                            End If
                        Next
                        strLength = ""
                        intX = InStr(m_strArray(1), "(")
                        If intX > 0 Then
                            strLength = Mid$(m_strArray(1), intX)
                            m_strArray(1) = Left$(m_strArray(1), intX - 1)
                        Else
                            If UBound(m_strArray) >= 2 Then
                                If Left$(m_strArray(2), 1) = "(" Then
                                    strLength = m_strArray(2)
                                End If
                            End If
                        End If
                        If Left$(strLength, 1) = "(" Then
                            intX = InStr(2, strLength, ")")
                            If intX = 0 Then
                                Call ProcessSchemaError("An invalid field length '" & strLength & "' was specified on the field '" & m_strArray(1) & "'")
                            Else
                                .lngLength = Val(Mid$(strLength, 2, intX - 2))
                            End If
                        End If
                        .strType = UCase$(Left$(m_strArray(1), 1)) & LCase$(Mid$(m_strArray(1), 2))
                        If Not ProcessSchemaField() Then
                            Call ProcessSchemaError("An invalid data type '" & .strType & "' was specified on the field '" & .strNameExternal & "'")
                        End If
                        If UBound(m_strArray) > 2 Then
                            For intX = 2 To UBound(m_strArray) - 1
                                If LCase$(m_strArray(intX)) = "autokey" Then
                                    .blnAutoKey = True
                                End If
                                If LCase$(m_strArray(intX)) = "primary" Then
                                    .blnPrimary = True
                                End If
                                If LCase$(m_strArray(intX)) = "index" Then
                                    .blnIndex = True
                                End If
                                If LCase$(m_strArray(intX)) = "unique" Then
                                    .blnUnique = True
                                End If
                            Next
                        End If
                        m_intFieldCount = m_intFieldCount + 1
                    End With
                    .intFieldCount = .intFieldCount + 1
                End If
            End If
        Loop
        m_intTableCount = m_intTableCount + 1
    End With
End Sub
