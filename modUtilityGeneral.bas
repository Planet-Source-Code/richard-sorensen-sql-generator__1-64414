Attribute VB_Name = "modUtilityGeneral"
Option Explicit
Option Base 0
Option Compare Text

'
'   Utility function library
'
Public Const vbKeyLF As Integer = 10

Public g_strAPIErrorMessage As String

Private m_lngBuffer As Long
Private m_lngHandle As Long
Private m_lngLength As Long
Private m_lngResult As Long
Private m_lngShift(31) As Long
Private m_strBuffer As String

Public Sub UtilityInitialize()
    On Error Resume Next
    m_strBuffer = Space$(4096)
    m_lngShift(0) = &H1&
    m_lngShift(1) = &H2&
    m_lngShift(2) = &H4&
    m_lngShift(3) = &H8&
    m_lngShift(4) = &H10&
    m_lngShift(5) = &H20&
    m_lngShift(6) = &H40&
    m_lngShift(7) = &H80&
    m_lngShift(8) = &H100&
    m_lngShift(9) = &H200&
    m_lngShift(10) = &H400&
    m_lngShift(11) = &H800&
    m_lngShift(12) = &H1000&
    m_lngShift(13) = &H2000&
    m_lngShift(14) = &H4000&
    m_lngShift(15) = &H8000&
    m_lngShift(16) = &H10000
    m_lngShift(17) = &H20000
    m_lngShift(18) = &H40000
    m_lngShift(19) = &H80000
    m_lngShift(20) = &H100000
    m_lngShift(21) = &H200000
    m_lngShift(22) = &H400000
    m_lngShift(23) = &H800000
    m_lngShift(24) = &H1000000
    m_lngShift(25) = &H2000000
    m_lngShift(26) = &H4000000
    m_lngShift(27) = &H8000000
    m_lngShift(28) = &H10000000
    m_lngShift(29) = &H20000000
    m_lngShift(30) = &H40000000
    m_lngShift(31) = &H80000000
End Sub

Public Sub ArrayClear(strArray() As String)
'
'   Clear a string array
'
    On Error Resume Next
    ReDim strArray(0)
    Err.Clear
End Sub

Public Sub ArrayInsert(strArray() As String, strValue As String)
'
'   Add an entry at the end of a string array
'
    On Error Resume Next
    ReDim Preserve strArray(UBound(strArray) + 1)
    strArray(UBound(strArray) - 1) = strValue
    Err.Clear
End Sub

Public Function BrowserLaunch(lngHWnd As Long, strURL As String) As Boolean
'
'   Launches the default browser and navigates to a URL
'
    Dim intFileNo As Integer
    Dim lngReturn As Long
    Dim strDummy As String
    Dim strEXEName As String
    Dim strFileName As String
    
    On Error Resume Next
    strFileName = SetSlash(GetDirectoryTemp()) & "Temp.htm"
    strEXEName = Space$(255)
    intFileNo = FreeFile
    Open strFileName For Output As #intFileNo
    Print #intFileNo, "<HTML></HTML>"
    Close #intFileNo
    lngReturn = FindExecutable(strFileName, strDummy, strEXEName)
    strEXEName = Trim$(strEXEName)
    Kill (strFileName)
    Err.Clear
    If lngReturn <= 32 Or IsEmpty(strEXEName) Then
        Exit Function
    End If
    lngReturn = ShellExecute(lngHWnd, "open", strEXEName, strURL, strDummy, SW_SHOWNORMAL)
    If lngReturn <= 32 Then
        Exit Function
    End If
    BrowserLaunch = True
    Err.Clear
End Function

Public Function ConvertStringToDate(strDate As String) As Date
'
'   Set a date value to a string, and use the current date if string is N/G
'
    Dim datReturn As Date
    
    On Error Resume Next
    datReturn = CDate(strDate)
    If Err Then
        datReturn = Now
        Err.Clear
    End If
    ConvertStringToDate = datReturn
End Function

Public Function ConvertTimeForDisplay(lngMillisecs As Long) As String
'
'   Return a string formatted as a time of day from a count of seconds
'
    Const lngStatusInterval As Long = 21600
    
    Dim lngHours As Long
    Dim lngMinutes As Long
    Dim lngSeconds As Long
    
    On Error Resume Next
    lngSeconds = lngMillisecs \ 1000
    Do While lngSeconds < 0
        lngSeconds = lngSeconds + lngStatusInterval
    Loop
    lngMinutes = lngSeconds \ 60
    lngSeconds = lngSeconds Mod 60
    lngHours = lngMinutes \ 60
    lngMinutes = lngMinutes Mod 60
    Err.Clear
    ConvertTimeForDisplay = Format$(lngHours, "####0") & ":" & Format$(lngMinutes, "00") & ":" & Format$(lngSeconds, "00")
End Function

Public Function ConvertUnixEpochToDate(dblEpoch As Double) As Date
'
'   Convert a unix-style epoch date to a VB date
'
    On Error Resume Next
    ConvertUnixEpochToDate = DateAdd("s", dblEpoch, #1/1/1970#)
    Err.Clear
End Function

Public Function CopyFile(strFileIn As String, strFileOut As String) As Boolean
'
'   File copy function
'
    On Error Resume Next
    FileCopy strFileIn, strFileOut
    If Err Then
        Err.Clear
    Else
        CopyFile = True
    End If
End Function

Public Function CreateDirectory(strPath As String) As Boolean
'
'   Creates a new directory
'
    On Error Resume Next
    MkDir (strPath)
    If Err Then
        CreateDirectory = False
    Else
        CreateDirectory = True
    End If
    Err.Clear
End Function

Public Function DeleteDirectory(strPath As String) As Boolean
'
'   Deletes a directory
'
    On Error Resume Next
    RmDir (strPath)
    If Err Then
        DeleteDirectory = False
    Else
        DeleteDirectory = True
    End If
    Err.Clear
End Function

Public Function DeleteFile(strFileName As String) As Boolean
'
'   Deletes a file
'
    On Error Resume Next
    SetAttr strFileName, vbNormal
    Err.Clear
    Kill (strFileName)
    If Err Then
        DeleteFile = False
    Else
        DeleteFile = True
    End If
    Err.Clear
End Function

Public Function DirectoryExists(strDirectoryName As String) As Boolean
'
'   Checks for the existance of a directory
'
    On Error Resume Next
    If strDirectoryName = "" Then
        Exit Function
    End If
    If Dir(SetSlash(strDirectoryName) & "*.*", vbDirectory + vbSystem) <> "" Then
        DirectoryExists = True
    End If
    Err.Clear
End Function

Public Function EncryptDecryptString(strPassPhrase As String, strKey As String) As String
'
'   Alternativly encrypts and decrypts a string value using the supplied key
'
    Dim intLength As Integer
    Dim intX As Integer
    Dim lngChar As Long
    Dim strResult As String
    
    On Error Resume Next
    If Len(strKey) < 1 Or Len(strPassPhrase) < 1 Then
        Exit Function
    End If
    intLength = Len(strKey)
    strResult = strPassPhrase
    For intX = 1 To Len(strPassPhrase)
        lngChar = Asc(Mid$(strKey, (intX Mod intLength) - intLength * ((intX Mod intLength) = 0), 1))
        Mid$(strResult, intX, 1) = Chr$(Asc(Mid$(strResult, intX, 1)) Xor lngChar)
    Next
    EncryptDecryptString = strResult
    Err.Clear
End Function

Public Function EscapeString(strValue As String, Optional strQuoteChar As String = "") As String
'
'   Perform URL-type escape operation on string
'
    Dim lngX As Long
    Dim strChar As String
    Dim strString As String
    
    On Error Resume Next
    For lngX = 1 To Len(strValue)
        strChar = Mid$(strValue, lngX, 1)
        If InStr("abcdefghijklmnopqrstuvwxyz0123456789*@-_./", LCase$(strChar)) > 0 Then
            strString = strString & strChar
        Else
            strChar = Hex(Asc(strChar))
            If Len(strChar) < 2 Then
                strChar = "0" & strChar
            End If
            strString = strString & "%" & strChar
        End If
    Next
    Err.Clear
    EscapeString = strQuoteChar & strString & strQuoteChar
End Function

Public Function FileExists(strFileName As String) As Boolean
'
'   Checks for the existence of a file
'
    On Error Resume Next
    m_lngResult = GetFileAttributes(strFileName)
    FileExists = IIf(m_lngResult = -1, False, True)
    Err.Clear
End Function

Public Function FileFilter(strFileName As String, strFilterString As String) As Boolean
'
'   Checks to see if a file name will pass a file filter
'
    Dim blnResult As Boolean
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim intFile As Integer
    Dim intFilter As Integer
    Dim strFilter As String
    
    On Error Resume Next
    If Trim$(strFilterString) = "" Then
        FileFilter = True
        Exit Function
    End If
    intBegin = 1
    Do
        blnResult = True
        intEnd = InStr(intBegin, strFilterString, "|")
        If intEnd > 0 Then
            strFilter = Trim$(Mid$(strFilterString, intBegin, intEnd - intBegin))
            intBegin = intEnd + 1
        Else
            strFilter = Mid$(strFilterString, intBegin)
            intBegin = Len(strFilter) + 1
        End If
        If strFilter = "" Then
            blnResult = False
            Exit Do
        End If
        intFile = 1
        intFilter = 1
        blnResult = True
        Do Until intFilter > Len(strFilter)
            If intFile > Len(strFileName) Then
                blnResult = False
                Exit Do
            End If
            If Mid$(strFilter, intFilter, 1) = "*" Then
                Do
                    intFilter = intFilter + 1
                    If intFilter > Len(strFilter) Then
                        Exit Do
                    End If
                Loop Until Mid$(strFilter, intFilter, 1) <> "*"
                If intFilter > Len(strFilter) Then
                    Exit Do
                End If
                For intFile = intFile To Len(strFileName)
                    If Mid$(strFileName, intFile, 1) = Mid$(strFilter, intFilter, 1) Then
                        Exit For
                    End If
                Next
                If intFile > Len(strFileName) Then
                    blnResult = False
                    Exit Do
                End If
            End If
            If Mid$(strFilter, intFilter, 1) <> "?" Then
                If Mid$(strFilter, intFilter, 1) <> Mid$(strFileName, intFile, 1) Then
                    blnResult = False
                    Exit Do
                End If
            End If
            intFile = intFile + 1
            intFilter = intFilter + 1
        Loop
        If blnResult Then
            Exit Do
        End If
    Loop
    FileFilter = blnResult
    Err.Clear
End Function

Public Function GetAddressOf(ByVal lngPtr As Long) As Long
'
'   AddressOf hack
'
    GetAddressOf = lngPtr
End Function

Public Function GetAPIErrorString(dwErrCode As Long) As String
'
'   Returns the system-defined description of an API error code
'
    Dim sErrDesc As String * 256   ' max string resource len
    
    On Error Resume Next
    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS Or _
        FORMAT_MESSAGE_MAX_WIDTH_MASK, _
        ByVal 0&, dwErrCode, LANG_USER_DEFAULT, _
        ByVal sErrDesc, 256, 0)
    GetAPIErrorString = GetStrFromBufferA(sErrDesc)
End Function

Public Function GetArrayEntry(strArray() As String, intIndex As Integer) As String
'
'   Return the contents of an array entry, if within bounds
'
    On Error Resume Next
    If intIndex <= UBound(strArray) Then
        GetArrayEntry = strArray(intIndex)
    End If
    Err.Clear
End Function

Public Function GetCursorPosition(ByRef lngX As Long, ByRef lngY As Long)
'
'   Returns the current mouse position
'
    Dim recPoint As typPOINT
    
    On Error Resume Next
    Call GetCursorPos(recPoint)
    lngX = recPoint.X
    lngY = recPoint.Y
End Function

Public Function GetDirectory(Optional strDrive As String = "") As String
'
'   Gets the current directory
'
    On Error Resume Next
    If strDrive = "" Then
        GetDirectory = CurDir$
    Else
        GetDirectory = CurDir$(strDrive)
    End If
    Err.Clear
End Function

Public Function GetDirectoryTemp() As String
'
'   Returns the Windows temp directory
'
    On Error Resume Next
    m_lngLength = GetTempPath(Len(m_strBuffer), m_strBuffer)
    GetDirectoryTemp = Left$(m_strBuffer, m_lngLength)
    Err.Clear
End Function

Public Function GetDirectoryWindows() As String
'
'   Returns the Windows base directory
'
    On Error Resume Next
    m_lngLength = GetWindowsDirectory(m_strBuffer, Len(m_strBuffer))
    GetDirectoryWindows = UCase$(Left$(m_strBuffer, m_lngLength))
    Err.Clear
End Function

Public Function GetDirectoryWindowsSystem() As String
'
'   Returns the Windows system directory
'
    On Error Resume Next
    m_lngLength = GetSystemDirectory(m_strBuffer, Len(m_strBuffer))
    GetDirectoryWindowsSystem = Left$(m_strBuffer, m_lngLength)
    Err.Clear
End Function

Public Function GetDrive() As String
'
'   Returns the drive letter of the current path
'
    Dim strPath As String
    
    On Error Resume Next
    strPath = CurDir$
    GetDrive = Left$(strPath, 2)
    Err.Clear
End Function

Public Function GetFileContents(strFileName As String) As String
'
'   Read in the entire contents of a file and return it as a string
'
    Dim intFileNo As Integer
    Dim strContents As String
    
    On Error Resume Next
    intFileNo = FreeFile
    Open strFileName For Binary Access Read As #intFileNo
    If Not Err Then
        strContents = Space$(FileLen(strFileName))
        Get #intFileNo, , strContents
        Close #intFileNo
        GetFileContents = strContents
        strContents = ""
    Else
        Err.Clear
    End If
End Function

Public Function GetFileExtension(strFileName As String) As String
'
'   Returns the file extension from a file name
'
    Dim intX As Integer
    Dim strName As String
    
    On Error Resume Next
    intX = InStrRev(strFileName, ".")
    If intX <= 0 Then
        strName = ""
    Else
        strName = Mid$(strFileName, intX + 1)
    End If
    Err.Clear
    GetFileExtension = LCase$(strName)
End Function

Public Function GetFileName(strFileName As String, Optional blnExcludeExtension As Boolean = False) As String
'
'   Strips the path and returns the filename, with or without extension
'
    Dim intX As Integer
    Dim strName As String
    
    On Error Resume Next
    intX = InStrRev(strFileName, "\")
    If intX <= 0 Then
        strName = strFileName
    Else
        strName = Mid$(strFileName, intX + 1)
    End If
    If blnExcludeExtension Then
        intX = InStr(strName, ".")
        If intX > 0 Then
            strName = Left$(strName, intX - 1)
        End If
    End If
    Err.Clear
    GetFileName = strName
End Function

Public Function GetFilePath(strFileName As String) As String
'
'   Strips the path from a file name and returns it
'
    Dim intX As Integer
    Dim strName As String
    
    On Error Resume Next
    intX = InStrRev(strFileName, "\")
    If intX <= 0 Then
        GetFilePath = ""
    Else
        GetFilePath = Left$(strFileName, intX - 1)
    End If
End Function

Public Function GetFileSize(strFileName As String) As Long
'
'   Return the file size
'
    On Error Resume Next
    m_lngLength = FileLen(strFileName)
    If Err.Number <> 0 Then
        m_lngLength = 0
    End If
    Err.Clear
    GetFileSize = m_lngLength
End Function

Public Function GetFileVersion(strFileName As String) As String
'
'   Returns the version string for an executable file
'
    Dim bytBuffer() As Byte
    Dim recFileVersionInfo As typFileVersionInfo
    Dim strVersion As String
    
    On Error Resume Next
    m_lngLength = GetFileVersionInfoSize(strFileName, m_lngHandle)
    If m_lngLength > LenB(recFileVersionInfo) Then
        ReDim bytBuffer(m_lngLength + 1)
        m_lngResult = GetFileVersionInfo(strFileName, m_lngHandle, m_lngLength, bytBuffer(0))
        m_lngResult = VerQueryValue(bytBuffer(0), "\", m_lngHandle, m_lngLength)
        If m_lngResult <> 0 Then
            Call CopyMemory(recFileVersionInfo, m_lngHandle, LenB(recFileVersionInfo))
        End If
        With recFileVersionInfo
            strVersion = CStr(CInt(.lngFileVersionMS / &H10000)) & "." & _
                CStr(CInt(.lngFileVersionMS And &HFFFF&)) & "." & _
                CStr(CInt(.lngFileVersionLS / &H10000)) & "." & _
                CStr(CInt(.lngFileVersionLS And &HFFFF&))
        End With
        Err.Clear
        GetFileVersion = strVersion
    End If
End Function

Public Function GetGUID(blnAddDelimiters As Boolean) As String
'
'   Returns a new GUID
'
    Dim udtGuid As typGuid
    
    On Error Resume Next
    If (CoCreateGuid(udtGuid) = 0) Then
        GetGUID = _
            IIf(blnAddDelimiters, "{", "") & _
            String$(8 - Len(Hex$(udtGuid.lngData1)), "0") & Hex$(udtGuid.lngData1) & _
            IIf(blnAddDelimiters, "-", "") & _
            String$(4 - Len(Hex$(udtGuid.intData2)), "0") & Hex$(udtGuid.intData2) & _
            IIf(blnAddDelimiters, "-", "") & _
            String$(4 - Len(Hex$(udtGuid.intData3)), "0") & Hex$(udtGuid.intData3) & _
            IIf(blnAddDelimiters, "-", "") & _
            IIf((udtGuid.bytData4(0) < &H10), "0", "") & Hex$(udtGuid.bytData4(0)) & _
            IIf((udtGuid.bytData4(1) < &H10), "0", "") & Hex$(udtGuid.bytData4(1)) & _
            IIf(blnAddDelimiters, "-", "") & _
            IIf((udtGuid.bytData4(2) < &H10), "0", "") & Hex$(udtGuid.bytData4(2)) & _
            IIf((udtGuid.bytData4(3) < &H10), "0", "") & Hex$(udtGuid.bytData4(3)) & _
            IIf((udtGuid.bytData4(4) < &H10), "0", "") & Hex$(udtGuid.bytData4(4)) & _
            IIf((udtGuid.bytData4(5) < &H10), "0", "") & Hex$(udtGuid.bytData4(5)) & _
            IIf((udtGuid.bytData4(6) < &H10), "0", "") & Hex$(udtGuid.bytData4(6)) & _
            IIf((udtGuid.bytData4(7) < &H10), "0", "") & Hex$(udtGuid.bytData4(7)) & _
            IIf(blnAddDelimiters, "}", "")
    End If
    Err.Clear
End Function

Public Function GetHighWord(dwValue As Long) As Integer
'
'   Returns the low 16-bit integer from a 32-bit long integer
'
    MoveMemory GetHighWord, ByVal VarPtr(dwValue) + 2, 2
End Function

Public Function GetINIString(strFileName As String, strSectionName As String, strStringName As String, strStringDefault As String) As String
'
'   Returns a string value from an INI file (or other text file formatted as an INI)
'
    On Error Resume Next
    m_lngLength = GetPrivateProfileString(strSectionName, strStringName, strStringDefault, m_strBuffer, Len(m_strBuffer), strFileName)
    GetINIString = Left$(m_strBuffer, m_lngLength)
    Err.Clear
End Function

Public Function GetINIValue(strFileName As String, strSectionName As String, strStringName As String, strStringDefault As String) As Variant
'
'   Returns a numeric or boolean value from an INI file (or other text file formatted as an INI)
'
    Dim strValue As String
    
    On Error Resume Next
    strValue = UCase$(GetINIString(strFileName, strSectionName, strStringName, strStringDefault))
    If strValue = "FALSE" Then
        GetINIValue = False
    ElseIf strValue = "TRUE" Then
        GetINIValue = True
    Else
        GetINIValue = Val(strValue)
    End If
    Err.Clear
End Function

Public Function GetInterval(ByRef lngIntervalLast As Long) As Long
'
'   Compute an interval in milliseconds since the last call
'
    Dim lngCurrent As Long
    Dim lngInterval As Long
    
    On Error Resume Next
    lngCurrent = GetTickCount
    If lngIntervalLast = 0 Then
        lngInterval = 0
    Else
        lngInterval = lngCurrent - lngIntervalLast
    End If
    lngIntervalLast = lngCurrent
    If lngInterval < 0 Then
        lngInterval = 0
    End If
    GetInterval = lngInterval
    Err.Clear
End Function

Public Function GetLocaleID() As Long
'
'   Returns the current locale ID set in the control panel
'
    On Error Resume Next
    GetLocaleID = GetThreadLocale
    Err.Clear
End Function

Public Function GetLowWord(dwValue As Long) As Integer
'
'   Returns the low 16-bit integer from a 32-bit long integer
'
    MoveMemory GetLowWord, dwValue, 2
End Function

Public Function GetNewFileName(strPrefix As String, Optional strPath As String = "", Optional strExtension As String = "") As String
'
'   Returns a unique file name from a selected directory
'
    Const strTempExtension As String = ".tmp"
    
    Dim lngNo As Long
    Dim strFileName As String
    Dim strPathName As String
    Dim strTestName As String
    
    On Error Resume Next
    If strPath = "" Then
        strPathName = GetDirectoryTemp()
    Else
        strPathName = strPath
    End If
    m_lngLength = GetTempFileName(strPathName, strPrefix, 0, m_strBuffer)
    m_lngLength = InStr(LCase$(m_strBuffer), strTempExtension)
    strFileName = Left$(m_strBuffer, m_lngLength + 3)
    Kill (strFileName)
    Err.Clear
    If Right$(strFileName, Len(strTempExtension)) = strTempExtension Then
        strFileName = Left$(strFileName, Len(strFileName) - Len(strTempExtension))
    End If
    lngNo = 0
    Do
        strTestName = strFileName & IIf(lngNo <= 0, "", CStr(lngNo)) & IIf(strExtension = "", "", "." & strExtension)
        If FileExists(strTestName) Then
            lngNo = lngNo + 1
        Else
            Exit Do
        End If
    Loop
    GetNewFileName = GetFileName(strFileName, IIf(strExtension = "", False, True)) & IIf(lngNo <= 0, "", CStr(lngNo)) & IIf(strExtension = "", "", "." & strExtension)
    Err.Clear
End Function

Public Function GetPlatform(blnFullVersion As Boolean) As String
'
'   Returns a value indicating the running O/S
'
    Dim intBuild As Integer
    Dim recVersion As typVersionInfo
    Dim strPlatform As String

    On Error Resume Next
    recVersion.lngBufferSize = Len(recVersion)
    m_lngResult = GetVersionEx(recVersion)
    If (recVersion.lngBuildNumber And &HFFFF&) > &H7FFF Then
        intBuild = (recVersion.lngBuildNumber And &HFFFF&) - &H10000
    Else
        intBuild = recVersion.lngBuildNumber And &HFFFF&
    End If
    strPlatform = ""
    Select Case recVersion.lngPlatformID
        Case Is = 0
            strPlatform = "Win31"
        Case Is = 1
            If recVersion.lngMinorVersion < 10 Then
                strPlatform = "Win95"
                If blnFullVersion Then
                    If intBuild > 950 Then
                        strPlatform = strPlatform & " B"
                    End If
                End If
            Else
                strPlatform = "Win98"
                If blnFullVersion Then
                    If recVersion.lngMinorVersion > 10 Then
                        strPlatform = strPlatform & " ME"
                    ElseIf intBuild > 1998 Then
                        strPlatform = strPlatform & " SE"
                    End If
                End If
            End If
        Case Is = 2
            Select Case recVersion.lngMajorVersion
                Case Is = 3
                    strPlatform = "WinNT"
                    If blnFullVersion Then
                        strPlatform = strPlatform & " 351"
                    End If
                Case Is = 4
                    strPlatform = "WinNT"
                    If blnFullVersion Then
                        strPlatform = strPlatform & " 40"
                    End If
                Case Is = 5
                    strPlatform = "Win2000"
            End Select
    End Select
    If blnFullVersion Then
        strPlatform = strPlatform & " " & CStr(recVersion.lngMajorVersion) & "." & CStr(recVersion.lngMinorVersion) & "." & CStr(intBuild)
    End If
    GetPlatform = strPlatform
    Err.Clear
End Function

Public Sub GetScreenResolution(lngHor As Long, lngVer As Long)
'
'   Returns the current screen resolution
'
    On Error Resume Next
    lngHor = GetSystemMetrics(SM_CXSCREEN)
    lngVer = GetSystemMetrics(SM_CYSCREEN)
    Err.Clear
End Sub

Public Function GetSerialHash(strSerialNo As String) As String
'
'   Accept a serial number and return a hash code
'
    Const strTable As String = "ABCDEFGHIJKLMNPS"
    
    Dim intX As Integer
    Dim intY As Integer
    Dim lngHash As Long
    Dim lngWord(4) As Long
    Dim strCode As String
    
    On Error Resume Next
    If Len(strSerialNo) <> 16 Or Not IsHex(strSerialNo) Then
        Exit Function
    End If
    For intX = 0 To 3
        intY = (intX * 4) + 1
        lngWord(intX) = CLng("&H" & Mid$(strSerialNo, intY, 4))
    Next
    lngHash = lngWord(0)
    For intX = 1 To 4
        lngHash = (((lngHash * 4) And 65532) Or ((lngHash And 65535) \ 16384))
        If intX < 4 Then
            lngHash = lngHash Xor lngWord(intX)
        End If
    Next
    strCode = ""
    For intX = 0 To 3
        strCode = strCode & Mid$(strTable, (lngHash And 15) + 1, 1)
        lngHash = lngHash \ 16
    Next
    GetSerialHash = strCode
    Err.Clear
End Function

Public Function GetShiftKeyState() As Boolean
'
'   Return the state of either shift key
'
    Dim intX As Integer
    Dim intY As Integer
        
    On Error Resume Next
    intX = GetKeyState(VK_LSHIFT) And &H80
    intY = GetKeyState(VK_RSHIFT) And &H80
    If intX > 0 Or intY > 0 Then
        GetShiftKeyState = True
    End If
End Function

Public Function GetShortFileName(strLongFileName As String) As String
'
'   Return the short (MS-DOS) name of a file
'
    On Error Resume Next
    m_lngLength = Len(m_strBuffer)
    m_lngResult = GetShortPathName(strLongFileName, m_strBuffer, m_lngLength)
    GetShortFileName = Left$(m_strBuffer, m_lngResult)
    Err.Clear
End Function

Public Function GetStrFromBufferA(sz As String) As String
'
'   Returns the string before first null char encountered (if any) from an ANSII string.
'
    On Error Resume Next
    If InStr(sz, vbNullChar) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
    Else
        ' If sz had no null char, the Left$ function
        ' above would return a zero length string ("").
        GetStrFromBufferA = sz
    End If
End Function

Public Function GetStrFromPtrA(lpszA As Long) As String
'
'   Returns an ANSII string from a pointer to an ANSII string.
'
    Dim sRtn As String
    
    On Error Resume Next
    sRtn = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal sRtn, ByVal lpszA)
    GetStrFromPtrA = sRtn
End Function

Public Function GetStrFromPtrW(lpszW As Long) As String
'
'   Returns an ANSI string from a pointer to a Unicode string.
'
    Dim sRtn As String
    
    On Error Resume Next
    sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
    '  sRtn = String$(WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, 0, 0, 0, 0), 0)
    Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
    GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function

Public Function GetTextLine(strBuffer As String, strLine As String) As Boolean
'
'   Extract a line from an input buffer - delimited by LF or CR/LF
'
    Dim intX As Integer
    
    On Error Resume Next
    intX = InStr(strBuffer, vbLf)
    If intX > 0 Then
        strLine = Mid$(strBuffer, 1, intX - 1)
        If Right$(strLine, 1) = vbCr Then
            strLine = Left$(strLine, Len(strLine) - 1)
        End If
        strBuffer = Mid$(strBuffer, intX + 1)
        GetTextLine = True
    Else
        strLine = strBuffer
        strBuffer = ""
        GetTextLine = False
    End If
    Err.Clear
End Function

Public Sub GetTimeZone(ByRef intTimeZone As Integer, ByRef lngOffset As Long, ByRef strTimeZone As String, ByRef strTimeZoneShort As String)
'
'   Returns the time zone name, value, and offset (from UTC/GMT) in minutes
'
    Dim intX As Integer
    Dim recTimeZone As typTimeZone
    Dim strZone As String
    
    On Error Resume Next
    m_lngResult = GetTimeZoneInformation(recTimeZone)
    strZone = ""
    If m_lngResult = TIME_ZONE_ID_DAYLIGHT Then
        For intX = 0 To UBound(recTimeZone.intDaylightName)
            If recTimeZone.intDaylightName(intX) = 0 Then
                Exit For
            End If
            strZone = strZone & Chr$(recTimeZone.intDaylightName(intX))
        Next
        lngOffset = recTimeZone.lngBias - 60
    Else
        For intX = 0 To UBound(recTimeZone.intStandardName)
            If recTimeZone.intStandardName(intX) = 0 Then
                Exit For
            End If
            strZone = strZone & Chr$(recTimeZone.intStandardName(intX))
        Next
        lngOffset = recTimeZone.lngBias
    End If
    intTimeZone = (lngOffset \ 60) * -1
    strTimeZone = strZone
    Select Case recTimeZone.lngBias
        Case Is = 0
            strTimeZoneShort = "GMT"
        Case Is = 60
            strTimeZoneShort = "AT"
        Case Is = 120
            strTimeZoneShort = "FST"
        Case Is = 180
            strTimeZoneShort = "BST"
        Case Is = 240
            strTimeZoneShort = "AST"
        Case Is = 300
            strTimeZoneShort = "EST"
        Case Is = 360
            strTimeZoneShort = "CST"
        Case Is = 420
            strTimeZoneShort = "MST"
        Case Is = 480
            strTimeZoneShort = "PST"
        Case Is = 540
            strTimeZoneShort = "AKST"
        Case Is = 600
            strTimeZoneShort = "HST"
        Case Is = 660
            strTimeZoneShort = "BEST"
        Case Is = 720
            strTimeZoneShort = "IDLW"
        Case Is = -60
            strTimeZoneShort = "CET"
        Case Is = -120
            strTimeZoneShort = "EET"
        Case Is = -180
            strTimeZoneShort = "MSK"
        Case Is = -240
            strTimeZoneShort = "GST"
        Case Is = -300
            strTimeZoneShort = "AQTT"
        Case Is = -360
            strTimeZoneShort = "ALMT"
        Case Is = -420
            strTimeZoneShort = "JT"
        Case Is = -480
            strTimeZoneShort = "AWST"
        Case Is = -540
            strTimeZoneShort = "JST"
        Case Is = -600
            strTimeZoneShort = "AEST"
        Case Is = -660
            strTimeZoneShort = "VUT"
        Case Is = -720
            strTimeZoneShort = "NZT"
    End Select
    strTimeZoneShort = strTimeZoneShort & " (" & CStr(lngOffset / -60) & ")"
    Err.Clear
End Sub

Public Function GetTopLevelParent(hWnd As Long) As Long
'
'   Returns the top level parent window from the specified window handle.
'
    Dim hwndParent As Long
    Dim hwndTmp As Long
    
    On Error Resume Next
    hwndParent = hWnd
    Do
        hwndTmp = GetParent(hwndParent)
        If hwndTmp Then
            hwndParent = hwndTmp
        End If
    Loop While hwndTmp
    GetTopLevelParent = hwndParent
End Function

Public Function GetWinUsername() As String
'
'   Returns the username of the current logged-in user
'
    On Error Resume Next
    m_lngLength = Len(m_strBuffer)
    m_lngResult = GetUsername(m_strBuffer, m_lngLength)
    If m_lngResult = 0 Then
        GetWinUsername = ""
    Else
        GetWinUsername = UCase$(Left$(m_strBuffer, m_lngLength - 1))
    End If
    Err.Clear
End Function

Public Function IsHex(strString As String) As Boolean
'
'   Determine if a string contains all hex characters
'
    Dim intX As Integer
    
    On Error Resume Next
    If Len(strString) > 0 Then
        For intX = 1 To Len(strString)
            If InStr("0123456789abcdef", LCase$(Mid$(strString, intX, 1))) = 0 Then
                Exit Function
            End If
        Next
        IsHex = True
    End If
    Err.Clear
End Function

Public Function IsIDE() As Boolean
'
'   Determine if VB is running in the IDE
'
    On Error GoTo Out
    Debug.Print 1 / 0
Out:
    IsIDE = Err
End Function

Public Function IsValidString(strString As String, blnLetters As Boolean, blnNumbers As Boolean, strOtherChars As String) As Boolean
'
'   Scan a string to determine if all of the characters are "valid"
'
    Dim blnOK As Boolean
    Dim lngX As Long
    Dim strCompare As String
    
    On Error Resume Next
    If blnLetters Then
        strCompare = "abcdefghijklmnopqrstuvwxyz"
    End If
    If blnNumbers Then
        strCompare = strCompare & "1234567890"
    End If
    strCompare = strCompare & strOtherChars
    blnOK = True
    For lngX = 1 To Len(strString)
        If InStr(strCompare, Mid$(strString, lngX, 1)) <= 0 Then
            blnOK = False
            Exit For
        End If
    Next
    IsValidString = blnOK
End Function

Public Function JustifyNumber(dblValue As Double, strFormat As String) As String
'
'   Accepts a numeric value and returns a string that is right-justified, using the format as the string length
'
    Dim strNumber As String
    
    On Error Resume Next
    strNumber = Format$(dblValue, strFormat)
    If Len(strNumber) < Len(strFormat) Then
        strNumber = Space$(Len(strFormat) - Len(strNumber)) & strNumber
    End If
    JustifyNumber = strNumber
End Function

Public Function JustifyString(strValue As String, intLength As Integer, Optional blnRightJustify As Boolean = False) As String
'
'   Accepts and returns a string that is justified in a fixed length area
'
    Dim strText As String
    
    On Error Resume Next
    strText = Trim$(strValue)
    If Len(strText) < intLength Then
        If blnRightJustify Then
            strText = Space$(intLength - Len(strText)) & strText
        Else
            strText = strText & Space$(intLength - Len(strText))
        End If
    Else
        strText = Left$(strText, intLength)
    End If
    JustifyString = strText
End Function

Public Function JustifyTrailer(strValue As String, intLength As Integer, Optional strTrailerChar As String = ".") As String
'
'   Accepts and returns a string that is left-justified in a fixed length area, and includes trailing dots
'
    Dim intX As Integer
    Dim strText As String
    
    On Error Resume Next
    strText = JustifyString(strValue, intLength)
    If Len(strText) > 0 Then
        For intX = Len(strText) To 1 Step -1
            If Mid$(strText, intX, 1) <> " " Then
                Exit For
            End If
        Next
        intX = Len(strText) - intX - 2
        If intX > 0 Then
            strText = Left$(strText, Len(strText) - intX - 1) & String$(intX, Left$(strTrailerChar, 1)) & " "
        End If
    End If
    JustifyTrailer = strText
End Function

Public Sub LockWindow(lngHWnd As Long)
'
'   Prevent a window/control from being visually updated
'
    On Error Resume Next
    Call LockWindowUpdate(lngHWnd)
    Err.Clear
End Sub

Public Function ParseStringToArray(strParseLine As String, strParseArray() As String, strSeparatorChar As String, Optional strQuoteChar As String = "", Optional blnBinaryScan As Boolean = False) As Boolean
'
'   Parses a line of text and returns an array of tokens
'
    Dim blnEndFlag As Boolean
    Dim blnResult As Boolean
    Dim intArrayIndex As Integer
    Dim lngEnd As Long
    Dim lngStart As Long
    Dim strQuote As String
    Dim strTab As String
    Dim strSeparator As String
    Dim strToken As String
    
    On Error Resume Next
    ReDim strParseArray(0)
    blnResult = True
    If strParseLine = "" Then
        ParseStringToArray = blnResult
        Exit Function
    End If
    If Len(strSeparatorChar) < 1 Then
        strSeparator = ","
    Else
        strSeparator = strSeparatorChar
    End If
    strQuote = strQuoteChar
    If strQuote <> "" And strQuote <> "'" And strQuote <> """" Then
        strQuote = """"
    End If
    strTab = Chr$(vbKeyTab)
    intArrayIndex = 0
    lngStart = 1
    Do
        strToken = ""
        Do
            If lngStart > Len(strParseLine) Then
                Exit Do
            End If
            If blnBinaryScan Or InStr(" " & strTab, Mid$(strParseLine, lngStart, 1)) = 0 Then
                Exit Do
            End If
            lngStart = lngStart + 1
        Loop
        Do
            If lngStart > Len(strParseLine) Then
                blnEndFlag = True
                Exit Do
            End If
            If strQuote <> "" And Mid$(strParseLine, lngStart, 1) = strQuote Then
                lngEnd = InStr(lngStart + 1, strParseLine, strQuote)
                If lngEnd = 0 Then
                    blnResult = False
                    blnEndFlag = True
                    Exit Do
                End If
                strToken = Mid$(strParseLine, lngStart + 1, lngEnd - lngStart - 1)
                lngStart = lngEnd
                lngEnd = InStr(lngStart + 1, strParseLine, strSeparator)
                If lngEnd = 0 Then
                    lngStart = lngStart + 1
                Else
                    lngStart = lngEnd + Len(strSeparator)
                End If
            Else
                If strSeparator = " " Or strSeparator = strTab Then
                    lngEnd = lngStart
                    Do
                        If lngEnd > Len(strParseLine) Then
                            lngEnd = 0
                            Exit Do
                        End If
                        If InStr(" " & strTab, Mid$(strParseLine, lngEnd, 1)) > 0 Then
                            Exit Do
                        End If
                        lngEnd = lngEnd + 1
                    Loop
                Else
                    lngEnd = InStr(lngStart, strParseLine, strSeparator)
                End If
                If lngEnd = 0 Then
                    strToken = Mid$(strParseLine, lngStart, Len(strParseLine) + 1 - lngStart)
                    lngStart = Len(strParseLine) + 1
                Else
                    strToken = Mid$(strParseLine, lngStart, lngEnd - lngStart)
                    lngStart = lngEnd + Len(strSeparator)
                End If
            End If
            Exit Do
        Loop
        If blnEndFlag Then
            Exit Do
        Else
            If Not blnBinaryScan Then
                Do
                    lngEnd = InStr(strToken, strTab)
                    If lngEnd <> 0 Then
                        strToken = Left$(strToken, lngEnd - 1) & Mid$(strToken, lngEnd + 1)
                    End If
                Loop Until lngEnd = 0
            End If
        End If
        intArrayIndex = intArrayIndex + 1
        ReDim Preserve strParseArray(intArrayIndex)
        strParseArray(intArrayIndex - 1) = strToken
    Loop
    ParseStringToArray = blnResult
    Err.Clear
End Function

Public Function ParseToken(strLine As String, strBeginChar As String, strEndChar As String) As String
'
'   Returns a substring between two separator characters (if found)
'
    Dim lngX As Long
    Dim lngY As Long
    
    On Error Resume Next
    If strBeginChar = "" Or strEndChar = "" Then
        Exit Function
    End If
    lngX = InStr(strLine, strBeginChar)
    If lngX > 0 Then
        lngY = InStr(lngX + Len(strBeginChar), strLine, strEndChar)
        If lngY > 0 Then
            ParseToken = Mid$(strLine, lngX + Len(strBeginChar), lngY - lngX - Len(strBeginChar))
        End If
    Else
        ParseToken = strLine
    End If
End Function

Public Sub PlaySound(strFileName As String)
'
'   Play a sound file
'
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    
    Dim lngFlags As Long
    
    On Error Resume Next
    lngFlags = SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP
    Call sndPlaySound(strFileName, lngFlags)
    Err.Clear
End Sub

Public Function PutFileContents(strFileName As String, strContents As String) As Boolean
'
'   Write a string out to a file
'
    Dim intFileNo As Integer
    
    On Error Resume Next
    intFileNo = FreeFile
    Open strFileName For Output As #intFileNo
    If Not Err Then
        Print #intFileNo, strContents;
        Close #intFileNo
        PutFileContents = True
    Else
        Err.Clear
    End If
End Function

Public Sub RegistryDeleteKey(lngStandardKey As Long, strKeyName As String)
'
'   Deletes a registry key
'
    On Error Resume Next
    If strKeyName = "" Then
        Exit Sub
    End If
    m_lngResult = RegDeleteKey(lngStandardKey, strKeyName)
    Err.Clear
End Sub

Public Sub RegistryDeleteValue(lngStandardKey As Long, strKeyName As String, strValueName As String)
'
'   Deletes a registry value
'
    On Error Resume Next
    If strKeyName = "" Or strValueName = "" Then
        Exit Sub
    End If
    If RegOpenKey(lngStandardKey, strKeyName, m_lngHandle) <> 0 Then
        Exit Sub
    End If
    m_lngResult = RegDeleteValue(m_lngHandle, strValueName)
    m_lngResult = RegCloseKey(m_lngHandle)
    Err.Clear
End Sub

Public Function RegistryReadValue(lngStandardKey As Long, strKeyName As String, Optional strValueName As String = "") As Variant
'
'   Reads a registry value
'
    Dim lngResultType As Long
    
    On Error Resume Next
    If strKeyName = "" Then
        Exit Function
    End If
    If strValueName = "" Then
        If RegOpenKey(lngStandardKey, strKeyName, m_lngHandle) <> 0 Then
            Exit Function
        End If
        m_lngLength = Len(m_strBuffer)
        m_lngResult = RegQueryValue(m_lngHandle, "", m_strBuffer, m_lngLength)
        lngResultType = REG_SZ
    Else
        If RegOpenKeyEx(lngStandardKey, strKeyName, 0, KEY_QUERY_VALUE, m_lngHandle) <> 0 Then
            Exit Function
        End If
        m_lngLength = 4
        m_lngResult = RegQueryValueExLong(m_lngHandle, strValueName, 0, lngResultType, m_lngBuffer, m_lngLength)
        If lngResultType <> REG_DWORD Then
            m_lngLength = Len(m_strBuffer)
            m_lngResult = RegQueryValueExString(m_lngHandle, strValueName, 0, lngResultType, m_strBuffer, m_lngLength)
        End If
    End If
    m_lngHandle = RegCloseKey(m_lngHandle)
    If m_lngResult = 0 Then
        If lngResultType = REG_DWORD Then
            RegistryReadValue = m_lngBuffer
        Else
            RegistryReadValue = Left$(m_strBuffer, m_lngLength - 1)
        End If
    End If
    Err.Clear
End Function

Public Sub RegistryWriteValueLong(lngStandardKey As Long, strKeyName As String, lngValue As Long, strValueName As String)
'
'   Writes a long (dword) registry value
'
    On Error Resume Next
    If strValueName = "" Then
        Exit Sub
    End If
    If RegOpenKeyEx(lngStandardKey, strKeyName, 0&, KEY_SET_VALUE, m_lngHandle) <> 0 Then
        If RegCreateKeyEx(lngStandardKey, strKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, m_lngHandle, m_lngResult) <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End If
    m_lngLength = 4
    m_lngResult = RegSetValueExLong(m_lngHandle, strValueName, 0&, REG_DWORD, lngValue, m_lngLength)
    m_lngResult = RegCloseKey(m_lngHandle)
    Err.Clear
End Sub

Public Sub RegistryWriteValueString(lngStandardKey As Long, strKeyName As String, strValue As String, Optional strValueName As String = "")
'
'   Writes a string registry value
'
    On Error Resume Next
    If RegOpenKeyEx(lngStandardKey, strKeyName, 0&, KEY_SET_VALUE, m_lngHandle) <> 0 Then
        If RegCreateKeyEx(lngStandardKey, strKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, m_lngHandle, m_lngResult) <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End If
    m_lngLength = Len(strValue)
    If strValueName = "" Then
        m_lngResult = RegSetValue(m_lngHandle, "", REG_SZ, strValue, m_lngLength)
    Else
        m_lngResult = RegSetValueExString(m_lngHandle, strValueName, 0&, REG_SZ, strValue, m_lngLength)
    End If
    m_lngResult = RegCloseKey(m_lngHandle)
    Err.Clear
End Sub

Public Function RemoveCharsFromString(strString As String, strRemoveChars As String) As String
'
'   Deletes every occurance of the specifed characters from a string
'
    Dim lngX As Long
    Dim lngY As Long
    Dim strResult As String
    
    On Error Resume Next
    strResult = strString
    If Len(strResult) > 0 And Len(strString) > 0 Then
        lngX = 1
        Do
            lngY = InStr(strRemoveChars, Mid$(strResult, lngX, 1))
            If lngY > 0 Then
                strResult = Left$(strResult, lngX - 1) & Mid$(strResult, lngX + 1)
            Else
                lngX = lngX + 1
            End If
        Loop While lngX <= Len(strResult)
    End If
    RemoveCharsFromString = strResult
    Err.Clear
End Function

Public Function RenameDirectory(strOldDirectoryName As String, strNewDirectoryName As String) As Boolean
'
'   Deletes a file
'
    On Error Resume Next
    RenameDirectory = RenameFile(strOldDirectoryName, strNewDirectoryName)
End Function

Public Function RenameFile(strOldFileName As String, strNewFileName As String) As Boolean
'
'   Deletes a file
'
    On Error Resume Next
    Name strOldFileName As strNewFileName
    If Err Then
        RenameFile = False
    Else
        RenameFile = True
    End If
    Err.Clear
End Function

Public Function RunFile(strFileName As String, lngHWnd As Long) As Boolean
'
'   Open a file with the application that supports it
'
    On Error Resume Next
    m_lngResult = ShellExecute(lngHWnd, "Open", strFileName, "", "", vbNormalFocus)
    If m_lngResult < 0 Or m_lngResult > 32 Then
        DoEvents
        DoEvents
        DoEvents
        RunFile = True
    End If
    Err.Clear
End Function

Public Function RunFunction(strFileName As String, strParams As String, lngHWnd As Long, Optional lngPidl As Long = 0) As Boolean
'
'   Exectues a function using the ShellExecuteEx function
'
    Dim lngResult As Long
    Dim recInfo As typSHELLEXECUTEINFOString
    
    On Error Resume Next
    With recInfo
        'This method was adopted to use a filename instead of a pidl reference
        .szFile = strFileName
        .szParameters = strParams
        .cbSize = Len(recInfo)
        .fMask = SEE_MASK_INVOKEIDLIST
        .hWnd = lngHWnd
        .nShow = SW_SHOWNORMAL
        .hInstApp = App.hInstance
    End With
    lngResult = ShellExecuteExString(recInfo)
    Err.Clear
    RunFunction = IIf(lngResult = 0, False, True)
End Function

Public Function RunProgram(strLaunchString As String, Optional lngWindowStyle As VbAppWinStyle = -1) As Double
'
'   Runs a program or bat file
'
    Dim dblResult As Double
    Dim lngWindow As Long
    
    On Error Resume Next
    If lngWindowStyle < 0 Then
        lngWindow = vbHide
    Else
        lngWindow = lngWindowStyle
    End If
    If lngWindow <> vbHide And lngWindow <> vbMaximizedFocus And lngWindow <> vbMinimizedFocus And lngWindow <> vbMinimizedNoFocus And lngWindow <> vbNormalFocus And lngWindow <> vbNormalNoFocus Then
        lngWindow = vbHide
    End If
    dblResult = Shell(strLaunchString, lngWindow)
    If Err.Number = 0 And dblResult <> 0 Then
        DoEvents
        DoEvents
        DoEvents
        RunProgram = True
    End If
    Err.Clear
End Function

Public Sub SetCursorPosition(lngX As Long, lngY As Long)
'
'   Sets the current mouse position
'
    Call SetCursorPos(lngX, lngY)
End Sub

Public Function SetDirectory(strPath As String) As Boolean
'
'   Sets the current directory
'
    On Error Resume Next
    ChDir (strPath)
    If Err Then
        SetDirectory = False
    Else
        SetDirectory = True
    End If
    Err.Clear
End Function

Public Function SetFileAssociation(strExtension As String, strAppName As String, strAppFileName As String, blnIcon As Boolean, Optional blnForce As Boolean) As Boolean
'
'   Create a file association to a program
'
    Dim strAssociation As String
    Dim strKey As String
    
    On Error Resume Next
    strAssociation = "." & LCase$(strExtension)
    If RegistryReadValue(HKEY_CLASSES_ROOT, strAssociation, "") <> "" And Not blnForce Then
        'Extension is already associated
        Exit Function
    End If
    Call RegistryWriteValueString(HKEY_CLASSES_ROOT, strAssociation, strAppName, "")
    Call RegistryWriteValueString(HKEY_CLASSES_ROOT, strAppName & "\Shell\Open\Command", """" & strAppFileName & """" & " " & """%1""", "")
    If blnIcon Then
        Call RegistryWriteValueString(HKEY_CLASSES_ROOT, strAppName & "\DefaultIcon", strAppFileName & ",1", "")
        Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, vbNullString, vbNullString)
    End If
    Err.Clear
    SetFileAssociation = True
End Function

Public Function SetFileContents(strData As String, strFileName As String) As Boolean
'
'   Create a file from the contents of a string
'
    Dim intFileNo As Integer
    
    On Error Resume Next
    intFileNo = FreeFile
    Open strFileName For Output As #intFileNo
    If Not Err Then
        Print #intFileNo, strData;
        Close #intFileNo
        SetFileContents = True
    Else
        Err.Clear
    End If
End Function

Public Sub SetINIString(strFileName As String, strSectionName As String, strStringName As String, strStringValue As String)
'
'   Writes a value to an INI file
'
    On Error Resume Next
    m_lngLength = WritePrivateProfileString(strSectionName, strStringName, strStringValue, strFileName)
    Err.Clear
End Sub

Public Sub SetLocaleID(lngLocaleID As Long)
'
'   Sets the locale for the application
'
    On Error Resume Next
    m_lngResult = SetThreadLocale(lngLocaleID)
    Err.Clear
End Sub

Public Sub SetLocaleData(lngLocaleType As Long, strLocaleValue As String)
'
'   Sets a locate value for the application
'
    On Error Resume Next
    m_lngResult = SetLocaleInfo(GetLocaleID, lngLocaleType, strLocaleValue)
    Err.Clear
End Sub

Public Function SetMaskEdBoxValue(strMask As String, strValue As String) As String
'
'   Sets a masked edit box field value
'
    Dim intIX0 As Integer
    Dim intIX1 As Integer
    Dim strText As String
    
    On Error Resume Next
    If strMask = "" Then
        SetMaskEdBoxValue = strValue
    Else
        strText = ""
        If Len(strValue) > 0 Then
            intIX0 = 1
            For intIX1 = 1 To Len(strMask)
                If Mid$(strMask, intIX1, 1) = "#" Then
                    If intIX0 <= Len(strValue) Then
                        strText = strText + Mid$(strValue, intIX0, 1)
                    Else
                        strText = strText + " "
                    End If
                    intIX0 = intIX0 + 1
                Else
                    strText = strText + Mid$(strMask, intIX1, 1)
                End If
            Next
        End If
        SetMaskEdBoxValue = strText
    End If
    Err.Clear
End Function

Public Function SetSlash(strPath As String) As String
'
'   Adds a backslash to the end of a string if one is not already there
'
    On Error Resume Next
    If Right$(strPath, 1) = "\" Then
        SetSlash = strPath
    Else
        SetSlash = strPath + "\"
    End If
    Err.Clear
End Function

Public Function SetSlashNone(strPath As String) As String
'
'   Remove all backslashes at the end of a string
'
    Dim strResult As String
    
    On Error Resume Next
    strResult = strPath
    Do
        If Right$(strResult, 1) = "\" Then
            strResult = Left$(strResult, Len(strResult) - 1)
        Else
            Exit Do
        End If
    Loop
    SetSlashNone = strResult
    Err.Clear
End Function

Public Sub SetTopmost(hWnd As Long, blnSetTopmost As Boolean)
'
'   Sets or resets a window as the topmost window
'
    Dim lngFlags As Long
    Dim lngWindow As Long
    
    On Error Resume Next
    lngWindow = IIf(blnSetTopmost, HWND_TOPMOST, HWND_NOTTOPMOST)
    lngFlags = SWP_NOSIZE + SWP_NOMOVE
    Call SetWindowPos(hWnd, lngWindow, 0, 0, 0, 0, lngFlags)
    Err.Clear
End Sub

Public Function ShiftLeft(ByVal lngValue As Long, ByVal intBitCount As Integer) As Long
'
'   Performs an unsigned left shift
'
    On Error Resume Next
    If intBitCount <= 0 Or intBitCount > 63 Then
        ShiftLeft = lngValue
    ElseIf intBitCount > 31 Then
        ShiftLeft = 0
    Else
        If (lngValue And m_lngShift(31)) = m_lngShift(31) Then
            ShiftLeft = (lngValue And &H7FFFFFFF) \ m_lngShift(intBitCount) Or m_lngShift(31 - intBitCount)
        Else
            ShiftLeft = lngValue \ m_lngShift(intBitCount)
        End If
    End If
End Function

Public Function ShiftRight(ByVal lngValue As Long, ByVal intBitCount As Integer) As Long
'
'   Performs an unsigned right shift
'
    On Error Resume Next
    If intBitCount <= 0 Or intBitCount > 63 Then
        ShiftRight = lngValue
    ElseIf intBitCount > 31 Then
        ShiftRight = 0
    Else
        If (lngValue And m_lngShift(31 - intBitCount)) = m_lngShift(31 - intBitCount) Then
            ShiftRight = (lngValue And (m_lngShift(31 - intBitCount) - 1)) * m_lngShift(intBitCount) Or m_lngShift(31)
        Else
            ShiftRight = (lngValue And (m_lngShift(31 - intBitCount) - 1)) * m_lngShift(intBitCount)
        End If
    End If
End Function

Public Sub Slumber(lngMilliseconds As Long)
'
'   Sleep for "n" milliseconds
'
    On Error Resume Next
    DoEvents
    Call Sleep(lngMilliseconds)
    DoEvents
    DoEvents
    DoEvents
    Err.Clear
End Sub

Public Function StripString(strValue As String, Optional blnStripCrLfHt As Boolean = False, Optional blnStripWeb As Boolean = False, Optional strStripOther As String = "", Optional strQuoteChar As String = "") As String
'
'   Remove all non-printable characters from a string and return with quotes, if desired
'
    Dim lngChar As Long
    Dim lngX As Long
    Dim strStrip As String
    Dim strString As String
    
    On Error Resume Next
    'Miscellaneous characters to be stripped
    'Other characters to be stripped could include:  %&/\^`~|"
    strStrip = strStripOther & vbNullChar
    If blnStripWeb Then
        strStrip = strStrip & "<>'"
    End If
    For lngX = 1 To Len(strValue)
        lngChar = Asc(Mid$(strValue, lngX, 1))
        If (lngChar < 32 And (blnStripCrLfHt Or (lngChar <> vbKeyReturn And lngChar <> vbKeyLF And lngChar <> vbKeyTab))) _
            Or lngChar > 126 Then
            lngChar = 0
        End If
        If InStr(strStrip, Chr$(lngChar)) = 0 Then
            strString = strString & Chr$(lngChar)
        End If
    Next
    Err.Clear
    StripString = strQuoteChar & strString & strQuoteChar
End Function

Public Function Succeeded(hr As Long) As Boolean   ' hr = HRESULT
'
'   Provides a generic test for success on any HResult status value.
'   Non-negative numbers indicate success.
'
    On Error Resume Next
    If (hr >= S_OK) Then
        Succeeded = True
    Else
        If IsIDE Then
            If (MsgBox("Error: &H" & Hex(hr) & ", " & GetAPIErrorString(hr) & vbCrLf & vbCrLf & _
                "View offending code?", vbExclamation Or vbYesNo) = vbYes) Then
                Stop
                ' hit Ctrl+L to view the call stack...
            End If
        Else
            g_strAPIErrorMessage = "Error: &H" & Hex(hr) & ", " & GetAPIErrorString(hr)
        End If
    End If
End Function

Public Function UnquoteString(strValue As String) As String
    Dim lngPos As Long
    Dim strString As String
    
    On Error Resume Next
    strString = Trim$(strValue)
    If Len(strString) > 1 And (Left$(strString, 1) = """" Or Left$(strString, 1) = "'") Then
        lngPos = InStr(Mid$(strString, 2), Left$(strString, 1))
        If lngPos > 0 Then
            strString = Mid$(strString, 2, lngPos - 1)
        End If
    End If
    Err.Clear
    UnquoteString = strString
End Function

Public Function WaitProgram(dblInstanceHandle As Double) As Boolean
'
'   Wait for a specific program (launched with the VB Shell function) to end
'
    Dim intInstanceHandle As Integer
    
    On Error Resume Next
    DoEvents
    DoEvents
    DoEvents
    intInstanceHandle = dblInstanceHandle
    m_lngHandle = FindWindow(vbNullString, vbNullString)
    Do Until m_lngHandle = 0
        If GetParent(m_lngHandle) = 0 Then
            If intInstanceHandle = GetWindowWord(m_lngHandle, GWW_HINSTANCE) Then
               Exit Do
            End If
        End If
        m_lngHandle = GetWindow(m_lngHandle, GW_HWNDNEXT)
    Loop
    If m_lngHandle = 0 Then
        WaitProgram = True
    End If
    Err.Clear
End Function

