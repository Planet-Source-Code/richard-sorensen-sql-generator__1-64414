Attribute VB_Name = "modDialog"
Option Explicit
Option Base 0

Public Const HH_CLOSE_ALL As Integer = &H12
Public Const HH_HELP_CONTEXT As Integer = &HF
Public Const OFN_FILEMUSTEXIST = &H1000

Public g_intFormLoadCount As Integer        'Count of forms loaded - used for controlling auto-signoff

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private m_blnHelp As Boolean                'Flag indicating if help was called on the current form
Private m_blnHtmlHelpLoaded As Boolean      'Flag indicating that ocx has been loaded
Private m_lngResult As Long                 'Working result

Public Function DialogColor(ctlDialog As CommonDialog, strTitle As String, lngColorDefault As Long) As Long
    On Error Resume Next
    If App.TaskVisible Then
        With ctlDialog
            .CancelError = True
            .ShowColor
            If Err.Number = 0 Then
                DialogColor = .Color
            Else
                DialogColor = lngColorDefault
            End If
        End With
        Err.Clear
    End If
End Function

Public Function DialogFileOpen(ctlDialog As CommonDialog, strTitle As String, strFileName As String, strPath As String, strFilter As String, Optional lngFlags As Long = 0) As String
'
'   Performs a common dialog file open function
'
    On Error Resume Next
    If App.TaskVisible Then
        With ctlDialog
            .CancelError = True
            .DialogTitle = strTitle
            .DefaultExt = ""
            .Filename = strFileName
            .InitDir = strPath
            .Filter = strFilter
            .Flags = cdlOFNExplorer Or lngFlags
            .ShowOpen
            If Err.Number = 0 Then
                DialogFileOpen = .Filename
            Else
                DialogFileOpen = ""
            End If
        End With
        Err.Clear
    End If
End Function

Public Function DialogFileSave(ctlDialog As CommonDialog, strTitle As String, strFileName As String, strPath As String, strFilter As String, Optional lngFlags As Long = 0) As String
'
'   Performs a common dialog file save function
'
    On Error Resume Next
    If App.TaskVisible Then
        With ctlDialog
            .CancelError = True
            .DialogTitle = strTitle
            .DefaultExt = ""
            .Filename = strFileName
            .InitDir = strPath
            .Filter = strFilter
            .Flags = cdlOFNExplorer Or lngFlags
            .ShowSave
            If Err.Number = 0 Then
                DialogFileSave = .Filename
            Else
                DialogFileSave = ""
            End If
        End With
        Err.Clear
    End If
End Function

Public Function DialogFont(ctlDialog As CommonDialog) As Boolean
'
'   Displays the common dialog font selection window
'
    On Error Resume Next
    If App.TaskVisible Then
        With ctlDialog
            .CancelError = True
            .Flags = cdlCFBoth Or cdlCFEffects
            .ShowFont
            If Err.Number = 0 Then
                DialogFont = True
            Else
                DialogFont = False
            End If
        End With
        Err.Clear
    End If
End Function

Public Function DialogPrinterSetup(ctlDialog As CommonDialog, lngFlags As Long, intOrientation As Integer, intCopies As Integer, intFromPage As Integer, intToPage As Integer) As Boolean
'
'   Performs the common dialog printer setup function
'
    On Error Resume Next
    If App.TaskVisible Then
        With ctlDialog
            .CancelError = True
            .Flags = lngFlags
            .ShowPrinter
        End With
        If Err.Number = 0 Then
            With ctlDialog
                lngFlags = .Flags
                intOrientation = .Orientation
                intCopies = .Copies
                intFromPage = .FromPage
                intToPage = .ToPage
            End With
            DialogPrinterSetup = True
        Else
            DialogPrinterSetup = False
        End If
        Err.Clear
    End If
End Function

Public Sub DialogWebHelp(lngHWnd As Long, strHelpFile As String, lngHelpContextID As Long)
'
'   Performs a common dialog help calling function - for "chm" help files
'
    Dim strHelpFileType As String
        
    On Error Resume Next
    If App.TaskVisible Then
        If Not m_blnHtmlHelpLoaded Then
            m_lngResult = LoadLibrary("hhctrl.ocx")
            m_blnHtmlHelpLoaded = True
        End If
        m_lngResult = HtmlHelp(lngHWnd, strHelpFile, HH_HELP_CONTEXT, lngHelpContextID)
        Err.Clear
    End If
End Sub

Public Sub DialogWebHelpClose(lngHWnd As Long)
'
'   Closes help window if it was open
'
    m_lngResult = HtmlHelp(lngHWnd, "", HH_CLOSE_ALL, 0)
End Sub

Public Sub FormHelp(lngHWnd As Long, lngContextID As Long)
'
'   Call for help on a form
'
    On Error Resume Next
    m_blnHelp = True
    Call DialogWebHelp(lngHWnd, App.HelpFile, lngContextID)
    Err.Clear
End Sub

Public Sub FormLoad()
'
'   Keep track of loaded application forms
'
    g_intFormLoadCount = g_intFormLoadCount + 1
End Sub

Public Sub FormUnload(lngHWnd As Long)
'
'   Make sure that help is closed when leaving form
'
    On Error Resume Next
    g_intFormLoadCount = g_intFormLoadCount - 1
    If m_blnHelp Then
        Call DialogWebHelpClose(lngHWnd)
        m_blnHelp = False
    End If
    Err.Clear
End Sub

