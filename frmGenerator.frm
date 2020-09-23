VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGenerator 
   Caption         =   "SQL Code Generator"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "frmGenerator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ctlDialog 
      Left            =   120
      Top             =   1830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6420
      TabIndex        =   2
      Top             =   630
      Width           =   3075
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run SQL Generation Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   630
      Width           =   3075
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "Locate XML Schema File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   3075
   End
   Begin VB.Label lbl 
      Caption         =   $"frmGenerator.frx":0442
      Height          =   465
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   9405
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Width           =   9405
   End
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const m_strRegKey As String = "Software\SqlGenerator"
Private Const m_strRegValue As String = "FileName"

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLocate_Click()
    Dim strFileName As String
    Dim strFilePath As String
    
    If g_strXMLFile = "" Then
        strFileName = ""
        strFilePath = ""
    Else
        strFileName = GetFileName(g_strXMLFile)
        strFilePath = GetFilePath(g_strXMLFile)
    End If
    g_strXMLFile = DialogFileOpen(ctlDialog, "Find XML File", strFileName, strFilePath, "XML Files|*.xml")
    lblFileName.Caption = g_strXMLFile
End Sub

Private Sub cmdRun_Click()
    If lblFileName.Caption = "" Then
        MsgBox "The XML schema file must be located prior to performing this function", vbExclamation
        Exit Sub
    End If
    If Not FileExists(lblFileName.Caption) Then
        MsgBox "The XML schema file is not present", vbExclamation
        Exit Sub
    End If
    Call RegistryWriteValueString(HKEY_LOCAL_MACHINE, m_strRegKey, g_strXMLFile, m_strRegValue)
    Call Process
End Sub

Private Sub Form_Load()
    g_strXMLFile = RegistryReadValue(HKEY_LOCAL_MACHINE, m_strRegKey, m_strRegValue)
    If Not FileExists(g_strXMLFile) Then
        g_strXMLFile = ""
    End If
    lblFileName.Caption = g_strXMLFile
End Sub
