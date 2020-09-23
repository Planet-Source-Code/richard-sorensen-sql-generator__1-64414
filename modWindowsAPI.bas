Attribute VB_Name = "modWindowsAPI"
Option Explicit
Option Base 0
'
'   Windows constants and declarations
'
'transparent color (the imagelist will use each icon's mask)
Public Const CLR_NONE = &HFFFFFFFF
'CodePage
Public Const CP_ACP = 0        ' ANSI code page
Public Const CP_OEMCP = 1   ' OEM code page
'Date Flags for GetDateFormat
Public Const DATE_SHORTDATE = &H1                  ' use short date picture
Public Const DATE_LONGDATE = &H2                     ' use long date picture
Public Const DATE_USE_ALT_CALENDAR = &H4   ' use alternate calendar (if any)
'
Public Const FILE_ATTRIBUTE_NORMAL = &H80
' dwFlags
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
'
Public Const GENERIC_READ = &H80000000
Public Const GW_HWNDNEXT = 2
Public Const GWW_HINSTANCE = (-6)
'Registry key definitions
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
'Window handle stuff
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2
' TVITEM.iImage/iSelectedImage, LVITEM.iImage
Public Const I_IMAGECALLBACK = (-1)
'FindFirstFile error rtn value
Public Const INVALID_HANDLE_VALUE = -1
'
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F
'dwLanguageId
Public Const LANG_USER_DEFAULT = &H400&
'
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)
'Local IDs
Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400
'
Public Const MAXDWORD = (2 ^ 32) - 1   ' 0xFFFFFFFF
Public Const MAX_PATH = 260
'Date Flag for GetDateFormat, Time Flag for GetTimeFormat
Public Const LOCALE_NOUSEROVERRIDE = &H80000000    ' do not use user overrides
' TV/LV_ITEM.pszText
Public Const LPSTR_TEXTCALLBACK = (-1)
'
Public Const OPEN_EXISTING = 3
'Registry values
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_DWORD As Long = 4
Public Const REG_SZ As Long = 1
'Hresult
Public Const S_OK = 0
Public Const S_FALSE = 1&   ' special HRESULT value
' SHELLEXECUTEINFO fMask
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0
Public Const SM_CXSCREEN As Long = 0
Public Const SM_CYSCREEN As Long = 1
' SHELLEXECUTEINFO nShow
Public Const SW_SHOWNORMAL As Long = 1
'Window movement
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
'Windows notification messages
Public Const NM_FIRST = -0&   ' (0U-  0U)       ' // generic to all controls
Public Const NM_CLICK = (NM_FIRST - 2)
Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_RETURN = (NM_FIRST - 4)
Public Const NM_RCLICK = (NM_FIRST - 5)
' Time Flags for GetTimeFormat
Public Const TIME_NOMINUTESORSECONDS = &H1  ' do not use minutes or seconds
Public Const TIME_NOSECONDS = &H2                        ' do not use seconds
Public Const TIME_NOTIMEMARKER = &H4                 ' do not use time marker, i.e AM/PM
Public Const TIME_FORCE24HOURFORMAT = &H8     ' always use 24 hour format
'Time Zone constants
Public Const TIME_ZONE_ID_UNKNOWN As Long = 0
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'
Public Const VK_LSHIFT As Long = &HA0
Public Const VK_RSHIFT As Long = &HA1
'Window messages
Public Const WM_CANCELMODE = &H1F
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_DESTROY = &H2
Public Const WM_DRAWITEM = &H2B
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_NOTIFY = &H4E
Public Const WM_MEASUREITEM = &H2C
Public Const WM_PASTE = &H302
Public Const WM_USER = &H400
Public Const OCM__BASE = (WM_USER + &H1C00)
Public Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)

'Enumerations
Public Enum BFFM_FromDlg
    BFFM_INITIALIZED = 1
    BFFM_SELCHANGED = 2
End Enum

'messages to browser
Public Enum BFFM_ToDlg
    BFFM_SETSTATUSTEXTA = (&H400 + 100)
    BFFM_ENABLEOK = (&H400 + 101)
    BFFM_SETSELECTIONA = (&H400 + 102)
    BFFM_SETSELECTIONW = (&H400 + 103)
    BFFM_SETSTATUSTEXTW = (&H400 + 104)
End Enum

'Browsing for directory.
Public Enum BF_Flags
    BIF_RETURNONLYFSDIRS = &H1      ' For finding a folder to start document searching
    BIF_DONTGOBELOWDOMAIN = &H2     ' For starting the Find Computer
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10               ' Add an editbox to the dialog.  Always on with BIF_USENEWUI
    BIF_VALIDATE = &H20              ' insist on valid result (or CANCEL)
    BIF_USENEWUI = &H40              ' Use the new dialog layout with the ability to resize.
    BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
    BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
    BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything
End Enum

Public Enum CBoolean
    CFalse = 0
    CTrue = 1
End Enum

Public Enum CSIDL_VALUES
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Public Enum SHGFI_flags
    SHGFI_LARGEICON = &H0
    SHGFI_SMALLICON = &H1
    SHGFI_OPENICON = &H2
    SHGFI_SHELLICONSIZE = &H4
    SHGFI_PIDL = &H8
    SHGFI_USEFILEATTRIBUTES = &H10
    SHGFI_ICON = &H100
    SHGFI_DISPLAYNAME = &H200
    SHGFI_TYPENAME = &H400
    SHGFI_ATTRIBUTES = &H800
    SHGFI_ICONLOCATION = &H1000
    SHGFI_EXETYPE = &H2000
    SHGFI_SYSICONINDEX = &H4000
    SHGFI_LINKOVERLAY = &H8000
    SHGFI_SELECTED = &H10000
    SHGFI_ATTR_SPECIFIED = &H20000
End Enum

Public Enum TPM_wFlags
    TPM_LEFTBUTTON = &H0
    TPM_RIGHTBUTTON = &H2
    TPM_LEFTALIGN = &H0
    TPM_CENTERALIGN = &H4
    TPM_RIGHTALIGN = &H8
    TPM_TOPALIGN = &H0
    TPM_VCENTERALIGN = &H10
    TPM_BOTTOMALIGN = &H20
    TPM_HORIZONTAL = &H0         ' Horz alignment matters more
    TPM_VERTICAL = &H40            ' Vert alignment matters more
    TPM_NONOTIFY = &H80           ' Don't send any notification msgs
    TPM_RETURNCMD = &H100
End Enum

Public Enum enumHookType
   WH_KEYBOARD = 2
   WH_MOUSE = 7
End Enum


'Type declarations
Public Type typFILETIME   ' ft
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type typFileVersionInfo
    lngSignature As Long
    lngStructVersion As Long
    lngFileVersionMS As Long
    lngFileVersionLS As Long
    lngProductVersionMS As Long
    lngProductVersionLS As Long
    lngFileFlagsMask As Long
    lngFileFlags As Long
    lngFileOS As Long
    lngFileType As Long
    lngFileSubType As Long
    lngFileDateMS As Long
    lngFileDateLS As Long
End Type

Public Type typGuid
    lngData1 As Long
    intData2 As Integer
    intData3 As Integer
    bytData4(8) As Byte
End Type

Public Type typNMHDR
    hwndFrom As Long   ' Window handle of control sending message
    idFrom As Long        ' Identifier of control sending message
    code  As Long          ' Specifies the notification code
End Type

Public Type typPOINT
    X As Long
    Y As Long
End Type

Public Type typRECT   ' rct
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type typSecurityAttributes
    lngLength As Long
    lngSecurityDescriptor As Long
    lngInheritHandle As Long
End Type

'Browse for folder structure
Public Type typSHBROWSEINFO
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Public Type typSHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long   ' String
    lpFile As Long   ' String
    lpParameters As Long   ' String
    lpDirectory As Long   ' String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As Long   ' String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type typSHELLEXECUTEINFOString
    cbSize As Long
    fMask As Long
    hWnd As Long
    szVerb As String
    szFile As String
    szParameters As String
    szDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As Long   ' String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type typSHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Type typSYSTEMTIME
    intYear As Integer
    intMonth As Integer
    intDayOfWeek As Integer
    intDay As Integer
    intHour As Integer
    intMinute As Integer
    intSecond As Integer
    intMilliseconds As Integer
End Type

Public Type typTimeZone
    lngBias As Long
    intStandardName(31) As Integer
    objStandardDate As typSYSTEMTIME
    lngStandardBias As Long
    intDaylightName(31) As Integer
    objDaylightDate As typSYSTEMTIME
    lngDaylightBias As Long
End Type

Public Type typVersionInfo
    lngBufferSize As Long
    lngMajorVersion As Long
    lngMinorVersion As Long
    lngBuildNumber As Long
    lngPlatformID As Long
    strPlatform As String * 128
End Type

Public Type typWIN32_FIND_DATA   ' wfd
    dwFileAttributes As Long
    ftCreationTime As typFILETIME
    ftLastAccessTime As typFILETIME
    ftLastWriteTime As typFILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

'Procedure declarations
'Kernel
Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, _
    ByVal Source As Long, _
    ByVal Length As Long)
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As typSecurityAttributes, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Public Declare Function DeleteFileX Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" _
    (lpFileTime As typFILETIME, _
    lpLocalFileTime As typFILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" _
    (lpFileTime As typFILETIME, _
    lpSystemTime As typSYSTEMTIME) As Long
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" _
    (pDest As Any, _
    ByVal dwLength As Long, _
    ByVal bFill As Byte)
Public Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Boolean
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
    lpFindFileData As typWIN32_FIND_DATA) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" _
    (ByVal Locale As Long, _
    ByVal dwFlags As Long, _
    lpDate As typSYSTEMTIME, _
    ByVal lpFormat As String, _
    ByVal lpDateStr As String, _
    ByVal cchDate As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileName As String) As Long
Public Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" _
    (ByVal Locale As Long, _
    ByVal dwFlags As Long, _
    lpTime As typSYSTEMTIME, _
    ByVal lpFormat As String, _
    ByVal lpTimeStr As String, _
    ByVal cchTime As Long) As Long
Public Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" _
    (ByVal strWhichDrive As String) As Long
Public Declare Function GetFileSizeX Lib "kernel32" Alias "GetFileSize" _
    (ByVal hFile As Long, _
    lpFileSizeHigh As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal lSize As Long, _
    ByVal lpFileName As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetThreadLocale Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As typVersionInfo) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As typTimeZone) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long
Public Declare Function LocalAlloc Lib "kernel32" _
    (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" _
    (ByVal hMem As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
Public Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Public Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Public Declare Function lstrcmpiA Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Public Declare Function lstrcmpiW Lib "kernel32" (lpString1 As Any, _
    lpString2 As Any) As Long
Public Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Public Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Public Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
    (ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String) As Long
Public Declare Function SetThreadLocale Lib "kernel32" _
    (ByVal Locale As Long) As Long
Public Declare Sub Sleep Lib "kernel32" _
    (ByVal lngMilliseconds As Long)
Public Declare Function WideCharToMultiByte Lib "kernel32" _
    (ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    lpWideCharStr As Any, _
    ByVal cchWideChar As Long, _
    lpMultiByteStr As Any, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As String, _
    ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As Any, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lplFileName As String) As Long

'Advapi DLL
Public Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, _
    nSize As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal lpValue As String, _
    lpcbValue As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal dwType As Long, _
    ByVal lpData As String, _
    ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpValue As Long, _
    ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal _
    lpValue As String, _
    ByVal cbData As Long) As Long

'Comctl32.dll
Public Declare Function ImageList_SetBkColor Lib "comctl32.dll" _
    (ByVal himl As Long, _
    ByVal clrBk As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" _
    (ByVal himl As Long) As Long

'MultiMedia
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
      (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'OLE32.DLL
Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal pv As Long)
Public Declare Function CoCreateGuid Lib "ole32.dll" _
    (gGuid As typGuid) As Long

'SetupToolkit DLL
Public Declare Function AllocUnit Lib "STKIT432.DLL" () As Long
Public Declare Function DiskSpaceFree Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long
Public Declare Function DLLSelfRegister Lib "STKIT432.DLL" _
    (ByVal lpDllName As String) As Integer
Public Declare Function fNTWithShell Lib "STKIT432.DLL" () As Boolean
Public Declare Function FSyncShell Lib "STKIT432.DLL" Alias "SyncShell" _
    (ByVal strCmdLine As String, _
    ByVal intCmdShow As Long) As Long
Public Declare Sub lmemcpy Lib "STKIT432.DLL" _
    (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
Public Declare Function OSfCreateShellGroup Lib "STKIT432.DLL" Alias "fCreateShellFolder" _
    (ByVal lpstrDirName As String) As Long
Public Declare Function OSfCreateShellLink Lib "STKIT432.DLL" Alias "fCreateShellLink" _
    (ByVal lpstrFolderName As String, _
    ByVal lpstrLinkName As String, _
    ByVal lpstrLinkPath As String, _
    ByVal lpstrLinkArguments As String) As Long
Public Declare Function OSGetLongPathName Lib "STKIT432.DLL" Alias "GetLongPathName" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Public Declare Function OSfRemoveShellLink Lib "STKIT432.DLL" Alias "fRemoveShellLink" _
    (ByVal lpstrFolderName As String, _
    ByVal lpstrLinkName As String) As Long
Public Declare Function SetTime Lib "STKIT432.DLL" _
    (ByVal strFileGetTime As String, _
    ByVal strFileSetTime As String) As Integer

'Shell DLL
Public Declare Function FileIconInit Lib "shell32.dll" Alias "#660" _
    (ByVal cmd As Boolean) As Boolean
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As typSHBROWSEINFO) As Long
Public Declare Function SHChangeNotify Lib "shell32.dll" _
    (ByVal wEventID As Long, _
    ByVal uFlags As Long, _
    ByVal dwItem1 As String, _
    ByVal dwItems As String) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
    (ByVal hWnd As Long, _
    ByVal szApp As String, _
    ByVal szOtherStuff As String, _
    ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" _
    (lpExecInfo As typSHELLEXECUTEINFO) As Long
Public Declare Function ShellExecuteExString Lib "shell32.dll" Alias "ShellExecuteEx" _
    (lpExecInfo As typSHELLEXECUTEINFOString) As Long
Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
    (ByVal pszPath As Any, _
    ByVal dwFileAttributes As Long, _
    psfi As typSHFILEINFO, _
    ByVal cbFileInfo As Long, _
    ByVal uFlags As Long) As Long
Public Declare Function SHGetFolderLocation Lib "shell32.dll" Alias "SHGetFolderLocationA" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ByVal hToken As Long, _
    ByVal dwReserved As Long, _
    ppidl As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long) As Long
Public Declare Function SHParseDisplayName Lib "shell32.dll" _
    (ByVal pszName As Any, _
    ByVal pbc As Long, _
    ppidl As Long, _
    ByVal sfgaoIn As Long, _
    psfgaoOut As Long) As Long

'User DLL
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" _
    (ByVal hWnd As Long, _
    lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" _
    (ByVal hMenu As Long) As Long
Public Declare Function DestroyWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hwndParent As Long, _
    ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
Public Declare Function GetClientRect Lib "user32" _
    (ByVal hWnd As Long, _
    lpRect As typRECT) As Long
Public Declare Function GetCursorPos Lib "user32" _
    (lpPoint As typPOINT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer
Public Declare Function GetParent Lib "user32" _
    (ByVal hWnd As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" _
    (ByVal hWnd As Long, _
    ByVal lpString As Any) As Long
Public Declare Function GetWindow Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal wCmd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, _
    lpRect As typRECT) As Long
Public Declare Function GetWindowWord Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Public Declare Function InvalidateRect Lib "user32" _
    (ByVal hWnd As Long, _
    lpRect As Any, _
    ByVal bErase As CBoolean) As CBoolean
Public Declare Function LoadString Lib "user32" Alias "LoadStringA" _
    (ByVal hInstance As Long, _
    ByVal uID As Long, _
    ByVal lpBuffer As String, _
    ByVal nBufferMax As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String) As Long
Public Declare Function ScreenToClient Lib "user32" _
    (ByVal hWnd As Long, _
    lpPoint As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SetCursorPos Lib "user32" _
    (ByVal X As Long, _
    ByVal Y As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" _
    (ByVal hMenu As Long, _
    ByVal wFlags As TPM_wFlags, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nReserved As Long, _
    ByVal hWnd As Long, _
    lprc As Any) As Long
Public Declare Function UpdateWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long

'Version DLL
Public Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" _
    (ByVal strFileName As String, _
    ByVal lVerHandle As Long, _
    ByVal lcbSize As Long, _
    lpvData As Byte) As Long
Public Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" _
    (ByVal strFileName As String, _
    lVerHandle As Long) As Long
Public Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" _
    (ByVal Flags&, _
    ByVal SrcName$, _
    ByVal DestName$, _
    ByVal SrcDir$, _
    ByVal DestDir$, _
    ByVal CurrDir As Any, _
    ByVal TmpName$, _
    lpTmpFileLen&) As Long
Public Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" _
    (lpvVerData As Byte, _
    ByVal lpszSubBlock As String, _
    lplpBuf As Long, _
    lpcb As Long) As Long

