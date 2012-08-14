Attribute VB_Name = "dicApi"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - dicApi.bas
'
' <b>Descrizione</b>: Elenco delle API
'
' @remarks
'
' @author
'
' @date 28/01/2011 17.57
Option Explicit

'Per il collegamento al sito
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' usato per simulare la pressione di un tasto
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_EXTENDEDKEY = &H1        'pressione del tasto (keyDown)
Public Const KEYEVENTF_KEYUP = &H2             'rilascio del tasto premuto (keyUp)
Public Const VK_TAB = &H9

' per determinare la posizione del form Calndario e orario e altri
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' per poter spostare il form Calendario
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Wparam As Long, Lparam As Any) As Long

' per eliminare la x nel form login
Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long

' per caricare tutti i driver
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
    (ByVal nDrive As String) As Long

' per caricare le info sul driver
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, _
    ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long)
    
' per determinare lo spazio libero su disco
Public Declare Function GetDiskFreeSpace_FAT32 _
    Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
    (ByVal lpRootPathName As String, _
    FreeBytesToCaller As Currency, BytesTotal _
    As Currency, FreeBytesTotal As Currency) _
    As Long

' collegamento con lo scanner
Public Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal BmpFilename$) As Integer
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Public Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public Declare Function BmpToJpeg Lib "PicFormat32.dll" (ByVal BmpFilename As String, ByVal JpegFileName As String, ByVal Quality As Integer) As Integer
Public Const COMPRESSIONE As Integer = 20

' varie costanti
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const MF_BYPOSITION = &H400&
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const DRIVE_REMOVABLE = 2
Public Const MAX_PATH = 260

' per attendere la fine di un processo
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName _
    As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, _
    ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal _
    dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory _
    As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&
