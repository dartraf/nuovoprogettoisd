Attribute VB_Name = "modGetSysOp"
Option Explicit

'Determina se il sysop è a 32 o 64 bit
Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
    
Private Declare Function GetCurrentProcess Lib "kernel32" _
    () As Long

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long

'Restituisce la versione di sysop
Public Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
  
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Public Const VER_PLATFORM_WIN32s = 0        ' Win32s on Windows 3.1
Public Const VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 95, Windows 98, or Windows Me
Public Const VER_PLATFORM_WIN32_NT = 2      ' Windows NT, Windows 2000, Windows XP, or Windows Server 2003 family.

'                           Major     Minor
' OS              Platform  Version   Version  Build
' Windows 95      1         4          0
' Windows 98      1         4         10       1998
' Windows 98SE    1         4         10       2222
' Windows Me      1         4         90       3000
' NT 3.51         2         3         51
' NT              2         4          0       1381
' 2000            2         5          0
' XP              2         5          1       2600
' Server 2003     2         5          2

Public Function getVersion() As String
Dim OSInfo As OSVERSIONINFO
Dim retvalue As Integer
  
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(OSInfo)
  
With OSInfo
     Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32s           ' Win32s on Windows 3.1
              getVersion = "Windows 3.1"
               
         Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
              Select Case .dwMinorVersion   ' or Windows Me
                  Case 0
                      getVersion = "Windows 95"
                  Case 10
                      If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                          getVersion = "Windows 98SE"
                      Else
                          getVersion = "Windows 98"
                      End If
                  Case 90
                      getVersion = "Windows Me"
              End Select
     
         Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
              Select Case .dwMajorVersion   ' or Windows Server 2003 family.
                  Case 3
                      getVersion = "Windows NT 3.51"
                  Case 4
                      getVersion = "Windows NT 4.0"
                  Case 5
                      Select Case .dwMinorVersion
                          Case 0
                              getVersion = "Windows 2000"
                          Case 1
                              getVersion = "Windows XP"
                          Case 2
                              getVersion = "Windows Server 2003"
                      End Select
              End Select
                     
         Case Else
              getVersion = "Failed"
               
     End Select
             
End With
End Function

'Private Sub Command1_Click()
'MsgBox getVersion
'End Sub
    
Public Function Is64bit() As Boolean
    Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function

