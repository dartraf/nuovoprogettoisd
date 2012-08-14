Attribute VB_Name = "modGestioneFileEsterni"
Option Explicit

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long

Private Type VS_FIXEDFILEINFO
   Signature As Long
   StrucVersionl As Integer     '  e.g. = &h0000 = 0
   StrucVersionh As Integer     '  e.g. = &h0042 = .42
   FileVersionMSl As Integer    '  e.g. = &h0003 = 3
   FileVersionMSh As Integer    '  e.g. = &h0075 = .75
   FileVersionLSl As Integer    '  e.g. = &h0000 = 0
   FileVersionLSh As Integer    '  e.g. = &h0031 = .31
   ProductVersionMSl As Integer '  e.g. = &h0003 = 3
   ProductVersionMSh As Integer '  e.g. = &h0010 = .1
   ProductVersionLSl As Integer '  e.g. = &h0000 = 0
   ProductVersionLSh As Integer '  e.g. = &h0031 = .31
   FileFlagsMask As Long        '  = &h3F for version "0.42"
   FileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   FileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   FileType As Long             '  e.g. VFT_DRIVER
   FileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   FileDateMS As Long           '  e.g. 0
   FileDateLS As Long           '  e.g. 0
End Type

''
' Restituisce la versione di un file dll, exe, ocx nel formato 00.00.0000
Function GetVersioneNumber(sFileName As String) As String
   On Error GoTo gestione
   Dim lFileHwnd As Long, lRet As Long, lBufferLen As Long, lplpBuffer As Long, lpuLen As Long
   Dim abytBuffer() As Byte
   Dim tVerInfo As VS_FIXEDFILEINFO
   Dim sBlock As String, sStrucVer As String

    'Get the size File version info structure
    lBufferLen = GetFileVersionInfoSize(sFileName, lFileHwnd)
    If lBufferLen = 0 Then
        GetVersioneNumber = ""
        Exit Function
    End If
    
    'Create byte array buffer, then copy memory into structure
    ReDim abytBuffer(lBufferLen)
    Call GetFileVersionInfo(sFileName, 0&, lBufferLen, abytBuffer(0))
    Call VerQueryValue(abytBuffer(0), "\", lplpBuffer, lpuLen)
    Call CopyMem(tVerInfo, ByVal lplpBuffer, Len(tVerInfo))
    
    'Determine structure version number (For info only)
    sStrucVer = Format$(tVerInfo.StrucVersionh) & "." & Format$(tVerInfo.StrucVersionl)
    
    'Concatenate file version number details into a result string
    GetVersioneNumber = Format$(tVerInfo.FileVersionMSh) & "." & Format$(tVerInfo.FileVersionMSl, "00") & "."
    If tVerInfo.FileVersionLSh > 0 Then
        GetVersioneNumber = GetVersioneNumber & Format$(tVerInfo.FileVersionLSh, "0000") & "." & Format$(tVerInfo.FileVersionLSl, "00")
    Else
        GetVersioneNumber = GetVersioneNumber & Format$(tVerInfo.FileVersionLSl, "0000")
    End If
    
    Exit Function
gestione:
    GetVersioneNumber = ""
End Function

Public Function IsCorrectVersion(inVersionRichiesta As String, inFileName As String, outVersionAttuale As String) As Boolean
    outVersionAttuale = GetVersioneNumber(inFileName)
    If inVersionRichiesta <> outVersionAttuale Then
        IsCorrectVersion = False
    Else
        IsCorrectVersion = True
    End If
End Function
