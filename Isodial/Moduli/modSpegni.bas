Attribute VB_Name = "modSpegni"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modSpegni.bas
'
' <b>Descrizione</b>: Insieme di funzioni e oggetti per forzare lo spegnimento del pc
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.08

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TokenPrivileges = 3
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const ANYSIZE_ARRAY = 1

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type Luid
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Function InitiateShutdownMachine(ByVal Machine As String, _
  Optional Force As Boolean = False, _
  Optional Restart As Boolean = False, _
  Optional Message As String = "") As Boolean
  
    Dim hProc As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuffLen As Long
    Dim pSize As Long


    If InStr(Machine, "\\") = 1 Then
        Machine = Right(Machine, Len(Machine) - 2)
    End If

    If (LCase(GetNomeComputer) = LCase(Machine)) Then

        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
            MsgBox "Errore: " & GetLastError()
            Exit Function
        End If

        If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
            MsgBox "Errore: " & GetLastError()
            Exit Function
        End If
        NewTokenStuff = OldTokenStuff
        NewTokenStuff.PrivilegeCount = 1
        NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        NewTokenStuffLen = Len(NewTokenStuff)
        pSize = Len(NewTokenStuff)

        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
            MsgBox "Errore: " & GetLastError()
            Exit Function
        End If

        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
        NewTokenStuff.Privileges(0).Attributes = 0

        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
            Exit Function
        End If
    Else

        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
    End If
    InitiateShutdownMachine = True
End Function

Private Function GetNomeComputer() As String
    Dim sLen As Long
    sLen = 255
    GetNomeComputer = Space(sLen)
    
    If GetComputerName(GetNomeComputer, sLen) Then
        GetNomeComputer = Left$(GetNomeComputer, sLen)
        GetNomeComputer = Trim$(GetNomeComputer)
    End If
End Function

Public Sub Spegni()
    Dim ret As Boolean
    ret = InitiateShutdownMachine(GetNomeComputer, True, False)
End Sub

