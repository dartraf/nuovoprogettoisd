Attribute VB_Name = "modWheel"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modWheel.bas
'
' <b>Descrizione</b>: Insieme di funzioni e oggetti per controllare la rotellina del mouse
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.07
Option Explicit

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, _
ByVal hWnd As Long, _
ByVal Msg As Long, _
ByVal Wparam As Long, _
ByVal Lparam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Const MK_CONTROL = &H8
Private Const MK_LBUTTON = &H1
Private Const MK_RBUTTON = &H2
Private Const MK_MBUTTON = &H10
Private Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Private LocalHwnd As Long
Private LocalPrevWndProc As Long
Private MyForm As Form
Private flx As MSFlexGrid

'Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long, ByVal a As Integer) As Long

'    Dim MouseKeys As Long
'    Dim Rotation As Long
'    Dim Xpos As Long
'    Dim Ypos As Long

'    If Lmsg = WM_MOUSEWHEEL Then
'        MouseKeys = Wparam And 65535
'        Rotation = Wparam / 65536
'        Xpos = Lparam And 65535
'        Ypos = Lparam / 65536
'        MyForm.MouseWheel flx, MouseKeys, Rotation, Xpos, Ypos
'    End If
'    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)
'End Function

'Public Sub WheelHook(PassedForm As Form, vflx As MSFlexGrid)

'    On Error Resume Next

'    Set MyForm = PassedForm
'    Set flx = vflx
'    LocalHwnd = PassedForm.hWnd
'    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
'End Sub


'Public Sub WheelUnHook()
'    Dim WorkFlag As Long

'    On Error Resume Next
'    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
'    Set MyForm = Nothing
'End Sub


