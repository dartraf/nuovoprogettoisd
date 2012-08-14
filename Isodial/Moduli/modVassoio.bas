Attribute VB_Name = "modVassoio"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modVassoio.bas
'
' <b>Descrizione</b>: Insieme di funzioni e oggetti per ridurre il software nella icon tray
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.08
Option Explicit

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const wmjjser = &H400
Public Const cbNotify& = wmjjser + 42
Public Const uID& = 61860
Public myNID As NOTIFYICONDATA


Declare Function Shell_NotifyIcon Lib "shell32" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0

Public Const NIM_DELETE = &H2

Public Const NIM_MODIFY = &H1

Public Const NIF_MESSAGE = &H1

Public Const NIF_ICON = &H2

Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200

Public Const WM_MBUTTONDOWN = &H207

Public Const WM_LBUTTONUP = &H202

Public Const WM_MBUTTONDBLCLK = &H209

Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_RBUTTONDOWN = &H204

Public Const WM_RBUTTONUP = &H205

Public Const WM_MBUTTONUP = &H208

Public Const WM_LBUTTONDOWN = &H201

Public Const WM_LBUTTONDBLCLK = &H203


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Global lpPrevWndProc As Long
Global gHW As Long

Public Sub HOOK()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub unHook()
    Dim tmp As Long
    tmp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal Wparam As Long, _
                            ByVal Lparam As Long) As Long
    If Wparam = uID Then
     Select Case Lparam
        Case WM_MOUSEMOVE
         ' spostamento del mouse
        Case WM_LBUTTONDOWN
         ' pressione pulsante sinistro
        Case WM_LBUTTONUP
         ' rilascio pulsante sinistro
        Case WM_LBUTTONDBLCLK
         ' doppio click
         ' visualizza il form
         frmMain.Visible = True
         Shell_NotifyIcon NIM_DELETE, myNID
         unHook
         Case WM_RBUTTONDOWN
         ' pressione pulsante destro
         ' visualizza il popup
         frmMain.PopupMenu frmMain.mnuVassoio, vbPopupMenuRightAlign
      End Select
     End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, Wparam, Lparam)
End Function

