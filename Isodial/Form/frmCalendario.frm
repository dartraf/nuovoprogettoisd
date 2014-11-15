VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendario 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "&Annulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "&Conferma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin MSComCtl2.MonthView calendario 
      Height          =   2820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   51118082
      TitleBackColor  =   -2147483645
      CurrentDate     =   38969
      MinDate         =   2
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmCalendario.frm
'
' <b>Descrizione</b>: Mostra un calendario
'
' @remarks
'
' @author
'
' @date 01/02/2011 21.49
Option Explicit

'' Carica il valore della data scelta nella var publica laData
Private Sub calendario_DateDblClick(ByVal DateDblClicked As Date)
    laData = calendario.Value
    Unload Me
End Sub

Private Sub calendario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

'' Setta le impostazioni iniziali e la posizione del form
Private Sub Form_Load()
    Dim PuntoX As Integer
    Dim PuntoY As Integer
    Call PosizioneCursore(PuntoX, PuntoY)
    Me.Top = PuntoY
    Me.Left = PuntoX
    If Me.Left + Me.Width > frmMain.Width Then
        Me.Left = frmMain.Width - Me.Width - 300
    End If
    If Me.Top + Me.Height > frmMain.Height Then
        Me.Top = frmMain.Height - Me.Height - 300
    End If
    Me.Top = IIf(Me.Top > 500, Me.Top, 1000)
    Me.Left = IIf(Me.Left > 500, Me.Left, 1000)
    calendario.Value = date
End Sub

Private Sub calendario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nRet As Long
    If Y < 100 Then
        ReleaseCapture
        nRet = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub cmdAnnulla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    laData = ""
    Unload Me
End Sub

Private Sub cmdConferma_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    laData = calendario.Value
    Unload Me
End Sub

