VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Verifìca password"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3345
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPassword 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtConferma 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Inserisci la password"
         Top             =   2280
         Width           =   2205
      End
      Begin VB.TextBox txtNuova 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Inserisci la password"
         Top             =   1440
         Width           =   2205
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Inserisci la password"
         Top             =   600
         Width           =   2205
      End
      Begin VB.Image Image2 
         Height          =   330
         Index           =   2
         Left            =   120
         Picture         =   "frmPass.frx":030A
         Top             =   2280
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   330
         Index           =   1
         Left            =   120
         Picture         =   "frmPass.frx":0494
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conferma password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuova password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inserisci password corrente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2865
      End
      Begin VB.Image Image2 
         Height          =   330
         Index           =   0
         Left            =   120
         Picture         =   "frmPass.frx":061E
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3135
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdAnnulla 
         Cancel          =   -1  'True
         Caption         =   "&Annulla"
         CausesValidation=   0   'False
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
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPass.frm
'
' <b>Descrizione</b>: Pannello per la verifica e il cambio di password degli utenti
'
' @remarks
'
' @author
'
' @date 22/02/2011 18.37
Option Explicit

'' indica se la pass è stata inserita bene
Private risPass As Boolean
Dim password As String

Public Property Get GetRisPass() As Boolean
    GetRisPass = risPass
End Property

Public Property Let LetRisPass(ByVal vRisPass As Boolean)
    risPass = vRisPass
End Property

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

'' Verifica se la password inserita è corretta
'
' @param nome stringa inserita da verificare con quella del db
' @return true se verificata
Private Function verificaPass(nome As String) As Boolean
    If nome = password Then
        verificaPass = True
    Else
        MsgBox "Password errata", vbCritical, "Attenzione"
        verificaPass = False
    End If
End Function

Private Sub cmdAnnulla_Click()
    risPass = False
    Me.Hide
End Sub

'' Cambia o verifica la password dell'utente corrente
Private Sub cmdOK_Click()
    If verificaPass(txtPassword.Text) Then
        If tipoPass.Tipo = tCAMBIA Then
            If txtConferma = txtNuova And Len(txtConferma) >= 8 And txtConferma <> txtPassword Then
                password = txtConferma
                Dim rsDataset As New Recordset
                rsDataset.Open "SELECT * FROM LOGIN WHERE KEY=" & tipoPass.key, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                rsDataset.Update "PASSWORD", password
                rsDataset.Update "DATA", date
                Set rsDataset = Nothing
                risPass = True
                Me.Hide
            Else
                If txtConferma <> txtNuova Then
                    MsgBox "Password di conferma errata", vbCritical, "Cambio password"
                ElseIf txtConferma = txtPassword Then
                    MsgBox "Impossibile inserire la stessa password", vbCritical, "Cambio password"
                Else
                    MsgBox "Lunghezza password inferiore a 8 caratteri", vbCritical, "Cambio password"
                End If
                txtConferma.Text = ""
                txtConferma.SetFocus
            End If
        Else
            risPass = True
            Me.Hide
        End If
    Else
        risPass = False
        txtPassword.Text = ""
        txtPassword.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If tipoPass.Tipo = tCAMBIA Then
        fraPassword.Height = 2775
        fraPulsanti.Top = fraPassword.Top + fraPassword.Height - 135
        Me.Height = fraPulsanti.Top + fraPulsanti.Height + 480
        Me.Caption = "Cambia password"
    End If
    password = tipoPass.password
End Sub

'' Evita di chiudere il form con alt f4
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) And KeyCode = vbKeyF4 Then
        KeyCode = 0
    End If
End Sub
