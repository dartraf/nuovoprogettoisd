VERSION 5.00
Begin VB.Form frmDirettoreSanitario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Direttore Sanitario"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPassword 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtCognome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
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
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton cmdMemorizza 
         Cancel          =   -1  'True
         Caption         =   "&Memorizza"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1620
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
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
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmDirettoreSanitario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmDirettoreSanitario.frm
'
' <b>Descrizione</b>: Scheda Direttore Sanitario associata alla tab DIRETTORE_SANITARIO
'
' @remarks
'
' @author
'
' @date 05/02/2011 16.36
Option Explicit

'' rs della scheda
Dim rsDataset As Recordset
'' indica se si è in modifica
Dim modifica As Boolean

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    modifica = False
    Call CaricaScheda
End Sub

'' Carica la scheda nel form
Private Sub CaricaScheda()
    Set rsDataset = New Recordset
    rsDataset.Open "DIRETTORE_SANITARIO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        modifica = True
        txtCognome = rsDataset("COGNOME")
        txtNome = rsDataset("NOME")
    Else
        modifica = False
        
    End If
    Set rsDataset = Nothing
End Sub

'' Verifica prima di memorizzare se sono presenti tutti i dati
Private Function Completo() As Boolean
    Completo = False
    If txtCognome = "" Then
        MsgBox "Inserire il campo COGNOME", vbCritical, "Attenzione"
        Exit Function
    ElseIf txtNome = "" Then
        MsgBox "Inserire il campo NOME", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

Private Sub cmdMemorizza_Click()
    If Completo Then
        Set rsDataset = New Recordset
        If modifica Then
            rsDataset.Open "SELECT * FROM DIRETTORE_SANITARIO WHERE KEY=1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDataset("COGNOME") = (txtCognome)
            rsDataset("NOME") = (txtNome)
            rsDataset.Update
        Else
            rsDataset.Open "DIRETTORE_SANITARIO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.AddNew
            rsDataset("COGNOME") = (txtCognome)
            rsDataset("NOME") = (txtNome)
            rsDataset("KEY") = 1
            rsDataset.Update
        End If
        Set rsDataset = Nothing
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub txtNome_GotFocus()
    txtNome.BackColor = colArancione
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = vbWhite
End Sub

Private Sub txtCognome_GotFocus()
    txtCognome.BackColor = colArancione
End Sub

Private Sub txtcogNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCognome_LostFocus()
    txtCognome.BackColor = vbWhite
End Sub
