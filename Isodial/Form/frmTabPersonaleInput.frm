VERSION 5.00
Begin VB.Form frmTabPersonaleInput 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfo 
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6015
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   1
         Top             =   720
         Width           =   3735
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   3735
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblNome 
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   630
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6015
      Begin VB.CommandButton cmdInserisci 
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
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1380
      End
      Begin VB.CommandButton cmdAnnulla 
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
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame fraAlreInfoMediciDialisi 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtCodiceAlbo 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Iscrizione Albo"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame fraAltreInfoInfermieri 
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optMansione 
         Caption         =   "Infermiere professionale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optMansione 
         Caption         =   "Coordinatore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipologia"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmTabPersonaleInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intTipoTabPersonale As enumTipoTabPersonale
Public blnRefresh As Boolean
Public intIDInserito As Integer
Dim strNomeTabella As String
Dim strNomeElemento As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdAnnulla_Click
    End If
End Sub

Private Sub Form_Load()
    blnRefresh = False
    Select Case intTipoTabPersonale
        Case enumTipoTabPersonale.MEDICI_DIALISI
            strNomeTabella = "MEDICI_DIALISI"
            strNomeElemento = "Il medico dialisi"
            Me.Caption = "Inserimento Medico Dialisi"
            fraAlreInfoMediciDialisi.Visible = True
            fraAlreInfoMediciDialisi.Enabled = True
            fraPulsanti.Top = fraAlreInfoMediciDialisi.Top + fraAlreInfoMediciDialisi.Height - 100
        Case enumTipoTabPersonale.INFERMIERI
            strNomeTabella = "INFERMIERI"
            strNomeElemento = "L'infermiere"
            Me.Caption = "Inserimento Infermiere"
            fraAltreInfoInfermieri.Visible = True
            fraAltreInfoInfermieri.Enabled = True
            fraPulsanti.Top = fraAltreInfoInfermieri.Top + fraAltreInfoInfermieri.Height - 100
        Case enumTipoTabPersonale.MEDICI_REFERTANTI
            strNomeTabella = "MEDICI_REFERTANTI"
            strNomeElemento = "Il medico refertante"
            Me.Caption = "Inserimento Medico Refertante"
        Case enumTipoTabPersonale.PSICOLOGI
            strNomeTabella = "PSICOLOGI"
            strNomeElemento = "Lo psicologo"
            Me.Caption = "Inserimento Psicologo"
    End Select
    
    Me.Height = fraPulsanti.Top + fraPulsanti.Height + 400
    fraPulsanti.ZOrder 1
End Sub

Private Function ControlloValori()
    Dim strNome As String
    
    If txtCognome.Text = "" Then
        strNome = "COGNOME"
        txtCognome.SetFocus
    ElseIf txtNome.Text = "" Then
        strNome = "NOME"
        txtNome.SetFocus
    End If

    If strNome = "" Then
        ControlloValori = True
    Else
        MsgBox "Il campo " & strNome & " è obbligatorio.", vbExclamation, Me.Caption
        ControlloValori = False
    End If
End Function

Private Function ControlloDuplicato() As Boolean
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "Select    count(Key) as Totale " & _
            "From " & strNomeTabella & " " & _
            "Where      Cognome like '" & txtCognome.Text & "' and" & _
            "           Nome like '" & txtNome.Text & "'"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly
    If rsDataset("Totale") <> 0 Then
        MsgBox strNomeElemento & " è gia presente in archivio.", vbExclamation, Me.Caption
        ControlloDuplicato = False
    Else
        ControlloDuplicato = True
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Function

Private Sub Memorizza()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim rsDataset As New Recordset

    intIDInserito = GetNumero(strNomeTabella)
    Select Case intTipoTabPersonale
        Case enumTipoTabPersonale.MEDICI_DIALISI
            v_Nomi = Array("KEY", "COGNOME", "NOME", "CODICE_ALBO")
            v_Val = Array(intIDInserito, txtCognome.Text, txtNome.Text, txtCodiceAlbo.Text)
        Case enumTipoTabPersonale.INFERMIERI
            v_Nomi = Array("KEY", "COGNOME", "NOME", "MANSIONE")
            v_Val = Array(intIDInserito, txtCognome.Text, txtNome.Text, IIf(optMansione(0).Value = True, 1, 2))
        Case enumTipoTabPersonale.MEDICI_REFERTANTI, enumTipoTabPersonale.PSICOLOGI
            v_Nomi = Array("KEY", "COGNOME", "NOME")
            v_Val = Array(intIDInserito, txtCognome.Text, txtNome.Text)
    End Select
  
    rsDataset.Open strNomeTabella, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
    rsDataset.AddNew v_Nomi, v_Val
    rsDataset.Update
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

Private Sub cmdAnnulla_Click()
    blnRefresh = False
    Unload Me
End Sub

Private Sub cmdInserisci_Click()
    If ControlloValori Then
        If ControlloDuplicato Then
            Call Memorizza
            blnRefresh = True
            Unload Me
        End If
    End If
End Sub

