VERSION 5.00
Begin VB.Form frmTabSingoloInput 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfo 
      Height          =   705
      Left            =   120
      TabIndex        =   3
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   4095
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
         TabIndex        =   5
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   600
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmTabSingoloInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intTipoTabSingolo As enumTipoTabSingolo
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
    Select Case intTipoTabSingolo
        Case enumTipoTabSingolo.AGO
            strNomeTabella = "AGO"
            strNomeElemento = "L'ago"
            Me.Caption = "Inserimento Ago"
            lblNome.Caption = "Ago"
        Case enumTipoTabSingolo.ANTICOAGULANTI
            strNomeTabella = "ANTICOAGULANTI"
            strNomeElemento = "L'anticoagulante"
            Me.Caption = "Inserimento Anticoagulante"
            lblNome.Caption = "Anticoagulante"
        Case enumTipoTabSingolo.filtro
            strNomeTabella = "FILTRI"
            strNomeElemento = "Il filtro"
            Me.Caption = "Inserimento Filtro"
            lblNome.Caption = "Filtro"
        Case enumTipoTabSingolo.LINEE
            strNomeTabella = "LINEE"
            strNomeElemento = "Le linee"
            Me.Caption = "Inserimento Linee"
            lblNome.Caption = "Linee"
        Case enumTipoTabSingolo.Medicinali
            strNomeTabella = "MEDICINALI"
            strNomeElemento = "Il farmaco"
            Me.Caption = "Inserimento Farmaco"
            lblNome.Caption = "Farmaco"
        Case enumTipoTabSingolo.ORGANO
            strNomeTabella = "ORGANI"
            strNomeElemento = "L'organo/Apparato"
            Me.Caption = "Inserimento Organo/Apparato"
            lblNome.Caption = "Organo"
        Case enumTipoTabSingolo.TITOLIDIARIO
            strNomeTabella = "TITOLI_DIARIO"
            strNomeElemento = "Il titolo di diario clinico"
            Me.Caption = "Inserimento Titolo diario"
            lblNome.Caption = "Titolo"
    End Select
End Sub

Private Function ControlloValori()
    If txtNome.Text = "" Then
        MsgBox "Inserimento Dati Obbligatori.", vbExclamation, Me.Caption
        ControlloValori = False
    Else
        ControlloValori = True
    End If
End Function

Private Function ControlloDuplicato() As Boolean
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "Select    count(Key) as Totale " & _
            "From " & strNomeTabella & " " & _
            "Where      Nome like '" & Apostrophe(txtNome.Text) & "'"
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
    
    Call SuperUcase(Me)

    intIDInserito = GetNumero(strNomeTabella)
    v_Nomi = Array("KEY", "NOME")
    v_Val = Array(intIDInserito, txtNome.Text)
  
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

