VERSION 5.00
Begin VB.Form frmMeseAnno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rielabora prestazioni"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox cboAnno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMeseAnno.frx":0000
         Left            =   720
         List            =   "frmMeseAnno.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboMese 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mese"
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
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anno"
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
         Index           =   32
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5175
      Begin VB.CommandButton cmdRielabora 
         Caption         =   "&Rielabora"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMeseAnno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmMeseAnno.frm
'
' <b>Descrizione</b>: Pannello per la scelta del mese e anno (per cambiare gli importi delle prescrizioni)
'
' @remarks
'
' @author
'
' @date 08/02/2011 21.01
Option Explicit

Private importoScontato As Boolean
Private importo As Single
Private keyId As Integer

Public Property Get getimportoScontato() As Boolean
    getimportoScontato = importoScontato
End Property

Public Property Let LetimportoScontato(ByVal vimportoScontato As Boolean)
    importoScontato = vimportoScontato
End Property

Public Property Get getkeyId() As Integer
    getkeyId = keyId
End Property

Public Property Let letKeyId(ByVal vKeyId As Integer)
    keyId = vKeyId
End Property

Public Property Get getimporto() As Single
    getimporto = importo
End Property

Public Property Let Letimporto(ByVal vimporto As Single)
    importo = vimporto
End Property

Private Sub cmdChiudi_Click()
     Unload Me
End Sub

'' Cambia tutti gli importi e importi scontati delle vecchie ricette del mese e anno selezionato
Private Sub cmdRielabora_Click()
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM PRESCRIZIONI WHERE CODICE_PRESTAZIONE=" & keyId & " AND CODICE_RICETTA IN (SELECT KEY FROM RICETTE WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " )", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        Do While Not rsDataset.EOF
            If importoScontato Then
                rsDataset("IMPORTO_SCONTATO") = importo
            Else
                rsDataset("IMPORTO") = importo
            End If
            rsDataset.MoveNext
        Loop
        MsgBox "Rielaborazione avvenuta con successo", vbInformation, "Rielaborazione prestazioni"
    Else
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbCritical, "Attenzione"
        rsDataset.Close
        Exit Sub
    End If
    rsDataset.Close
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    For i = 1 To 12
        cboMese.AddItem UCase(MonthName(i))
    Next i
    cboMese.ListIndex = Month(Now) - 1
End Sub
