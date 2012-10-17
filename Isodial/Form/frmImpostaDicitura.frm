VERSION 5.00
Begin VB.Form frmImpostaDicitura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imposta Dicitura"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Imposta dicitura da stampare:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtDicitura 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         MaxLength       =   190
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
      Begin VB.CommandButton cmdMemorizza 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1380
      End
      Begin VB.CommandButton cmdAnnulla 
         Cancel          =   -1  'True
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
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmImpostaDicitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDataset As Recordset

Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    
    Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM INTESTAZIONE_FATTURA WHERE KEY=1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsDataset("DICITURA_ESAMI_PERIODICI") = (txtDicitura)
        rsDataset.Update
    Set rsDataset = Nothing
        
    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"

End Sub

Private Sub Form_Load()

    txtDicitura_GotFocus

    Set rsDataset = New Recordset
    rsDataset.Open "INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        txtDicitura = rsDataset("DICITURA_ESAMI_PERIODICI") & ""
    End If
    Set rsDataset = Nothing
    
End Sub

Private Sub txtDicitura_GotFocus()
    txtDicitura.BackColor = colArancione
End Sub

Private Sub txtDicitura_LostFocus()
    txtDicitura.BackColor = vbWhite
End Sub
