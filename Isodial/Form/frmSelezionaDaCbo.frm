VERSION 5.00
Begin VB.Form frmSelezionaDaCbo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleziona gruppo"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cboDati 
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
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Il trasferimento potrebbe richiedere diversi minuti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label lblNomeDato 
         AutoSize        =   -1  'True
         Caption         =   "Gruppo"
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
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   780
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
      Begin VB.CommandButton cmdNuovo 
         Cancel          =   -1  'True
         Caption         =   "&Nuovo"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1380
      End
      Begin VB.CommandButton cmdSeleziona 
         Caption         =   "&Seleziona"
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
         TabIndex        =   3
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
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmSelezionaDaCbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nomeTabella As String

Private Sub cmdAnnulla_Click()
    tSelezionaDaCbo.valoreSelezionato = 0
    Unload Me
End Sub

Private Sub cmdNuovo_Click()
    Dim v_nomi(1 To 2) As Variant
    Dim v_val(1 To 2) As Variant
    Dim num As Integer
    Dim primo As Boolean
    Dim rsDataset As New Recordset
    
    primo = True
    tInput.mantieniDati = False
    tInput.Tipo = tpIESAMI
    Do
        If Not primo Then
            MsgBox "Il gruppo inserito è già presente", vbCritical, "Attenzione"
            tInput.mantieniDati = True
        End If
        Unload frmInput
        frmInput.Show 1
        primo = False
    Loop While Esiste(frmTipiEsamiLab.flxNomi, 1, 0, tInput.v_valori(1))
    
    If Not (tInput.v_valori(1) = "") Then
        v_nomi(1) = "KEY"
        v_nomi(2) = "NOME"
        num = GetNumero("GRUPPI_ESAMI")
        v_val(1) = num
        v_val(2) = tInput.v_valori(1)
        
        Set rsDataset = New Recordset
        rsDataset.Open "GRUPPI_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_nomi, v_val
        rsDataset.Update
        Set rsDataset = Nothing
        
        tSelezionaDaCbo.nuovoInserimento = True
        tSelezionaDaCbo.valoreSelezionato = num
        Unload Me
    End If
End Sub

Private Sub cmdSeleziona_Click()
    tSelezionaDaCbo.nuovoInserimento = False
    If cboDati.ListIndex = -1 Then
        tSelezionaDaCbo.valoreSelezionato = -1
    Else
        tSelezionaDaCbo.valoreSelezionato = cboDati.ItemData(cboDati.ListIndex)
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    cboDati.SetFocus
End Sub

Private Sub Form_Load()
    Select Case tSelezionaDaCbo.tipoCampo
        Case tpGRUPPI_ESAMI: nomeTabella = "GRUPPI_ESAMI"
    End Select
    
    Call RicaricaComboBox("SELECT NOME, KEY FROM " & nomeTabella & " WHERE NOT KEY=" & tSelezionaDaCbo.valoreDaEvitare, "NOME", cboDati)
End Sub
