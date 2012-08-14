VERSION 5.00
Begin VB.Form frmEsamiPeriodiciInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserisci esame periodico"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTerapia 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox chkTutto 
         Caption         =   "Seleziona tutto il gruppo di esami"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   4695
      End
      Begin VB.OptionButton optTipoEsame 
         Caption         =   "Esami di Laboratorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.OptionButton optTipoEsame 
         Caption         =   "Esami Strumentali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboEsami 
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
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   4335
      End
      Begin VB.ComboBox cboGruppi 
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
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   640
         Width           =   4335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Esami"
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
         TabIndex        =   5
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label lblTipoGruppo 
         AutoSize        =   -1  'True
         Caption         =   "Raggruppamento"
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
         TabIndex        =   3
         Top             =   660
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   6735
      Begin VB.OptionButton optPeriodoProblemi 
         Caption         =   "Se problemi clinici"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optPeriodoBimestrale 
         Caption         =   "Bimestrale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optPeriodoAnnuali 
         Caption         =   "Annuali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPeriodoSemestrali 
         Caption         =   "Semestrali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optPeriodoTrimestrali 
         Caption         =   "Trimestrali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optPeriodoMensili 
         Caption         =   "Mensili"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frequenza:"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   6735
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
         Left            =   3720
         TabIndex        =   17
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
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmEsamiPeriodiciInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmEsamiInput.frm
'
' <b>Descrizione</b>: Pannello per l'input degli esami periodici
'
' @remarks
'
' @author
'
' @date 05/02/2011 17.59
Option Explicit

'' rs della scheda
Dim rsDataset As Recordset
Dim m_intPeriodo As tipoPeriodo
Public Sub LetPeriodo(Vperiodo As tipoPeriodo)
    m_intPeriodo = Vperiodo
End Sub

Private Sub Form_Activate()
    optTipoEsame(0).Value = True
    
    tEsamiPeriodici.interoGruppo = -1
    tEsamiPeriodici.periodo = -1
    tEsamiPeriodici.codiceAssociazione = 0
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2 - 500
    Me.Left = (Screen.Width - Me.Width) / 2
    Call TakeCloseOff(Me.hWnd)
    Select Case m_intPeriodo
        Case tipoPeriodo.tpMENSILE: optPeriodoMensili.Value = True
        Case tipoPeriodo.tpBIMESTRALE: optPeriodoBimestrale.Value = True
        Case tipoPeriodo.tpTRIMESTRALE: optPeriodoTrimestrali.Value = True
        Case tipoPeriodo.tpSEMESTRALE: optPeriodoSemestrali.Value = True
        Case tipoPeriodo.tpANNUALE: optPeriodoAnnuali.Value = True
        Case tipoPeriodo.tpPROBLEMI: optPeriodoProblemi.Value = True
        Case Else: optPeriodoMensili.Value = True
    End Select
End Sub

'' Restituisce il periodo selezionato dagli opt
Private Function GetOptperiodo() As Integer
    If optPeriodoMensili.Value Then GetOptperiodo = tipoPeriodo.tpMENSILE
    If optPeriodoBimestrale.Value Then GetOptperiodo = tipoPeriodo.tpBIMESTRALE
    If optPeriodoTrimestrali.Value Then GetOptperiodo = tipoPeriodo.tpTRIMESTRALE
    If optPeriodoSemestrali.Value Then GetOptperiodo = tipoPeriodo.tpSEMESTRALE
    If optPeriodoAnnuali.Value Then GetOptperiodo = tipoPeriodo.tpANNUALE
    If optPeriodoProblemi.Value Then GetOptperiodo = tipoPeriodo.tpPROBLEMI
End Function

'' Filtra solo gli esami di lab o quelli strumentali
Private Sub Filtra()
    Dim strSql As String
    If optTipoEsame(0).Value Then
        ' Carica tutti gli esami dell'organo scelto
        Call RicaricaComboBox("SELECT * FROM ESAMI WHERE CODICE_ORGANO=" & cboGruppi.ItemData(cboGruppi.ListIndex), "NOME", cboEsami)
    Else
        ' Carica tutti gli esami del solo gruppo scelto
        strSql = "SELECT    VOCI_ESAMI.NOME, VOCI_ESAMI.KEY " & _
                "FROM       (ASSOCIAZIONE_ESAMI_LAB " & _
                "           INNER JOIN VOCI_ESAMI ON VOCI_ESAMI.KEY=ASSOCIAZIONE_ESAMI_LAB.CODICE_ESAME) " & _
                "WHERE      CODICE_GRUPPO=" & cboGruppi.ItemData(cboGruppi.ListIndex)
        Call RicaricaComboBox(strSql, "NOME", cboEsami)
    End If
End Sub

'' Richiama Filtra
Private Sub cboGruppi_Click()
    Call Filtra
End Sub

Private Sub optTipoEsame_Click(Index As Integer)
    cboEsami.Clear
    If Index = 1 Then
        Call RicaricaComboBox("GRUPPI_ESAMI", "NOME", cboGruppi)
        lblTipoGruppo = "Raggruppamento"
    Else
        Call RicaricaComboBox("ORGANI", "NOME", cboGruppi)
        lblTipoGruppo = "Organo"
    End If
End Sub

Private Sub cmdAnnulla_Click()
    tEsamiPeriodici.interoGruppo = -1
    Unload Me
End Sub

'' Passa il valori al form degli esami periodici
' Se viene scelto un esame di lab codiceAssociazione è negativo
' Se viene scelto tutto il gruppo interoGruppo > 0
Private Sub cmdInserisci_Click()
    If cboEsami.ListIndex = -1 And cboGruppi.ListIndex = -1 Then
        MsgBox "Selezionare almeno un elemento", vbCritical, "Attenzione"
        Exit Sub
    End If

    If Not (chkTutto.Value = Unchecked And cboEsami.ListIndex = -1) Then
        Set rsDataset = New Recordset
        If optTipoEsame(0).Value Then
            If chkTutto.Value = Checked Then
                tEsamiPeriodici.interoGruppo = 1
                tEsamiPeriodici.codiceAssociazione = cboGruppi.ItemData(cboGruppi.ListIndex)
            Else
                tEsamiPeriodici.interoGruppo = 0
                rsDataset.Open "SELECT KEY FROM ESAMI WHERE KEY=" & cboEsami.ItemData(cboEsami.ListIndex) & " AND CODICE_ORGANO=" & cboGruppi.ItemData(cboGruppi.ListIndex), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                tEsamiPeriodici.codiceAssociazione = rsDataset("KEY")
                rsDataset.Close
            End If
        Else
            If chkTutto.Value = Checked Then
                tEsamiPeriodici.interoGruppo = 1
                tEsamiPeriodici.codiceAssociazione = -cboGruppi.ItemData(cboGruppi.ListIndex)
            Else
                tEsamiPeriodici.interoGruppo = 0
                rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & cboGruppi.ItemData(cboGruppi.ListIndex) & " AND CODICE_ESAME=" & cboEsami.ItemData(cboEsami.ListIndex), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                tEsamiPeriodici.codiceAssociazione = -rsDataset("CODICE_ESAME")
                rsDataset.Close
            End If
        End If
        Set rsDataset = Nothing
        tEsamiPeriodici.periodo = GetOptperiodo
        Unload Me
    Else
        MsgBox "Selezionare l'esame", vbCritical, "Attenzione"
        Exit Sub
    End If
End Sub

