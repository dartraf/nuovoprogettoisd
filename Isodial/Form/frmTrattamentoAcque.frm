VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmTrattamentoAcque 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attuazione protocollo di trattamento delle acque"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox cboAttivita 
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
         ItemData        =   "frmTrattamentoAcque.frx":0000
         Left            =   1800
         List            =   "frmTrattamentoAcque.frx":0010
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox cboEsito 
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
         ItemData        =   "frmTrattamentoAcque.frx":006B
         Left            =   1800
         List            =   "frmTrattamentoAcque.frx":0075
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtEseguitoDa 
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2235
         Width           =   735
      End
      Begin DataTimeBox.uDataTimeBox oDataTime 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin DataTimeBox.uDataTimeBox oDataTime 
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   17
         Top             =   2160
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo attività"
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
         TabIndex        =   12
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
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
         TabIndex        =   11
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
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
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   1725
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eseguito da"
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
         Index           =   11
         Left            =   120
         TabIndex        =   9
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estremi referto:      n°"
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
         Index           =   12
         Left            =   120
         TabIndex        =   8
         Top             =   2235
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "data"
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
         Index           =   13
         Left            =   3720
         TabIndex        =   7
         Top             =   2235
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   6495
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
         Enabled         =   0   'False
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
         TabIndex        =   14
         Top             =   240
         Width           =   1335
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
         Left            =   5160
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
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
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   -120
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmTrattamentoAcque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTrattamento As Recordset
Dim modifica As Integer
Dim keyId As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i  As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    modifica = False
    Call EliminaScansioniSospese("SCAN_TRATT_ACQUE")
End Sub

Private Sub PulisciTutto()
    oDataTime(0).data = ""
    cmdStampa.Enabled = False
End Sub

Private Sub CaricaScheda()
    Dim data As Date
    If oDataTime(0).data <> "" Then
        Set rsTrattamento = New Recordset
        ' la data americana
        data = oDataTime(0).DataAmericana
        rsTrattamento.Open "SELECT * FROM MON_TRAT_ACQUE WHERE (DATA=#" & data & "#)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsTrattamento.EOF And rsTrattamento.BOF) Then
            cboAttivita.ListIndex = rsTrattamento("TIPO_ATTIVITA")
            txtEseguitoDa = rsTrattamento("ESEGUITA_DA") & ""
            cboEsito.ListIndex = rsTrattamento("ESITO") - 1
            txtNumero = IIf(rsTrattamento("ESTREMI_NUM") = 0, "", rsTrattamento("ESTREMI_NUM"))
            oDataTime(1).data = rsTrattamento("ESTREMI_DATA") & ""
            keyId = rsTrattamento("KEY")
            modifica = True
        Else
            modifica = False
        End If
        Set rsTrattamento = Nothing
    End If
End Sub

Private Sub Pulisci()
    oDataTime(1).Pulisci
    modifica = False
    keyId = 0
    Call PulisciForm(Me)
    Call EliminaScansioniSospese("SCAN_TRATT_ACQUE")
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim numKey As Integer
    
    If oDataTime(0).data = "" Then
        MsgBox "Selezionare la data", vbInformation, "Informazione"
        Exit Sub
    End If
        
    If modifica Then
        numKey = keyId
    Else
        numKey = GetNumero("MON_TRAT_ACQUE")
    End If
    v_Nomi = Array("KEY", "DATA", "TIPO_ATTIVITA", "ESEGUITA_DA", "ESITO", "ESTREMI_NUM")
    v_Val = Array(numKey, oDataTime(0).data, cboAttivita.ListIndex, txtEseguitoDa, cboEsito.ListIndex + 1, IIf(txtNumero = "", 0, txtNumero))
    Set rsTrattamento = New Recordset
    
    If modifica Then
        rsTrattamento.Open "SELECT * FROM MON_TRAT_ACQUE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        ' gestisce la data separatamente solo se questa e stata inserita
        rsTrattamento("ESTREMI_DATA") = Null
        If oDataTime(1).data <> "" Then
            rsTrattamento("ESTREMI_DATA") = oDataTime(1).data
        End If
        rsTrattamento.Update v_Nomi, v_Val
    Else
        rsTrattamento.Open "MON_TRAT_ACQUE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsTrattamento.AddNew v_Nomi, v_Val
        ' gestisce la data separatamente solo se questa e stata inserita pulendola cmq
        rsTrattamento("ESTREMI_DATA") = Null
        If oDataTime(1).data <> "" Then
            rsTrattamento("ESTREMI_DATA") = oDataTime(1).data
        End If
        rsTrattamento.Update
        rsTrattamento.Close
        ' controlla eventuali scansioni memorizzate in sospeso
        rsTrattamento.Open "SELECT * FROM SCAN_TRATT_ACQUE WHERE CODICE_SCHEDA=0", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        Do While Not rsTrattamento.EOF
            rsTrattamento("CODICE_SCHEDA") = numKey
            rsTrattamento.Update
            rsTrattamento.MoveNext
        Loop
        rsTrattamento.Close
    End If
    Set rsTrattamento = Nothing
    
    Call PulisciTutto
    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
End Sub

Private Sub cmdGestioneReferti_Click()
    Unload frmGestioneDocumentiEsterni
    Load frmGestioneDocumentiEsterni
    frmGestioneDocumentiEsterni.LetCodicePaziente = 0
    If modifica Then
        frmGestioneDocumentiEsterni.letcodiceRecord = keyId
    Else
        frmGestioneDocumentiEsterni.letcodiceRecord = 0
    End If
    frmGestioneDocumentiEsterni.LetNomeFile = T_AC & " " & Replace(oDataTime(0).data, "/", "-")
    tDocumentiEsterni = tpSCANTRATTAMENTOACQUE
    frmGestioneDocumentiEsterni.Show 1
End Sub

Private Sub cmdStampa_Click()
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
        
    strSql = "SHAPE APPEND  NEW adVarChar (10) as DATA, " & _
                    "       NEW adVarChar (40) as TIPO_ATTIVITA, " & _
                    "       NEW adVarChar (40) as ESEGUITO, " & _
                    "       NEW adVarChar (10) as ESITO, " & _
                    "       NEW adVarChar (6)  as ESTREMI_REFERTO, " & _
                    "       NEW adVarChar (10) as DATA_ESTREMI_REFERTO "
                    
                                              
                                              
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
           
 
    If oDataTime(0).data = "" Then
        MsgBox "Selezionare la data", vbInformation, "Informazione"
        Exit Sub
    Else
        Set rsTrattamento = New Recordset
        rsTrattamento.Open "MON_TRAT_ACQUE" & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
        Do While Not rsTrattamento.EOF
            With rsMain
                .AddNew
                .Fields("DATA") = rsTrattamento("DATA")
                .Fields("TIPO_ATTIVITA") = Choose(rsTrattamento("TIPO_ATTIVITA") + 1, "Disinfezione impianto ed anello", "esami chimico - fisici", "esami microbiologici", "Lal test")
                .Fields("ESEGUITO") = rsTrattamento("ESEGUITA_DA") & ""
                .Fields("ESITO") = Choose(rsTrattamento("ESITO"), "NEGATIVO", "POSITIVO")
                .Fields("ESTREMI_REFERTO") = IIf(rsTrattamento("ESTREMI_NUM") = 0, "", rsTrattamento("ESTREMI_NUM"))
                .Fields("DATA_ESTREMI_REFERTO") = rsTrattamento("ESTREMI_DATA") & ""
            End With
            rsTrattamento.MoveNext
        Loop
        rsTrattamento.Close
        Set rsTrattamento = Nothing
    End If
    Set rptTrattamentoAcque.DataSource = rsMain
    rptTrattamentoAcque.TopMargin = 0
    rptTrattamentoAcque.BottomMargin = 0
    rptTrattamentoAcque.PrintReport True, rptRangeAllPages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EliminaScansioniSospese("SCAN_TRATT_ACQUE")
End Sub

Private Sub oDataTime_OnDataChange(Index As Integer)
    If Index = 0 Then
        Call Pulisci
        If oDataTime(Index).data <> "" Then
            Call CaricaScheda
        End If
    End If
    cmdStampa.Enabled = True
End Sub

Private Sub oDataTime_OnDataClick(Index As Integer)
    oDataTime(Index).Pulisci
    laData = ""
End Sub

Private Sub oDataTime_OnElencaClick(Index As Integer)
    ' setta le variabili che saranno viste dal frmElencaDate
    tElenca.Tipo = tpMON_TRAT_ACQUE
    tElenca.condizione = ""
    frmElencaDate.Show 1
    If laData <> "" Then oDataTime(0).data = laData
End Sub

Private Sub txtNumero_LostFocus()
    txtNumero.BackColor = vbWhite
End Sub

Private Sub txtEseguitoDa_GotFocus()
    txtEseguitoDa.BackColor = colArancione
End Sub

Private Sub txtEseguitoDa_LostFocus()
    txtEseguitoDa.BackColor = vbWhite
End Sub

Private Sub txtNumero_GotFocus()
    txtNumero.BackColor = colArancione
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

