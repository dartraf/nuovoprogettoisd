VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmInserisciManutenzione 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraManutenzione 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin DataTimeBox.uDataTimeBox oDataScadenzaManutenzione 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.CheckBox chkSicurezza 
         Caption         =   "Verifica Sicurezza Elettrica"
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
         Left            =   5880
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox chkFunzionalità 
         Caption         =   "Verifica Funzionalità Meccanica"
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
         Left            =   5880
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cboDescrizone 
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
         Index           =   0
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   6855
      End
      Begin DataTimeBox.uDataTimeBox oDataEffettivaManutenzione 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataRichiestaManutenzione 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Effettiva Manut."
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   750
         Width           =   2145
      End
      Begin VB.Label lblDescrizioneManutenzione 
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
         Top             =   1320
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Scadenza Manut."
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
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Richiesta Manut."
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
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   2280
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   9975
      Begin VB.TextBox txtNumeroDocumneto 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboDettagliIntervento 
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
         Index           =   1
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dettagli Intervento"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "N° Rif. Doc. di Lavoro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   9975
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa Richiesta Intervento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4560
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtTipoManutenzione 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
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
         Height          =   600
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   1455
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
         Height          =   600
         Left            =   8640
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmInserisciManutenzione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsManutenzione As Recordset
Dim rsCercaManutenzione As Recordset
Dim rsDataRevisioneFunzionale As Recordset
Dim rsDataRevisioneSicurezza As Recordset
Dim NumeroApparato As Integer
Dim ModificaApparato As Boolean
Dim ProxRevFun As Date
Dim ProxRevSic As Date

'SelezionatoManutenzione = False    Inserisco la nuova manutenzione anche se seleziono con un click l' apparato
'SelezionatoManutenzione = True     Seleziono la manutenzione desiderata e la carico

'Private Sub cboDescrizone_GotFocus(Index As Integer)
'    cboDescrizone(0).BackColor = colArancione
'End Sub

Private Sub cboDescrizone_LostFocus(Index As Integer)
    If Len(cboDescrizone(0)) > 120 Then
        MsgBox "Impossibile memorizzare più di 120 caratteri", vbCritical, "Attenzione"
        cboDescrizone(0).Text = ""
        cboDescrizone(0).SetFocus
        Exit Sub
    End If
    
    If cboDescrizone(0).Text <> "" Then
        Call GestisciNuovoApparato("DESCRIZIONE_MANUTENZIONE", cboDescrizone(0))
    End If

'    cboDescrizone(0).BackColor = vbWhite
End Sub

'Private Sub cboDettagliIntervento_GotFocus(Index As Integer)
'    cboDettagliIntervento(1).BackColor = colArancione
'End Sub

Private Sub cboDettagliIntervento_LostFocus(Index As Integer)
    If Len(cboDettagliIntervento(1)) > 120 Then
        MsgBox "Impossibile memorizzare più di 120 caratteri", vbCritical, "Attenzione"
        cboDettagliIntervento(1).Text = ""
        cboDettagliIntervento(1).SetFocus
        Exit Sub
    End If
    
    If cboDettagliIntervento(1).Text <> "" Then
        Call GestisciNuovoApparato("DETTAGLIO_MANUTENZIONE", cboDettagliIntervento(1))
    End If

 '   cboDettagliIntervento(1).BackColor = vbWhite
End Sub

Private Sub cmdChiudi_Click()
    If KeyReturnManutenzione > 0 Then
        SelezionatoManutenzione = False
        Unload frmInserisciManutenzione
    Else
        SelezionatoManutenzione = False
        KeyReturnManutenzione = -2
        Unload frmInserisciManutenzione
    End If
    rptRichiestaIntervento.Sections("Intestazione").Controls("lblRagioneSociale").Caption = ""

End Sub

Private Sub cmdMemorizza_Click()
Dim v_Nomi() As Variant
Dim v_Val() As Variant
Dim numKey As Integer
               
    Set rsManutenzione = New Recordset
    
    If tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA Then
        If oDataRichiestaManutenzione(0).txtBox = "" Then
            MsgBox "Inserire la Data di Richiesta Manutenzione", vbInformation, "Informazione"
            Exit Sub
        ElseIf CDate(oDataRichiestaManutenzione(0).data) > date Then
            MsgBox "La Data della Richiesta Straordinaria non può essere successiva alla Data Odierna", vbInformation, "Informazione"
            Exit Sub
        ElseIf CDate(oDataEffettivaManutenzione(1).data) > date Then
            MsgBox "La Data di Effettiva Manutenzione non può essere successiva alla Data Odierna", vbInformation, "Informazione"
            Exit Sub
        ElseIf cboDescrizone(0).Text = "" Then
            MsgBox "Inserire la Motivazione della Richiesta", vbInformation, "Informazione"
            Exit Sub
        End If
    
    ElseIf tTabellaManutenzione = tpMANUTENZIONEORDINARIA Then
        If oDataScadenzaManutenzione(1).txtBox = "" Then
            MsgBox "Inserire la Data di Scadenza Manutenzione", vbInformation, "Informazione"
            Exit Sub
        ElseIf CDate(oDataScadenzaManutenzione(1).data) > date Then
            MsgBox "La Data di Scadenza Manutenzione non può essere successiva alla Data Odierna", vbInformation, "Informazione"
            Exit Sub
        ElseIf CDate(oDataEffettivaManutenzione(1).data) > date Then
            MsgBox "La Data di Effettiva Manutenzione non può essere superiore alla Data Odierna", vbInformation, "Informazione"
            Exit Sub
        ElseIf chkFunzionalità.Value = Unchecked And chkSicurezza.Value = Unchecked Then
            MsgBox "Selezionare la Funzionalità o la Sicurezza", vbInformation, "Informazione"
            Exit Sub
        ElseIf cboDescrizone(0).Text = "" Then
            MsgBox "Inserire la Descrizione della Manutenzione", vbInformation, "Informazione"
            Exit Sub
        End If

    End If
            
    If SelezionatoManutenzione = False Then
        numKey = GetNumero("MANUTENZIONE_APPARATI")
    Else
        numKey = KeyReturnManutenzione
    End If
    
    If tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA Then
        'sostituisce l'apostrofo con `
        cboDescrizone(0).Text = Replace(cboDescrizone(0).Text, Chr(39), Chr(96))
        cboDettagliIntervento(1).Text = Replace(cboDettagliIntervento(1).Text, Chr(39), Chr(96))
     
        v_Nomi = Array("KEY", "CODICE_APPARATO", "TIPO_MANUTENZIONE", "DATA_RICHIESTA_MANUTENZIONE", "DATA_EFFETTIVA_MANUTENZIONE", "DESCRIZIONE_MANUTENZIONE", "DETTAGLI_INTERVENTO", "NUMERO_DOCUMENTO")
        
        v_Val = Array(numKey, KeyApparato, txtTipoManutenzione.Text, IIf(oDataRichiestaManutenzione(0).data = "", Null, oDataRichiestaManutenzione(0).data), IIf(oDataEffettivaManutenzione(1).data = "", Null, oDataEffettivaManutenzione(1).data), cboDescrizone(0).Text, cboDettagliIntervento(1).Text, txtNumeroDocumneto)
            
        If SelezionatoManutenzione = False Then
            rsManutenzione.Open "MANUTENZIONE_APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsManutenzione.AddNew v_Nomi, v_Val
        Else
            rsManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsManutenzione.Update v_Nomi, v_Val
        End If
            
        Set rsManutenzione = Nothing
                
    ElseIf tTabellaManutenzione = tpMANUTENZIONEORDINARIA Then
    
        If chkFunzionalità.Value = Checked And chkSicurezza.Value = Unchecked Then
            txtTipoManutenzione.Text = "ORD. FUNZ."
        ElseIf chkSicurezza.Value = Checked And chkFunzionalità.Value = Unchecked Then
            txtTipoManutenzione.Text = "ORD. SICUR."
        ElseIf chkFunzionalità.Value = Checked And chkSicurezza.Value = Checked Then
            txtTipoManutenzione.Text = "ORD. FUN. SIC."
        End If
            
        v_Nomi = Array("KEY", "CODICE_APPARATO", "TIPO_MANUTENZIONE", "DATA_SCADENZA_MANUTENZIONE", "DATA_EFFETTIVA_MANUTENZIONE", "DESCRIZIONE_MANUTENZIONE", "DETTAGLI_INTERVENTO", "NUMERO_DOCUMENTO", "FUNZIONALITA", "SICUREZZA")
        
        v_Val = Array(numKey, KeyApparato, txtTipoManutenzione.Text, IIf(oDataScadenzaManutenzione(1).data = "", Null, oDataScadenzaManutenzione(1).data), IIf(oDataEffettivaManutenzione(1).data = "", Null, oDataEffettivaManutenzione(1).data), cboDescrizone(0).Text, cboDettagliIntervento(1).Text, txtNumeroDocumneto, IIf(chkFunzionalità.Value = Checked, True, False), IIf(chkSicurezza.Value = Checked, True, False))
            
        If SelezionatoManutenzione = False Then
            rsManutenzione.Open "MANUTENZIONE_APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsManutenzione.AddNew v_Nomi, v_Val
        Else
            rsManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsManutenzione.Update v_Nomi, v_Val
        End If
            
        Set rsManutenzione = Nothing
    End If
    
        ' AUTOMATISMO PER LA
        If chkFunzionalità.Value = Checked And chkSicurezza.Value = Checked Then
            Call CalcoloProxRevFun
            Call CalcoloProxRevSic
        ' DATA PROSSIMA REV. FUNZ.
        ElseIf chkFunzionalità.Value = Checked Then
            Call CalcoloProxRevFun
        ' DATA PROSSIMA REV. SIC.
        ElseIf chkSicurezza.Value = Checked Then
            Call CalcoloProxRevSic
        End If
                
    Call Pulisci
        
    If KeyReturnManutenzione > 0 Then
        SelezionatoManutenzione = False
        Unload frmInserisciManutenzione
    Else
        KeyReturnManutenzione = 0
        Unload frmInserisciManutenzione
    End If
        
End Sub

'' AUTOMATISMO PER IL CALCOLO DELLA DATA DELLA PROSSIMA REVISIONE FUNZIONALE
Private Sub CalcoloProxRevFun()

    Set rsDataRevisioneFunzionale = New Recordset
    rsDataRevisioneFunzionale.Open "SELECT * FROM APPARATI WHERE KEY=" & KeyApparato, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        
    If oDataEffettivaManutenzione(1).data > rsDataRevisioneFunzionale("DATA_COLLAUDO") Then
        ProxRevFun = rsDataRevisioneFunzionale("DATA_COLLAUDO")
            
            Select Case rsDataRevisioneFunzionale("FUNZIONALITA")
                Case Is = 0
                ' funzione per sommare la date
                ' d=day, m=month, y=year
                    ProxRevFun = DateAdd("m", 1, oDataEffettivaManutenzione(1).data)
                Case Is = 1
                    ProxRevFun = DateAdd("m", 2, oDataEffettivaManutenzione(1).data)
                Case Is = 2
                    ProxRevFun = DateAdd("m", 3, oDataEffettivaManutenzione(1).data)
                Case Is = 3
                    ProxRevFun = DateAdd("m", 4, oDataEffettivaManutenzione(1).data)
                Case Is = 4
                    ProxRevFun = DateAdd("m", 6, oDataEffettivaManutenzione(1).data)
                Case Is = 5
                    ' calcolo l' aggiunta dell' anno con la somma dei mesi
                    ' in quanto la funzione "year" aggiunge il giorno
                    ProxRevFun = DateAdd("m", 12, oDataEffettivaManutenzione(1).data)
                Case Is = 6
                    ProxRevFun = DateAdd("m", 24, oDataEffettivaManutenzione(1).data)
                Case Is = 7
                    ProxRevFun = DateAdd("m", 36, oDataEffettivaManutenzione(1).data)
            End Select
    
        rsDataRevisioneFunzionale("PROXREVFUN") = ProxRevFun
        rsDataRevisioneFunzionale("LETTO") = False
        rsDataRevisioneFunzionale.Update
        
    End If
        
    Set rsDataRevisioneFunzionale = Nothing

End Sub

'' AUTOMATISMO PER IL CALCOLO DELLA DATA DELLA PROSSIMA REVISIONE SICUREZZA
Private Sub CalcoloProxRevSic()

    Set rsDataRevisioneSicurezza = New Recordset
    rsDataRevisioneSicurezza.Open "SELECT * FROM APPARATI WHERE KEY=" & KeyApparato, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        
    If oDataEffettivaManutenzione(1).data > rsDataRevisioneSicurezza("DATA_COLLAUDO") Then
        ProxRevSic = rsDataRevisioneSicurezza("DATA_COLLAUDO")
            
            Select Case rsDataRevisioneSicurezza("SICUREZZA")
                Case Is = 0
                ' funzione per sommare la date
                ' d=day, m=month, y=year
                    ProxRevSic = DateAdd("m", 1, oDataEffettivaManutenzione(1).data)
                Case Is = 1
                    ProxRevSic = DateAdd("m", 2, oDataEffettivaManutenzione(1).data)
                Case Is = 2
                    ProxRevSic = DateAdd("m", 3, oDataEffettivaManutenzione(1).data)
                Case Is = 3
                    ProxRevSic = DateAdd("m", 4, oDataEffettivaManutenzione(1).data)
                Case Is = 4
                    ProxRevSic = DateAdd("m", 6, oDataEffettivaManutenzione(1).data)
                Case Is = 5
                    ' calcolo l' aggiunta dell' anno con la somma dei mesi
                    ' in quanto la funzione "year" aggiunge il giorno
                    ProxRevSic = DateAdd("m", 12, oDataEffettivaManutenzione(1).data)
                Case Is = 6
                    ProxRevSic = DateAdd("m", 24, oDataEffettivaManutenzione(1).data)
                Case Is = 7
                    ProxRevSic = DateAdd("m", 36, oDataEffettivaManutenzione(1).data)
            End Select
    
        rsDataRevisioneSicurezza("PROXREVSIC") = ProxRevSic
        rsDataRevisioneSicurezza("LETTO") = False
        rsDataRevisioneSicurezza.Update
        
    End If
        
    Set rsDataRevisioneSicurezza = Nothing

End Sub

Private Sub Pulisci()
    oDataRichiestaManutenzione(0).Pulisci
    oDataEffettivaManutenzione(1).Pulisci
    cboDescrizone(0).ListIndex = -1
    cboDettagliIntervento(1).ListIndex = -1
    txtNumeroDocumneto.Text = ""
End Sub

Private Sub cmdStampa_Click()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim KeyProduttore As Integer
    
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(120) AS TIPOLOGIA, " & _
                "       NEW adVarChar(120) AS MODELLO, " & _
                "       NEW adVarChar(120) AS MATRICOLA, " & _
                "       NEW adVarChar(120) AS MOTIVAZIONE_RICHIESTA "
                
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM APPARATI WHERE (KEY=" & KeyApparato & ") ORDER BY KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            Do While Not rsDataset.EOF
                .AddNew
                .Fields("TIPOLOGIA") = rsDataset("TIPO_APPARATO")
                .Fields("MODELLO") = rsDataset("MODELLO")
                .Fields("MATRICOLA") = rsDataset("MATRICOLA")
                If rsDataset("KEY_PRODUTTORE") > 0 Then
                    KeyProduttore = rsDataset("KEY_PRODUTTORE")
                End If
                rsDataset.MoveNext
            Loop
               
                .Fields("MOTIVAZIONE_RICHIESTA") = cboDescrizone(0).Text
            
        End With
    End If
    Set rsDataset = Nothing
    
    ' Ricerca del Produttore
    If KeyProduttore > 0 Then
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM PRODUTTORE_MANUTENTORE WHERE (KEY=" & KeyProduttore & ") ORDER BY KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            With rsMain
                   .AddNew
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblRagioneSociale").Caption = rsDataset("RAGIONE_SOCIALE")
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblCap").Caption = rsDataset("CAP")
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblCitta").Caption = rsDataset("Citta")
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblProvincia").Caption = rsDataset("PROV")
                    rptRichiestaIntervento.Sections("Intestazione").Controls("lblFax").Caption = "Fax: " & rsDataset("FAX")
            End With
        Set rsDataset = Nothing
    Else
     ' Pulisco i campi per evitare di ricaricare i dati in caso non ci sia il produttore
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblRagioneSociale").Caption = ""
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblIndirizzo").Caption = ""
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblCap").Caption = ""
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblCitta").Caption = ""
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblProvincia").Caption = ""
        rptRichiestaIntervento.Sections("Intestazione").Controls("lblFax").Caption = ""
    End If
    
        
    Set rptRichiestaIntervento.DataSource = rsMain
    rptRichiestaIntervento.RightMargin = 0
    rptRichiestaIntervento.LeftMargin = 0
    rptRichiestaIntervento.Sections("Pie").Controls("lblMese").Caption = "lì" & " " & date
    rptRichiestaIntervento.PrintReport True, rptRangeAllPages
        
End Sub


Private Sub Form_Activate()
    Call RicaricaComboBox("DESCRIZIONE_MANUTENZIONE", "NOME", cboDescrizone(0))
    Call RicaricaComboBox("DETTAGLIO_MANUTENZIONE", "NOME", cboDettagliIntervento(1))
End Sub

Private Sub Form_Load()
    Select Case tTabellaManutenzione
        Case tpMANUNTENZIONESTRAORDINARIA
            frmInserisciManutenzione.Caption = "Manutenzione Straordinaria"
            Label1(12).Visible = True
            oDataRichiestaManutenzione(0).Visible = True
            cmdStampa.Visible = True
            txtTipoManutenzione = "STRAORDINARIA"
            lblDescrizioneManutenzione(1).Caption = "Motivazione Richiesta"
            If Selezionato = True Then
                Call CaricaManutenzione
            Else
            'Se non è selezionata nessuna manutenzione
            'si trova nella fase di inserimento e
            'precompila il campo RchiestaManutenzione con la data di sistema
                oDataRichiestaManutenzione(0).txtBox = date
            End If
            
        Case tpMANUTENZIONEORDINARIA
            frmInserisciManutenzione.Caption = "Manutenzione Ordinaria"
            chkFunzionalità.Visible = True
            chkSicurezza.Visible = True
            Label1(5).Visible = True
            oDataScadenzaManutenzione(1).Visible = True
            lblDescrizioneManutenzione(1).Caption = "Descrizione Manutenzione"
            If Selezionato = True Then
                Call CaricaManutenzione
            Else
            'Se non è selezionata nessuna manutenzione
            'si trova nella fase di inserimento e
            'precompila il campo ScadenzaManutenzione con la data di sistema
                oDataScadenzaManutenzione(1).data = date
            End If
            
    End Select
    
    Selezionato = False
    
End Sub

Private Sub CaricaManutenzione()
    Set rsCercaManutenzione = New Recordset
    
    rsCercaManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY =" & KeyReturnManutenzione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        txtTipoManutenzione.Text = rsCercaManutenzione("TIPO_MANUTENZIONE")
        If tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA Then
            oDataRichiestaManutenzione(0).txtBox = rsCercaManutenzione("DATA_RICHIESTA_MANUTENZIONE") & ""
        ElseIf tTabellaManutenzione = tpMANUTENZIONEORDINARIA Then
            oDataScadenzaManutenzione(1).txtBox = rsCercaManutenzione("DATA_SCADENZA_MANUTENZIONE") & ""
            chkFunzionalità.Value = IIf(CBool(rsCercaManutenzione("FUNZIONALITA")), Checked, Unchecked)
            chkSicurezza.Value = IIf(CBool(rsCercaManutenzione("SICUREZZA")), Checked, Unchecked)
        End If
        oDataEffettivaManutenzione(1).txtBox = rsCercaManutenzione("DATA_EFFETTIVA_MANUTENZIONE") & ""
        cboDescrizone(0).Text = rsCercaManutenzione("DESCRIZIONE_MANUTENZIONE")
        cboDettagliIntervento(1).Text = rsCercaManutenzione("DETTAGLI_INTERVENTO")
        txtNumeroDocumneto.Text = rsCercaManutenzione("NUMERO_DOCUMENTO")
        
    Set rsCercaManutenzione = Nothing
End Sub

Private Sub txtNumeroDocumneto_GotFocus()
    txtNumeroDocumneto.BackColor = colArancione
End Sub

Private Sub txtNumeroDocumneto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumeroDocumneto_LostFocus()
    txtNumeroDocumneto.BackColor = vbWhite
End Sub
