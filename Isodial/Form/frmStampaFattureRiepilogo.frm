VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStampaFattureRiepilogo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Riepilogo per Impegnative"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAutorizzazioneBollo 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   5295
      Begin VB.CheckBox chkBollo 
         Caption         =   "Dicitura Bollo Virtuale"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkNumeroAutorizzazione 
         Caption         =   "N° Autorizzazione"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkImportoBollo 
         Caption         =   "Bollo su Fattura"
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Width           =   3015
      End
   End
   Begin VB.Frame fraPeriodo 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
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
         ItemData        =   "frmStampaFattureRiepilogo.frx":0000
         Left            =   4200
         List            =   "frmStampaFattureRiepilogo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
         ItemData        =   "frmStampaFattureRiepilogo.frx":0004
         Left            =   840
         List            =   "frmStampaFattureRiepilogo.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
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
         Index           =   1
         Left            =   3480
         TabIndex        =   7
         Top             =   250
         Width           =   540
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
         Left            =   120
         TabIndex        =   6
         Top             =   250
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   -120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraNumFattura 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5295
      Begin VB.OptionButton Option2 
         Caption         =   "3 fatture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         ToolTipText     =   "Divide i pazienti ASL/Fuori ASL/Fuori Regione"
         Top             =   335
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2 fatture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         ToolTipText     =   "Divide i pazienti ASL/Fuori ASL+Fuori Regione"
         Top             =   110
         Width           =   1095
      End
      Begin VB.TextBox txtNumFattura 
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
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   3
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dividi in"
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
         Left            =   2760
         TabIndex        =   15
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° fattura"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   220
         Width           =   960
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton ControllaFE 
         Caption         =   "C&ontrolla Fatt.EL."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1960
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   910
      End
      Begin VB.CommandButton VisualizzaFE 
         Caption         =   "&Visualizza Fatt.EL."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   980
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   970
      End
      Begin VB.CommandButton fattelettr 
         Caption         =   "&Fattura Elettr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdStampa 
         Cancel          =   -1  'True
         Caption         =   "&Stampa"
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
         Left            =   2920
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdEsci 
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
         TabIndex        =   5
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmStampaFattureRiepilogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' oggetto documentoXML
Dim doc As New DOMDocument60

Dim napoli3 As Boolean
Dim numfat As Integer
Dim txtPercorso As String
Dim ret As Integer
    
Dim FE_Descrizione As String
Dim FE_Quantita As Integer
Dim FE_PrezzoUnit As Single
Dim FE_PrezzoTot As Single
Dim FE_NumLinea As Integer

Private Sub ControllaFE_Click()
    'Apre nel browser il link per il controllo della FE
    'SHOW_SHOWNORMAL = 1
    'SHOW_SHOWMAXIMIZED = 3
    ret = ShellExecute(Me.hWnd, "open", "http://sdi.fatturapa.gov.it/SdI2FatturaPAWeb/AccediAlServizioAction.do?pagina=controlla_fattura", vbNullString, vbNullString, 1)
    If ret < 32 Then MsgBox "Si è verificato un errore aprendo il browser di default", vbCritical, "ATTENZIONE!!!"
End Sub

Private Sub Form_Load()
    Dim rsDataset As New Recordset
    Me.Top = 0
    Me.Left = 10
    Select Case tStampeRiepilogo
        Case tpFATTURA
            Me.Caption = "Stampa Fattura"
            fraNumFattura.Top = fraPulsanti.Top
            fraNumFattura.Left = fraPulsanti.Left
            fraNumFattura.ZOrder
            fraAutorizzazioneBollo.Top = fraNumFattura.Top + fraNumFattura.Height - 135
            fraPulsanti.Top = fraAutorizzazioneBollo.Top + fraAutorizzazioneBollo.Height - 135
            fattelettr.Visible = True
            VisualizzaFE.Visible = True
            ControllaFE.Visible = True
        Case tpXMAZZETTEDISTRETTI
            Me.Caption = "Stampa Mazzette per Distretti"
        Case tpXPAZIENTE
            Me.Caption = "Stampa Riepilogo per Paziente"
        Case tpXTOTALIPERPRESTAZIONE
            Me.Caption = "Stampa Riepilogo per Totali per Prestazione"
        Case TPXMAZZETTEMENSILI
            Me.Caption = "Stampa Riepilogo per Mazzette - Mensili"
            cboMese.AddItem "Giu. - Set. 2010", 0
        Case tpXMAZZETTASINGOLA
            Me.Caption = "Stampa Riepilogo per Mazzetta - Singola"
        Case tpXASLDISTRETTI
            Me.Caption = "Stampa Riepilogo per Asl - Distretti"
        Case tpXTOTALIPERASL
            Me.Caption = "Stampa Riepilogo per Totali per Asl"
    End Select
    Me.Height = fraPulsanti.Top + fraPulsanti.Height + 480
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    
    cboMese.ListIndex = Month(Now) - 1
        
    rsDataset.Open "SELECT CODICE_ASL FROM INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset("CODICE_ASL") = 6 And tStampeRiepilogo = tpFATTURA Then
        napoli3 = True
        Option1.Visible = True
        Option2.Visible = True
        Label1(0).Visible = True
    Else
        napoli3 = False
        Option1.Visible = False
        Option2.Visible = False
        Label1(0).Visible = False
    End If
    rsDataset.Close
    
    
    rsDataset.Open "SELECT CODICE_ASL FROM INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset("CODICE_ASL") = 5 And tStampeRiepilogo = tpFATTURA Then
        chkImportoBollo.Visible = True
    Else
        chkImportoBollo.Visible = False
        chkBollo.Top = 380
        chkNumeroAutorizzazione.Top = 380
    End If
    rsDataset.Close
    Set rsDataset = Nothing
    
    txtPercorso = Environ$("USERPROFILE") & "\Desktop"
    
End Sub

Private Sub StampaPerMazzetteMensili()
    On Error GoTo gestione
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim strCondizione As String

    ' totale dei totali
    Dim importoTotale As Currency
    Dim quantitaTotale As Integer

    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter


    If cboMese.ListIndex >= 6 And cboMese.ListIndex <= 9 And cboAnno.Text = 2010 Then
        MsgBox "Impossibile stampare solo per il mese di " & cboMese.Text & " " & cboAnno.Text & vbCrLf & "Selezionare l'opzione Stampa Giu. - Sett. 2010", vbCritical, "Attenzione"
        Exit Sub
    End If
        
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(10) AS NUMERO_MAZZETTA, " & _
                "       NEW adVarChar(30) AS NOME_DISTRETTO, " & _
                "       NEW adDate AS DATA_RICETTA, " & _
                "       NEW adDate AS DATA_PRENOTAZIONE, " & _
                "       NEW adVarChar(16) AS NUMERO_RICETTA, " & _
                "       NEW adVarChar(35) AS COGNOME, " & _
                "       NEW adVarChar(35) AS NOME, " & _
                "       NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "       NEW adInteger AS QUANTITA," & _
                "       NEW adCurrency AS IMPORTO_TOTALE, " & _
                "       NEW adDate AS DATA_INIZIO, " & _
                "       NEW adDate AS DATA_FINE "

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    If cboMese.ListIndex = 0 Then
        strCondizione = " (MESE=6 OR MESE=7 OR MESE=8 OR MESE=9) AND ANNO=2010 "
    Else
        strCondizione = " MESE=" & cboMese.ListIndex & "AND ANNO=" & cboAnno.Text
    End If
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset

    strSql = "SELECT        RICETTE.*, DISTRETTI.NOME AS DISTRETTINOME, PAZIENTI.COGNOME, PAZIENTI.NOME AS PAZIENTINOME, NOMENCLATORE_TARIFFARIO.CODICE AS NOMENCLATORE_TARIFFARIOCODICE, PRESCRIZIONI.IMPORTO, PRESCRIZIONI.QUANTITA, PRESCRIZIONI.DATA_INIZIO, PRESCRIZIONI.DATA_FINE " & _
             "FROM          ((((RICETTE " & _
             "              INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
             "              INNER JOIN PRESCRIZIONI ON PRESCRIZIONI.CODICE_RICETTA=RICETTE.KEY) " & _
             "              INNER JOIN NOMENCLATORE_TARIFFARIO ON NOMENCLATORE_TARIFFARIO.KEY=PRESCRIZIONI.CODICE_PRESTAZIONE) " & _
             "              LEFT OUTER JOIN DISTRETTI ON DISTRETTI.KEY=PAZIENTI.CODICE_DISTRETTO) " & _
             "WHERE         " & strCondizione & " AND " & _
             "              NOT FLAG=3 " & _
             "ORDER BY      MAZZETTA1, MAZZETTA2"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("NUMERO_MAZZETTA") = rsDataset("MAZZETTA1") & "/" & rsDataset("MAZZETTA2")
            .Fields("NOME_DISTRETTO") = rsDataset("DISTRETTINOME") & ""
            .Fields("DATA_RICETTA") = rsDataset("DATA_RICETTA")
            .Fields("DATA_PRENOTAZIONE") = rsDataset("DATA_PRENOTAZIONE")
            .Fields("NUMERO_RICETTA") = rsDataset("NUMERO_RICETTA")
            .Fields("COGNOME") = rsDataset("COGNOME")
            .Fields("NOME") = rsDataset("PAZIENTINOME")
            .Fields("CODICE_PRESTAZIONE") = rsDataset("NOMENCLATORE_TARIFFARIOCODICE")
            .Fields("QUANTITA") = rsDataset("QUANTITA")
            quantitaTotale = quantitaTotale + .Fields("QUANTITA")
            .Fields("IMPORTO_TOTALE") = rsDataset("IMPORTO") * rsDataset("QUANTITA")
            importoTotale = importoTotale + .Fields("IMPORTO_TOTALE")
            .Fields("DATA_INIZIO") = rsDataset("DATA_INIZIO")
            .Fields("DATA_FINE") = rsDataset("DATA_FINE")
            .Update
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NUMERO_MAZZETTA") = ""
        .Fields("DATA_RICETTA") = Null
        .Fields("DATA_PRENOTAZIONE") = Null
        .Fields("NUMERO_RICETTA") = ""
        .Fields("COGNOME") = ""
        .Fields("NOME") = ""
        .Fields("CODICE_PRESTAZIONE") = "TOTALI"
        .Fields("QUANTITA") = quantitaTotale
        .Fields("IMPORTO_TOTALE") = Round(importoTotale, 2)
        .Fields("DATA_INIZIO") = Null
        .Fields("DATA_FINE") = Null
        .Update
    End With
    
    
    If rsMain.RecordCount = 2 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        rptRiepilogoPerMazzetteMensili.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & IIf(cboMese.ListIndex = 0, "", cboAnno.Text)
        Set rptRiepilogoPerMazzetteMensili.DataSource = rsMain
        rptRiepilogoPerMazzetteMensili.Orientation = rptOrientLandscape
        rptRiepilogoPerMazzetteMensili.LeftMargin = rptRiepilogoPerMazzetteMensili.LeftMargin / 4
        rptRiepilogoPerMazzetteMensili.RightMargin = rptRiepilogoPerMazzetteMensili.RightMargin / 4
        rptRiepilogoPerMazzetteMensili.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaPerPaziente()
    On Error GoTo gestione
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rsAppo2 As Recordset
    

    ' totale dei totali
    Dim importoTotale As Currency
    Dim quantitaTotale As Integer
    Dim importoTicket As Integer
    Dim importoQuotaNazionale As Integer
    Dim importoQuotaAggiuntiva As Integer
    
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim coeffTicket As Integer
    Dim coeffQuotaNazionale As Integer
    Dim coeffQuota As Single
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
        
    strShape = "SHAPE APPEND " & _
                "               NEW adVarChar(45) AS PAZIENTE, " & _
                "               NEW adVarChar(16) AS NUMERO_RICETTA, " & _
                "               NEW adVarChar(10) AS ESENZIONE, " & _
                "               NEW adCurrency AS TICKET, " & _
                "               NEW adCurrency AS QUOTA_AGGIUNTIVA, " & _
                "               NEW adCurrency AS QUOTA_NAZIONALE, " & _
                "               NEW adDate AS DATA_RICETTA, " & _
                "               NEW adDate AS DATA_PRENOTAZIONE, " & _
                "               NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "               NEW adInteger AS QUANTITA," & _
                "               NEW adCurrency AS IMPORTO_TOTALE, " & _
                "               NEW adDate AS DATA_INIZIO, " & _
                "               NEW adDate AS DATA_FINE "

        
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    Set rsAppo2 = New Recordset
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close


    strSql = "SELECT    DISTINCT PAZIENTI.KEY, COGNOME, NOME, DATA_NASCITA " & _
             "FROM      (RICETTE " & _
             "          INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
             "WHERE     MESE=" & cboMese.ListIndex + 1 & " AND " & _
             "          ANNO=" & cboAnno.Text & " AND " & _
             "          NOT FLAG=3 " & _
             "ORDER BY  COGNOME"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        strSql = "SELECT    NUMERO_RICETTA, TIPOLOGIE_ESENZIONE.CODICE AS TIPOLOGIE_ESENZIONECODICE, " & _
                 "          DATA_RICETTA, DATA_PRENOTAZIONE, NOMENCLATORE_TARIFFARIO.CODICE AS NOMENCLATORE_TARIFFARIOCODICE, " & _
                 "          QUANTITA, PRESCRIZIONI.IMPORTO, DATA_INIZIO, DATA_FINE, CODICE_PRESTAZIONE, RICETTE.KEY  " & _
                 "FROM      (((RICETTE " & _
                 "          INNER JOIN PRESCRIZIONI ON PRESCRIZIONI.CODICE_RICETTA=RICETTE.KEY) " & _
                 "          INNER JOIN NOMENCLATORE_TARIFFARIO ON NOMENCLATORE_TARIFFARIO.KEY=PRESCRIZIONI.CODICE_PRESTAZIONE) " & _
                 "          INNER JOIN TIPOLOGIE_ESENZIONE ON TIPOLOGIE_ESENZIONE.KEY=RICETTE.CODICE_ESENZIONE) " & _
                 "WHERE     MESE=" & cboMese.ListIndex + 1 & " AND " & _
                 "          ANNO=" & cboAnno.Text & " AND " & _
                 "          CODICE_PAZIENTE=" & rsDataset("KEY") & " AND " & _
                 "          NOT FLAG=3"
        rsAppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsAppo.EOF
            With rsMain
                .AddNew
                .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME") & " - " & rsDataset("DATA_NASCITA")
                .Fields("NUMERO_RICETTA") = rsAppo("NUMERO_RICETTA")
                .Fields("ESENZIONE") = rsAppo("TIPOLOGIE_ESENZIONECODICE")
                .Fields("DATA_RICETTA") = rsAppo("DATA_RICETTA")
                .Fields("DATA_PRENOTAZIONE") = rsAppo("DATA_PRENOTAZIONE")
                .Fields("CODICE_PRESTAZIONE") = rsAppo("NOMENCLATORE_TARIFFARIOCODICE")
                .Fields("QUANTITA") = rsAppo("QUANTITA")
                quantitaTotale = quantitaTotale + .Fields("QUANTITA")
                .Fields("IMPORTO_TOTALE") = rsAppo("QUANTITA") * rsAppo("IMPORTO")
                importoTotale = importoTotale + .Fields("IMPORTO_TOTALE")
                .Fields("DATA_INIZIO") = rsAppo("DATA_INIZIO")
                .Fields("DATA_FINE") = rsAppo("DATA_FINE")
                
                
                rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsAppo("CODICE_PRESTAZIONE") & " AND R.KEY=" & rsAppo("KEY") & " AND (NOT FLAG=3) AND (CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                coeffTicket = rsAppo2("TOTALE_R")
                coeffQuotaNazionale = rsAppo2("TOTALE_R")
                rsAppo2.Close
                
                rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsAppo("CODICE_PRESTAZIONE") & " AND R.KEY=" & rsAppo("KEY") & " AND (NOT FLAG=3) AND CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                coeffQuota = rsAppo2("TOTALE_R")
                rsAppo2.Close
                rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsAppo("CODICE_PRESTAZIONE") & " AND R.KEY=" & rsAppo("KEY") & " AND (NOT FLAG=3) AND (NOT CODICE_ESENZIONE=-1) AND T.ESENZIONE_QUOTA=FALSE  AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                coeffQuota = coeffQuota + rsAppo2("TOTALE_R") / 2
                rsAppo2.Close
                
                .Fields("QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
                importoQuotaAggiuntiva = importoQuotaAggiuntiva + .Fields("QUOTA_AGGIUNTIVA")
                
                .Fields("QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
                importoQuotaNazionale = importoQuotaNazionale + .Fields("QUOTA_NAZIONALE")
                
                .Fields("TICKET") = ticket * coeffTicket
                importoTicket = importoTicket + .Fields("TICKET")
                
                .Update
            End With
            coeffQuota = 0
            rsAppo.MoveNext
        Loop
        rsAppo.Close
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    If quantitaTotale = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
        Exit Sub
    End If
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .Fields("PAZIENTE") = ""
        .Fields("NUMERO_RICETTA") = ""
        .Fields("DATA_RICETTA") = Null
        .Fields("DATA_PRENOTAZIONE") = Null
        .Fields("CODICE_PRESTAZIONE") = "TOTALI"
        .Fields("QUANTITA") = quantitaTotale
        .Fields("IMPORTO_TOTALE") = Round(importoTotale, 2)
        .Fields("DATA_INIZIO") = Null
        .Fields("DATA_FINE") = Null
        .Fields("TICKET") = importoTicket
        .Fields("QUOTA_AGGIUNTIVA") = importoQuotaAggiuntiva
        .Fields("QUOTA_NAZIONALE") = importoQuotaNazionale
        .Update
    End With
    
    rptRiepilogoPerPaziente.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
    Set rptRiepilogoPerPaziente.DataSource = rsMain
    rptRiepilogoPerPaziente.Orientation = rptOrientLandscape
    rptRiepilogoPerPaziente.PrintReport False, rptRangeAllPages

    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaPerTotaliPerPrestazione()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset

    ' totale dei totali
    Dim importoTotale As Currency
    Dim quantitaTotale As Integer
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "       NEW adVarChar(100) AS NOME_PRESTAZIONE, " & _
                "       NEW adInteger AS QUANTITA, " & _
                "       NEW adCurrency AS IMPORTO_UNITARIO, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE "
                  
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    

    Set rsDataset = New Recordset
    Set rsAppo = New Recordset

    rsDataset.Open "SELECT DISTINCT PR.CODICE_PRESTAZIONE,PR.IMPORTO, CODICE, NOME FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN NOMENCLATORE_TARIFFARIO N ON N.KEY=PR.CODICE_PRESTAZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            .Fields("QUANTITA") = rsAppo("TOTALEQ")
            rsAppo.Close
            .Fields("CODICE_PRESTAZIONE") = rsDataset("CODICE")
            .Fields("NOME_PRESTAZIONE") = rsDataset("NOME")
            quantitaTotale = quantitaTotale + .Fields("QUANTITA")
            .Fields("IMPORTO_UNITARIO") = rsDataset("IMPORTO")
            .Fields("IMPORTO_TOTALE") = .Fields("QUANTITA") * .Fields("IMPORTO_UNITARIO")
            importoTotale = importoTotale + .Fields("IMPORTO_TOTALE")
            .Update
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("CODICE_PRESTAZIONE") = "TOTALI"
        .Fields("NOME_PRESTAZIONE") = ""
        .Fields("QUANTITA") = quantitaTotale
        .Fields("IMPORTO_UNITARIO") = 0
        .Fields("IMPORTO_TOTALE") = Round(importoTotale, 2)
    End With
    
    
    If rsMain.RecordCount = 2 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        rptRiepilogoPerTotaliPerPrestazione.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptRiepilogoPerTotaliPerPrestazione.DataSource = rsMain
        rptRiepilogoPerTotaliPerPrestazione.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaFattura()
    On Error GoTo gestione
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim totaleRicette As Integer
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
    ' totale dei totali
    Dim totaleAsl As Integer
    Dim totaleRegione As Integer
    Dim totaleFuoriRegione As Integer
    Dim importoTotale As Currency
    Dim importoTotaleScontato As Currency
    Dim importoTotaleTicket As Currency
    Dim importoTotaleQuotaAggiuntiva As Currency
    Dim importoTotaleQuotaNazionale As Currency
    Dim importoTotaleNetto As Currency
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    importoTotale = 0
    importoTotaleNetto = 0
    importoTotaleScontato = 0
    importoTotaleTicket = 0
    importoTotaleQuotaAggiuntiva = 0
    importoTotaleQuotaNazionale = 0
    
        
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "       NEW adInteger AS TOTALE_ASL, " & _
                "       NEW adInteger AS TOTALE_REGIONE, " & _
                "       NEW adInteger AS TOTALE_FUORI_REGIONE, " & _
                "       NEW adCurrency AS IMPORTO_UNITARIO, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE, " & _
                "       NEW adCurrency AS IMPORTO_SCONTATO, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE_SCONTATO, " & _
                "       NEW adCurrency AS TOTALE_TICKET, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE, " & _
                "       NEW adCurrency AS IMPORTO_NETTO"

        
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    rsDataset.Open "SELECT DISTINCT PR.CODICE_PRESTAZIONE, CODICE FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN NOMENCLATORE_TARIFFARIO N ON N.KEY=PR.CODICE_PRESTAZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("CODICE_PRESTAZIONE") = rsDataset("CODICE")
            rsAppo.Open "SELECT  SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_ASL") = 0
            Else
                .Fields("TOTALE_ASL") = rsAppo("TOTALEQ")
            End If
            rsAppo.Close
            totaleAsl = totaleAsl + .Fields("TOTALE_ASL")
        
            rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND P.CODICE_ASL IN (SELECT KEY FROM ASL WHERE CODICE_REGIONE=16) AND NOT P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_REGIONE") = 0
            Else
                .Fields("TOTALE_REGIONE") = rsAppo("TOTALEQ")
            End If
            rsAppo.Close
            totaleRegione = totaleRegione + .Fields("TOTALE_REGIONE")
            
            rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT P.CODICE_ASL IN (SELECT KEY FROM ASL WHERE CODICE_REGIONE=16) AND NOT P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_FUORI_REGIONE") = 0
            Else
                .Fields("TOTALE_FUORI_REGIONE") = rsAppo("TOTALEQ")
            End If
            rsAppo.Close
            totaleFuoriRegione = totaleFuoriRegione + .Fields("TOTALE_FUORI_REGIONE")
        
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND T.ESENZIONE_QUOTA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            totaleRicette = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND (CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND (NOT CODICE_ESENZIONE=-1) AND T.ESENZIONE_QUOTA=FALSE  AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            rsAppo.Open "SELECT DISTINCT IMPORTO, IMPORTO_SCONTATO FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            .Fields("IMPORTO_UNITARIO") = rsAppo("IMPORTO")
            .Fields("IMPORTO_TOTALE") = .Fields("IMPORTO_UNITARIO") * (.Fields("TOTALE_ASL") + .Fields("TOTALE_REGIONE") + .Fields("TOTALE_FUORI_REGIONE"))
            importoTotale = importoTotale + .Fields("IMPORTO_TOTALE")
            .Fields("IMPORTO_SCONTATO") = rsAppo("IMPORTO_SCONTATO")
            .Fields("IMPORTO_TOTALE_SCONTATO") = .Fields("IMPORTO_SCONTATO") * (.Fields("TOTALE_ASL") + .Fields("TOTALE_REGIONE") + .Fields("TOTALE_FUORI_REGIONE"))
            
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("IMPORTO_NETTO") = .Fields("IMPORTO_TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            
            importoTotaleQuotaAggiuntiva = importoTotaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
            importoTotaleQuotaNazionale = importoTotaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
            importoTotaleTicket = importoTotaleTicket + .Fields("TOTALE_TICKET")
            importoTotaleNetto = importoTotaleNetto + .Fields("IMPORTO_NETTO")
            importoTotaleScontato = importoTotaleScontato + .Fields("IMPORTO_TOTALE_SCONTATO")
            totaleRicette = 0
            rsAppo.Close
            .Update
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("CODICE_PRESTAZIONE") = "TOTALI"
        .Fields("TOTALE_ASL") = totaleAsl
        .Fields("TOTALE_REGIONE") = totaleRegione
        .Fields("TOTALE_FUORI_REGIONE") = totaleFuoriRegione
        .Fields("IMPORTO_UNITARIO") = 0
        .Fields("IMPORTO_TOTALE") = Round(importoTotale, 2)
        .Fields("IMPORTO_SCONTATO") = 0
        .Fields("IMPORTO_TOTALE_SCONTATO") = Round(importoTotaleScontato, 2)
        .Fields("TOTALE_TICKET") = Round(importoTotaleTicket, 2)
        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = Round(importoTotaleQuotaAggiuntiva, 2)
        .Fields("TOTALE_QUOTA_NAZIONALE") = Round(importoTotaleQuotaNazionale, 2)
        .Fields("IMPORTO_NETTO") = Round(importoTotaleNetto, 2)
    End With
  
    If rsMain.RecordCount = 2 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.NOME AS ASLNOME, COMUNI.NOME AS COMUNINOME " & _
                 "FROM      ((INTESTAZIONE_FATTURA " & _
                 "          INNER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
                 "          INNER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE) "
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            rptFattura.Sections("intestazione").Controls("lblAsl").Caption = "ASL " & rsDataset("ASLNOME")
            rptFattura.Sections("intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
            rptFattura.Sections("intestazione").Controls("lblCap").Caption = rsDataset("CAP")
            rptFattura.Sections("intestazione").Controls("lblProvincia").Caption = rsDataset("COMUNINOME") & " (" & rsDataset("PROV") & ")"
            rptFattura.Sections("intestazione").Controls("lblIva").Caption = rsDataset("P_IVA")
            rptFattura.Sections("pie").Controls("lblDicitura").Caption = rsDataset("DICITURA")
            rptFattura.Sections("pie").Controls("lblIntestatario").Caption = rsDataset("INTESTATARIO_CC")
            rptFattura.Sections("pie").Controls("lblIban").Caption = rsDataset("IBAN")
            
            If chkNumeroAutorizzazione.Value = Checked Then
                rptFattura.Sections("pie").Controls("lblAutorizzazione").Caption = "Autorizzazione N° " & rsDataset("NUMERO_AUTORIZZAZIONE")
            Else
                rptFattura.Sections("pie").Controls("lblAutorizzazione").Caption = ""
            End If
            
            If chkBollo.Value = Checked Then
                rptFattura.Sections("pie").Controls("lblBollo").Caption = "Obbligo bollo assolto in maniera virtuale"
            Else
                rptFattura.Sections("pie").Controls("lblBollo").Caption = ""
            End If
            
        End If
        rsDataset.Close
        rptFattura.Sections("intestazione").Controls("lblNumeroFattura").Caption = txtNumFattura.Text & " / " & cboAnno.Text
        rptFattura.Sections("intestazione").Controls("lblData").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptFattura.Sections("intestazione").Controls("lblNomeAsl").Caption = "Emodialisi " & GetNome(structIntestazione.sCodiceAsl, "ASL")
        rptFattura.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        rptFattura.Sections("intestazione").Controls("lblTicket").Caption = "Ticket    " & Format(ticket, "###.00") & " €"
        Set rptFattura.DataSource = rsMain
        rptFattura.RightMargin = 0
        rptFattura.LeftMargin = 0
        rptFattura.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaPerAslDistretto()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rsAppo2 As Recordset
    Dim rsRicette As Recordset
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
    ' totali di una singola asl
    Dim totaleLordo As Currency
    Dim totaleScontato As Currency
    Dim totaleRicette As Integer
    Dim totalePrestazioni As Integer
    Dim totaleMazzette As Integer
    Dim totaleNetto As Currency
    Dim totaleTicket As Currency
    Dim totaleQuotaAggiuntiva As Currency
    Dim totaleQuotaNazionale As Currency
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter

    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS NOME_ASL, " & _
                "       NEW adCurrency AS TOTALE_LORDO_T, " & _
                "       NEW adCurrency AS SCONTO_T, " & _
                "       NEW adInteger AS MAZZETTE_T, " & _
                "       NEW adInteger AS RICETTE_T, " & _
                "       NEW adInteger AS PRESTAZIONI_T, " & _
                "       NEW adCurrency AS TOTALE_SCONTATO_T, " & _
                "       NEW adCurrency AS TOTALE_NETTO_T, " & _
                "       NEW adCurrency AS TOTALE_TICKET_T, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA_T, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE_T, " & _
                "       NEW adInteger AS CODICE_ASL, " & _
                "       (( SHAPE APPEND NEW adInteger AS CODICE_ASL, " & _
                "           NEW adVarChar(30) AS NOME_DISTRETTO, " & _
                "           NEW adInteger AS MAZZETTE, " & _
                "           NEW adInteger AS RICETTE, " & _
                "           NEW adCurrency AS TOTALE_LORDO, " & _
                "           NEW adCurrency AS SCONTO, " & _
                "           NEW adInteger AS PRESTAZIONI, " & _
                "           NEW adCurrency AS TOTALE_SCONTATO, " & _
                "           NEW adCurrency AS TOTALE_NETTO, " & _
                "           NEW adCurrency AS TOTALE_TICKET, " & _
                "           NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "           NEW adCurrency AS TOTALE_QUOTA_NAZIONALE ) RELATE CODICE_ASL TO CODICE_ASL ) AS Res1 "
                
                       
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    Set rsRicette = New Recordset
    Set rsAppo2 = New Recordset
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    With rsMain
        rsDataset.Open "SELECT DISTINCT A.KEY, A.NOME, ABS(A.KEY-" & structIntestazione.sCodiceAsl & ") AS INDICE FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) left outer JOIN ASL A ON A.KEY=P.CODICE_ASL) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataset.EOF
            .AddNew
            .Fields("NOME_ASL") = "ASL " & rsDataset("NOME")
            .Fields("CODICE_ASL") = IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY"))
            rsAppo.Open "SELECT * FROM DISTRETTI WHERE CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsAppo.EOF
                rsRicette.Open "SELECT COUNT(R.KEY) AS TOTALE_RICETTE FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=" & rsAppo("KEY") & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If rsRicette("TOTALE_RICETTE") = 0 Then
                    rsRicette.Close
                Else
                    Set rsFiglio = .Fields("Res1").Value
                    With rsFiglio
                        .AddNew
                        .Fields("CODICE_ASL") = rsDataset("KEY")
                        .Fields("NOME_DISTRETTO") = "      Distretto: " & rsAppo("NOME")
                        .Fields("RICETTE") = rsRicette("TOTALE_RICETTE")
                        .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)
                        rsRicette.Close
                        
                        rsAppo2.Open "SELECT SUM(QUANTITA) AS TOTALE_Q,SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=" & rsAppo("KEY") & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        .Fields("PRESTAZIONI") = rsAppo2("TOTALE_Q")
                        .Fields("TOTALE_SCONTATO") = rsAppo2("TOTALE_SCONTATO")
                        .Fields("TOTALE_LORDO") = rsAppo2("TOTALE_LORDO")
                        rsAppo2.Close
                        
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05') AND CODICE_DISTRETTO=" & rsAppo("KEY"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffTicket = rsAppo2("TOTALE_R")
                        coeffQuotaNazionale = rsAppo2("TOTALE_R")
                        rsAppo2.Close
                        
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1 AND CODICE_DISTRETTO=" & rsAppo("KEY"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffQuota = rsAppo2("TOTALE_R")
                        rsAppo2.Close
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND CODICE_DISTRETTO=" & rsAppo("KEY") & " AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffQuota = coeffQuota + rsAppo2("TOTALE_R") / 2
                        rsAppo2.Close
                        
                        .Fields("TOTALE_TICKET") = ticket * coeffTicket
                        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
                        .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
                        .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
                        .Fields("SCONTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_SCONTATO")
                        
                        totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
                        totaleScontato = totaleScontato + .Fields("TOTALE_SCONTATO")
                        totaleRicette = totaleRicette + .Fields("RICETTE")
                        totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
                        totaleMazzette = totaleMazzette + .Fields("MAZZETTE")
                        totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
                        totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
                        totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
                        totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
                        .Update
                    End With
                End If
                rsAppo.MoveNext
            Loop
            rsAppo.Close
            
            ' ricette senza distretto
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_RICETTE FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=-1 AND CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")) & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rsAppo("TOTALE_RICETTE") <> 0 Then
                Do While Not rsAppo.EOF
                    Set rsFiglio = .Fields("Res1").Value
                    With rsFiglio
                        .AddNew
                        .Fields("CODICE_ASL") = IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY"))
                        .Fields("NOME_DISTRETTO") = "      Distretto: -- "
                        .Fields("RICETTE") = rsAppo("TOTALE_RICETTE")
                        .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)
                        
                        rsAppo2.Open "SELECT SUM(QUANTITA) AS TOTALE_Q,SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=-1 AND CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")) & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        .Fields("PRESTAZIONI") = rsAppo2("TOTALE_Q")
                        .Fields("TOTALE_SCONTATO") = rsAppo2("TOTALE_SCONTATO")
                        .Fields("TOTALE_LORDO") = rsAppo2("TOTALE_LORDO")
                        rsAppo2.Close
                        
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05') AND CODICE_DISTRETTO=-1 AND CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffTicket = rsAppo2("TOTALE_R")
                        coeffQuotaNazionale = rsAppo2("TOTALE_R")
                        rsAppo2.Close
                        
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1 AND CODICE_DISTRETTO=-1 AND CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffQuota = rsAppo2("TOTALE_R")
                        rsAppo2.Close
                        rsAppo2.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND CODICE_DISTRETTO=-1 AND CODICE_ASL=" & IIf(IsNull(rsDataset("KEY")), -1, rsDataset("KEY")) & " AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        coeffQuota = coeffQuota + rsAppo2("TOTALE_R") / 2
                        rsAppo2.Close
                        
                        .Fields("TOTALE_TICKET") = ticket * coeffTicket
                        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
                        .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
                        .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
                        .Fields("SCONTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_SCONTATO")
                       
                        totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
                        totaleScontato = totaleScontato + .Fields("TOTALE_SCONTATO")
                        totaleRicette = totaleRicette + .Fields("RICETTE")
                        totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
                        totaleMazzette = totaleMazzette + .Fields("MAZZETTE")
                        totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
                        totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
                        totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
                        totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
                        .Update
                    End With
                    rsAppo.MoveNext
                Loop
            End If
            
            rsAppo.Close
            .Fields("TOTALE_LORDO_T") = Round(totaleLordo, 2)
            .Fields("TOTALE_SCONTATO_T") = Round(totaleScontato, 2)
            .Fields("SCONTO_T") = Round(totaleLordo, 2) - Round(totaleScontato, 2)
            .Fields("MAZZETTE_T") = totaleMazzette
            .Fields("RICETTE_T") = totaleRicette
            .Fields("PRESTAZIONI_T") = totalePrestazioni
            .Fields("TOTALE_NETTO_T") = Round(totaleNetto, 2)
            .Fields("TOTALE_TICKET_T") = Round(totaleTicket, 2)
            .Fields("TOTALE_QUOTA_AGGIUNTIVA_T") = Round(totaleQuotaAggiuntiva, 2)
            .Fields("TOTALE_QUOTA_NAZIONALE_T") = Round(totaleQuotaNazionale, 2)
            .Update
            totaleLordo = 0
            totaleNetto = 0
            totaleScontato = 0
            totaleMazzette = 0
            totaleRicette = 0
            totalePrestazioni = 0
            totaleTicket = 0
            totaleQuotaAggiuntiva = 0
            totaleQuotaNazionale = 0
            rsDataset.MoveNext
        Loop
        rsDataset.Close
    End With
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
    rptRiepilogoPerAslDistretto.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
    rptRiepilogoPerAslDistretto.Sections("intestazione").Controls("lblTicket").Caption = "Ticket    " & Format(ticket, "###.00") & " €"
    Set rptRiepilogoPerAslDistretto.DataSource = rsMain
    rptRiepilogoPerAslDistretto.RightMargin = 0
    rptRiepilogoPerAslDistretto.LeftMargin = 0
    rptRiepilogoPerAslDistretto.PrintReport False, rptRangeAllPages
    End If
    
    Set rsRicette = Nothing
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaTotaliPerAsl()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    
    Dim nomeAsl As String
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(40) AS NOME, " & _
                "       NEW adCurrency AS TOTALE_LORDO, " & _
                "       NEW adInteger AS MAZZETTE, " & _
                "       NEW adInteger AS RICETTE, " & _
                "       NEW adInteger AS PRESTAZIONI, " & _
                "       NEW adCurrency AS TOTALE_TICKET, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE, " & _
                "       NEW adCurrency AS TOTALE_SCONTATO, " & _
                "       NEW adCurrency AS TOTALE_NETTO "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT CODICE_ASL, NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    With rsMain
        rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If rsDataset("TOTALE_R") = 0 Then
            rsDataset.Close
            MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
            Exit Sub
        End If
        .AddNew
        .Fields("NOME") = "TOTALE GENERALE"
        .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
        .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
        .Fields("RICETTE") = rsDataset("TOTALE_R")
        .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
        .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)
        
        rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        coeffTicket = rsAppo("TOTALE_R")
        coeffQuotaNazionale = rsAppo("TOTALE_R")
        rsAppo.Close
        
        rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        coeffQuota = rsAppo("TOTALE_R")
        rsAppo.Close
        rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
        rsAppo.Close
        
        .Fields("TOTALE_TICKET") = ticket * coeffTicket
        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
        .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
        .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - -.Fields("TOTALE_QUOTA_NAZIONALE")
        .Update
        rsDataset.Close
        
        rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        .AddNew
        .Fields("NOME") = "TOTALE ASL " & nomeAsl
        If rsDataset("TOTALE_R") <> 0 Then
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)

            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
        
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - -.Fields("TOTALE_QUOTA_NAZIONALE")
        Else
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTATO") = 0
            .Fields("RICETTE") = 0
            .Fields("PRESTAZIONI") = 0
            .Fields("MAZZETTE") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
        End If
        .Update
        rsDataset.Close
        
        rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        .AddNew
        .Fields("NOME") = "TOTALE FUORI ASL " & nomeAsl
        If rsDataset("TOTALE_R") <> 0 Then
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
        
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - -.Fields("TOTALE_QUOTA_NAZIONALE")
        Else
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTATO") = 0
            .Fields("RICETTE") = 0
            .Fields("PRESTAZIONI") = 0
            .Fields("MAZZETTE") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
        End If
        .Update
        rsDataset.Close
        
        rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        .AddNew
        .Fields("NOME") = "TOTALE FUORI REGIONE CAMPANIA"
        If rsDataset("TOTALE_R") <> 0 Then
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("MAZZETTE") = Int(.Fields("RICETTE") / 50 + 1)
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND NOT CODICE_REGIONE=16 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND NOT CODICE_REGIONE=16 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3 AND NOT CODICE_REGIONE=16 AND NOT R.CODICE_ESENZIONE=-1 AND T.ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
        
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - -.Fields("TOTALE_QUOTA_NAZIONALE")
        Else
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTATO") = 0
            .Fields("RICETTE") = 0
            .Fields("PRESTAZIONI") = 0
            .Fields("MAZZETTE") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
        End If
        .Update
        rsDataset.Close
    End With
    
    rptRiepilogoPerTotaliPerAsl.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
    Set rptRiepilogoPerTotaliPerAsl.DataSource = rsMain
    rptRiepilogoPerTotaliPerAsl.RightMargin = 0
    rptRiepilogoPerTotaliPerAsl.LeftMargin = 0
    rptRiepilogoPerTotaliPerAsl.PrintReport False, rptRangeAllPages
    
    Set rsDataset = Nothing

gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaRiepilogoImpegnative()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    
    ' totali
    Dim totaleRicette As Integer
    Dim totalePrestazioni As Integer
    Dim totaleLordo As Currency
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(25) AS NOME_ASL, " & _
                "       NEW adVarChar(4) AS NOME_DISTRETTO, " & _
                "       NEW adInteger AS PROGRESSIVO_RICETTA, " & _
                "       NEW adVarChar(15) AS NUMERO_RICETTA, " & _
                "       NEW adInteger AS MAZZETTA, " & _
                "       NEW adVarChar(30) AS COGNOME, " & _
                "       NEW adVarChar(30) AS NOME, " & _
                "       NEW adInteger AS NUMERO_PRESTAZIONI, " & _
                "       NEW adVarChar(15) AS CODICE_PRESTAZIONE, " & _
                "       NEW adDate AS DATA_INIZIO, " & _
                "       NEW adDate AS DATA_FINE, " & _
                "       NEW adCurrency AS TOTALE_LORDO "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    With rsMain
        rsDataset.Open "SELECT R.KEY, COGNOME, P.NOME, A.NOME, D.NOME, MAZZETTA1, PROGRESSIVO_RICETTA, NUMERO_RICETTA FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) left outer JOIN ASL A ON A.KEY=P.CODICE_ASL) LEFT OUTER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 ORDER BY PROGRESSIVO_RICETTA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataset.EOF
            rsAppo.Open "SELECT QUANTITA, CODICE, DATA_INIZIO, DATA_FINE, PR.IMPORTO FROM (PRESCRIZIONI PR INNER JOIN NOMENCLATORE_TARIFFARIO N ON N.KEY=PR.CODICE_PRESTAZIONE) WHERE CODICE_RICETTA=" & rsDataset("KEY"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsAppo.EOF
                .AddNew
                .Fields("NOME_ASL") = rsDataset("A.NOME") & ""
                .Fields("NOME_DISTRETTO") = rsDataset("D.NOME") & ""
                .Fields("PROGRESSIVO_RICETTA") = rsDataset("PROGRESSIVO_RICETTA")
                .Fields("NUMERO_RICETTA") = rsDataset("NUMERO_RICETTA")
                .Fields("MAZZETTA") = rsDataset("MAZZETTA1")
                .Fields("COGNOME") = rsDataset("COGNOME")
                .Fields("NOME") = rsDataset("P.NOME")
                .Fields("NUMERO_PRESTAZIONI") = rsAppo("QUANTITA")
                .Fields("CODICE_PRESTAZIONE") = rsAppo("CODICE")
                .Fields("DATA_INIZIO") = rsAppo("DATA_INIZIO")
                .Fields("DATA_FINE") = rsAppo("DATA_FINE")
                .Fields("TOTALE_LORDO") = rsAppo("QUANTITA") * rsAppo("IMPORTO")
                totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
                totalePrestazioni = totalePrestazioni + .Fields("NUMERO_PRESTAZIONI")
                .Update
                rsAppo.MoveNext
            Loop
            rsAppo.Close
            rsDataset.MoveNext
        Loop
        totaleRicette = rsDataset.RecordCount
        rsDataset.Close
    End With
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        rptRiepilogoImpegnative.Sections("intestazione").Controls("lblTotale").Caption = Format(totaleLordo, "###,###.00")
        rptRiepilogoImpegnative.Sections("intestazione").Controls("lblNumeroImpegnative").Caption = totaleRicette
        rptRiepilogoImpegnative.Sections("intestazione").Controls("lblNumeroPrestazioni").Caption = totalePrestazioni
        rptRiepilogoImpegnative.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptRiepilogoImpegnative.DataSource = rsMain
        rptRiepilogoImpegnative.LeftMargin = 1000
        rptRiepilogoImpegnative.RightMargin = 0
        rptRiepilogoImpegnative.Orientation = rptOrientLandscape
        rptRiepilogoImpegnative.PrintReport False, rptRangeAllPages
    End If
    
    Set rsAppo = Nothing
    Set rsDataset = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaPerAslNapoli2Nord()
    On Error GoTo gestione
    Dim SQLString As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim nomeAsl As String
    
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
    Dim totalePrestazioni As Integer
    Dim intTotaleRicette As Integer
    Dim totaleLordo As Currency
    Dim totaleSconto As Currency
    Dim totaleTicket As Currency
    Dim totaleQuotaAggiuntiva As Currency
    Dim totaleQuotaNazionale As Currency
    Dim totaleNetto As Currency
    Dim totaleNettoScontato As Currency
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter

    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(30) AS NOME_ASL, " & _
                "       NEW adInteger AS RICETTE, " & _
                "       NEW adInteger AS PRESTAZIONI, " & _
                "       NEW adCurrency AS TOTALE_LORDO, " & _
                "       NEW adCurrency AS TOTALE_TICKET, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE, " & _
                "       NEW adCurrency AS TOTALE_NETTO, " & _
                "       NEW adCurrency AS TOTALE_SCONTO, " & _
                "       NEW adCurrency AS TOTALE_NETTO_SCONTATO "

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT CODICE_ASL, NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close

    rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_RICETTE, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsMain
        .AddNew
        .Fields("NOME_ASL") = "Ambito ASL "
        If rsDataset("TOTALE_Q") <> 0 Then
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_RICETTE")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTO") = .Fields("TOTALE_LORDO") - rsDataset("TOTALE_SCONTATO")
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND (NOT FLAG=3) AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            .Fields("TOTALE_NETTO_SCONTATO") = .Fields("TOTALE_NETTO") - .Fields("TOTALE_SCONTO")
        Else
            .Fields("PRESTAZIONI") = 0
            .Fields("RICETTE") = 0
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTO") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
            .Fields("TOTALE_NETTO_SCONTATO") = 0
        End If
        totalePrestazioni = .Fields("PRESTAZIONI")
        intTotaleRicette = .Fields("RICETTE")
        totaleLordo = .Fields("TOTALE_LORDO")
        totaleSconto = .Fields("TOTALE_SCONTO")
        totaleTicket = .Fields("TOTALE_TICKET")
        totaleQuotaAggiuntiva = .Fields("TOTALE_QUOTA_AGGIUNTIVA")
        totaleQuotaNazionale = .Fields("TOTALE_QUOTA_NAZIONALE")
        totaleNetto = .Fields("TOTALE_NETTO")
        totaleNettoScontato = .Fields("TOTALE_NETTO_SCONTATO")
        .Update
    End With
    rsDataset.Close
    
    rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_RICETTE, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsMain
        .AddNew
        .Fields("NOME_ASL") = "Fuori ASL"
        If rsDataset("TOTALE_Q") <> 0 Then
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_RICETTE")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTO") = .Fields("TOTALE_LORDO") - rsDataset("TOTALE_SCONTATO")
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            .Fields("TOTALE_NETTO_SCONTATO") = .Fields("TOTALE_NETTO") - .Fields("TOTALE_SCONTO")
        Else
            .Fields("PRESTAZIONI") = 0
            .Fields("RICETTE") = 0
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTO") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
            .Fields("TOTALE_NETTO_SCONTATO") = 0
        End If
        totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
        intTotaleRicette = intTotaleRicette + .Fields("RICETTE")
        totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
        totaleSconto = totaleSconto + .Fields("TOTALE_SCONTO")
        totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
        totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
        totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
        totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
        totaleNettoScontato = totaleNettoScontato + .Fields("TOTALE_NETTO_SCONTATO")
        .Update
    End With
    rsDataset.Close
    
    rsDataset.Open "SELECT COUNT(R.KEY) AS TOTALE_RICETTE, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsMain
        .AddNew
        .Fields("NOME_ASL") = "Fuori Regione"
        If rsDataset("TOTALE_Q") <> 0 Then
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_RICETTE")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTO") = .Fields("TOTALE_LORDO") - rsDataset("TOTALE_SCONTATO")
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            .Fields("TOTALE_NETTO_SCONTATO") = .Fields("TOTALE_NETTO") - .Fields("TOTALE_SCONTO")
        Else
            .Fields("PRESTAZIONI") = 0
            .Fields("RICETTE") = 0
            .Fields("TOTALE_LORDO") = 0
            .Fields("TOTALE_SCONTO") = 0
            .Fields("TOTALE_TICKET") = 0
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = 0
            .Fields("TOTALE_QUOTA_NAZIONALE") = 0
            .Fields("TOTALE_NETTO") = 0
            .Fields("TOTALE_NETTO_SCONTATO") = 0
        End If
        totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
        intTotaleRicette = intTotaleRicette + .Fields("RICETTE")
        totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
        totaleSconto = totaleSconto + .Fields("TOTALE_SCONTO")
        totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
        totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
        totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
        totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
        totaleNettoScontato = totaleNettoScontato + .Fields("TOTALE_NETTO_SCONTATO")
        .Update
    End With
    rsDataset.Close
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE"
        .Fields("PRESTAZIONI") = totalePrestazioni
        .Fields("RICETTE") = intTotaleRicette
        .Fields("TOTALE_LORDO") = Round(totaleLordo, 2)
        .Fields("TOTALE_SCONTO") = Round(totaleSconto, 2)
        .Fields("TOTALE_TICKET") = Round(totaleTicket, 2)
        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = Round(totaleQuotaAggiuntiva, 2)
        .Fields("TOTALE_QUOTA_NAZIONALE") = Round(totaleQuotaNazionale, 2)
        .Fields("TOTALE_NETTO") = Round(totaleNetto, 2)
        .Fields("TOTALE_NETTO_SCONTATO") = Round(totaleNettoScontato, 2)
        .Update
    End With
    
    If totalePrestazioni = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.NOME AS ASLNOME, COMUNI.NOME AS COMUNINOME " & _
                 "FROM      ((INTESTAZIONE_FATTURA " & _
                 "          INNER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
                 "          INNER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE) "
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblAsl").Caption = "ASL " & rsDataset("ASLNOME")
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblCap").Caption = rsDataset("CAP")
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblProvincia").Caption = rsDataset("COMUNINOME") & " (" & rsDataset("PROV") & ")"
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblIva").Caption = rsDataset("P_IVA")
            rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblCodiceFiscaleFattura").Caption = rsDataset("CODICE_FISCALE")
             
            rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblDicitura").Caption = rsDataset("DICITURA")                                                                        '
            rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblTotaleFattura").Caption = Format((totaleNettoScontato) + rsDataset("IMPORTO_BOLLO"), "#,##0.00")
                                                                                                                                                                          
            If chkNumeroAutorizzazione.Value = Checked Then
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblAutorizzazione").Caption = "Autorizzazione N° " & rsDataset("NUMERO_AUTORIZZAZIONE")
            Else
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblAutorizzazione").Caption = ""
            End If
            
            If chkBollo.Value = Checked Then
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblBollo").Caption = "Obbligo bollo assolto in maniera virtuale"
            Else
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblBollo").Caption = ""
            End If
                                                                                           
            If chkImportoBollo.Value = Checked Then
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblImportoBollo").Caption = "€ " & rsDataset("IMPORTO_BOLLO")
            Else
                rptFatturaAslNapoli2Nord.Sections("Pie").Controls("lblImportoBollo").Caption = ""
            End If
                                        
        End If
        rsDataset.Close
        rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblNumeroFattura").Caption = IIf(napoli3 And Option1, txtNumFattura + 1, txtNumFattura) & " / " & cboAnno.Text
        rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblData").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptFatturaAslNapoli2Nord.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptFatturaAslNapoli2Nord.DataSource = rsMain
        rptFatturaAslNapoli2Nord.RightMargin = 0
        rptFatturaAslNapoli2Nord.LeftMargin = 0
        rptFatturaAslNapoli2Nord.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaFattura2()
    On Error GoTo gestione
    Dim SQLString As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim nomeAsl As String
    Dim primo As Boolean
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
    ' totale dei totali
    Dim totaleRicette As Integer
    Dim totalePrestazioni As Integer
    Dim totaleScontato As Currency
    Dim totaleLordo As Currency
    Dim totaleNetto As Currency
    Dim totaleTicket As Currency
    Dim totaleQuotaAggiuntiva As Currency
    Dim totaleQuotaNazionale As Currency
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(30) AS NOME_ASL, " & _
                "       NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "       NEW adInteger AS RICETTE, " & _
                "       NEW adInteger AS PRESTAZIONI, " & _
                "       NEW adCurrency AS IMPORTO, " & _
                "       NEW adCurrency AS IMPORTO_SCONTATO, " & _
                "       NEW adCurrency AS TOTALE_LORDO, " & _
                "       NEW adCurrency AS TOTALE_NETTO, " & _
                "       NEW adCurrency AS SCONTO, " & _
                "       NEW adCurrency AS TOTALE_SCONTATO, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE, " & _
                "       NEW adCurrency AS TOTALE_TICKET "
         
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    

    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT CODICE_ASL, NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close

    primo = True
    rsDataset.Open "SELECT CODICE_PRESTAZIONE, COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 GROUP BY CODICE_PRESTAZIONE ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            If primo Then
                .Fields("NOME_ASL") = "ASL " & nomeAsl
            Else
                .Fields("NOME_ASL") = ""
            End If
            primo = False
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND (NOT FLAG=3) AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("SCONTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_SCONTATO")
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            
            rsAppo.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE KEY=" & rsDataset("CODICE_PRESTAZIONE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("IMPORTO") = rsAppo("IMPORTO")
            .Fields("IMPORTO_SCONTATO") = rsAppo("IMPORTO_SCONTATO")
            .Fields("CODICE_PRESTAZIONE") = rsAppo("CODICE")
            rsAppo.Close
            .Update
            
            totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
            totaleScontato = totaleScontato + .Fields("TOTALE_SCONTATO")
            totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
            totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
            totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
            totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    
    
   'stampa fattura pazienti asl (la prima)
    If napoli3 And Option1 Or Option2 Then
        If totaleRicette <> 0 Then
            ' totali dei totali
            With rsMain
                .AddNew
                .Update
                .AddNew
                .Fields("NOME_ASL") = "TOTALE"
                .Fields("PRESTAZIONI") = totalePrestazioni
                .Fields("RICETTE") = totaleRicette
                .Fields("TOTALE_LORDO") = Round(totaleLordo, 2)
                .Fields("TOTALE_SCONTATO") = Round(totaleScontato, 2)
                .Fields("TOTALE_TICKET") = Round(totaleTicket, 2)
                .Fields("TOTALE_QUOTA_AGGIUNTIVA") = Round(totaleQuotaAggiuntiva, 2)
                .Fields("TOTALE_QUOTA_NAZIONALE") = Round(totaleQuotaNazionale, 2)
                .Fields("TOTALE_NETTO") = Round(totaleNetto, 2)
                .Fields("SCONTO") = Round(totaleLordo, 2) - Round(totaleScontato, 2)
                .Fields("IMPORTO") = 0
                .Fields("IMPORTO_SCONTATO") = 0
                .Fields("CODICE_PRESTAZIONE") = ""
                .Update
            End With
            strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.NOME AS ASLNOME, COMUNI.NOME AS COMUNINOME " & _
                 "FROM      ((INTESTAZIONE_FATTURA " & _
                 "          INNER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
                 "          INNER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE) "
            rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                rptFattura2.Sections("intestazione").Controls("lblAsl").Caption = "ASL " & rsDataset("ASLNOME")
                rptFattura2.Sections("intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
                rptFattura2.Sections("intestazione").Controls("lblCap").Caption = rsDataset("CAP")
                rptFattura2.Sections("intestazione").Controls("lblProvincia").Caption = rsDataset("COMUNINOME") & " (" & rsDataset("PROV") & ")"
                rptFattura2.Sections("intestazione").Controls("lblIva").Caption = rsDataset("P_IVA")
                rptFattura2.Sections("pie").Controls("lblDicitura").Caption = rsDataset("DICITURA")
                rptFattura2.Sections("pie").Controls("lblIntestatario").Caption = rsDataset("INTESTATARIO_CC")
                rptFattura2.Sections("pie").Controls("lblIban").Caption = rsDataset("IBAN")
            
                If chkNumeroAutorizzazione.Value = Checked Then
                    rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = "Autorizzazione N° " & rsDataset("NUMERO_AUTORIZZAZIONE")
                Else
                    rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = ""
                End If
            
                If chkBollo.Value = Checked Then
                    rptFattura2.Sections("pie").Controls("lblBollo").Caption = "Obbligo bollo assolto in maniera virtuale"
                Else
                    rptFattura2.Sections("pie").Controls("lblBollo").Caption = ""
                End If
                
            End If
            rsDataset.Close
            rptFattura2.Sections("intestazione").Controls("lblNumeroFattura").Caption = txtNumFattura.Text & " / " & cboAnno.Text
            rptFattura2.Sections("intestazione").Controls("lblData").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
            rptFattura2.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
            rptFattura2.Sections("intestazione").Controls("lblTicket").Caption = "Ticket    " & Format(ticket, "###.00") & " €"
            Set rptFattura2.DataSource = rsMain
            rptFattura2.RightMargin = 0
            rptFattura2.LeftMargin = 0
            rptFattura2.PrintReport False, rptRangeAllPages
        End If
        rsMain.Close
        rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        totaleLordo = 0
        totaleScontato = 0
        totalePrestazioni = 0
        totaleRicette = 0
        totaleTicket = 0
        totaleNetto = 0
        totaleQuotaAggiuntiva = 0
        totaleQuotaNazionale = 0
    End If
    
    primo = True
    rsDataset.Open "SELECT CODICE_PRESTAZIONE, COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 GROUP BY CODICE_PRESTAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            If primo Then
                .Fields("NOME_ASL") = "REGIONE CAMPANIA FUORI ASL"
            Else
                .Fields("NOME_ASL") = ""
            End If
            primo = False
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
             
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND CODICE_REGIONE=16 AND NOT CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND (NOT FLAG=3) AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("SCONTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_SCONTATO")
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            
            rsAppo.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE KEY=" & rsDataset("CODICE_PRESTAZIONE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("IMPORTO") = rsAppo("IMPORTO")
            .Fields("IMPORTO_SCONTATO") = rsAppo("IMPORTO_SCONTATO")
            .Fields("CODICE_PRESTAZIONE") = rsAppo("CODICE")
            rsAppo.Close
            
            .Update
            totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
            totaleScontato = totaleScontato + .Fields("TOTALE_SCONTATO")
            totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
            totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
            totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
            totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    
    
'stampa fattura pazienti regione campania fuori asl (la seconda)
    If napoli3 And Option2 Then
        If totaleRicette <> 0 Then
          ' totali dei totali
            With rsMain
                .AddNew
                .Update
                .AddNew
                .Fields("NOME_ASL") = "TOTALE"
                .Fields("PRESTAZIONI") = totalePrestazioni
                .Fields("RICETTE") = totaleRicette
                .Fields("TOTALE_LORDO") = Round(totaleLordo, 2)
                .Fields("TOTALE_SCONTATO") = Round(totaleScontato, 2)
                .Fields("TOTALE_TICKET") = Round(totaleTicket, 2)
                .Fields("TOTALE_QUOTA_AGGIUNTIVA") = Round(totaleQuotaAggiuntiva, 2)
                .Fields("TOTALE_QUOTA_NAZIONALE") = Round(totaleQuotaNazionale, 2)
                .Fields("TOTALE_NETTO") = Round(totaleNetto, 2)
                .Fields("SCONTO") = Round(totaleLordo, 2) - Round(totaleScontato, 2)
                .Fields("IMPORTO") = 0
                .Fields("IMPORTO_SCONTATO") = 0
                .Fields("CODICE_PRESTAZIONE") = ""
                .Update
            End With

            strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.NOME AS ASLNOME, COMUNI.NOME AS COMUNINOME " & _
                     "FROM      ((INTESTAZIONE_FATTURA " & _
                     "          INNER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
                     "          INNER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE) "
            rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                rptFattura2.Sections("intestazione").Controls("lblAsl").Caption = "ASL " & rsDataset("ASLNOME")
                rptFattura2.Sections("intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
                rptFattura2.Sections("intestazione").Controls("lblCap").Caption = rsDataset("CAP")
                rptFattura2.Sections("intestazione").Controls("lblProvincia").Caption = rsDataset("COMUNINOME") & " (" & rsDataset("PROV") & ")"
                rptFattura2.Sections("intestazione").Controls("lblIva").Caption = rsDataset("P_IVA")
                rptFattura2.Sections("pie").Controls("lblDicitura").Caption = rsDataset("DICITURA")
                rptFattura2.Sections("pie").Controls("lblIntestatario").Caption = rsDataset("INTESTATARIO_CC")
                rptFattura2.Sections("pie").Controls("lblIban").Caption = rsDataset("IBAN")
            
                If chkNumeroAutorizzazione.Value = Checked Then
                    rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = "Autorizzazione N° " & rsDataset("NUMERO_AUTORIZZAZIONE")
                Else
                    rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = ""
                End If
            
                If chkBollo.Value = Checked Then
                    rptFattura2.Sections("pie").Controls("lblBollo").Caption = "Obbligo bollo assolto in maniera virtuale"
                Else
                    rptFattura2.Sections("pie").Controls("lblBollo").Caption = ""
                End If
            End If
            rsDataset.Close
            rptFattura2.Sections("intestazione").Controls("lblNumeroFattura").Caption = IIf(napoli3 And Option2, txtNumFattura + 1, txtNumFattura) & " / " & cboAnno.Text
            rptFattura2.Sections("intestazione").Controls("lblData").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
            rptFattura2.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
            rptFattura2.Sections("intestazione").Controls("lblTicket").Caption = "Ticket    " & Format(ticket, "###.00") & " €"
            Set rptFattura2.DataSource = rsMain
            rptFattura2.RightMargin = 0
            rptFattura2.LeftMargin = 0
            rptFattura2.PrintReport False, rptRangeAllPages
        End If
        
        rsMain.Close
        rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        totaleLordo = 0
        totaleScontato = 0
        totalePrestazioni = 0
        totaleRicette = 0
        totaleTicket = 0
        totaleNetto = 0
        totaleQuotaAggiuntiva = 0
        totaleQuotaNazionale = 0
        
    End If
    
    primo = True
    rsDataset.Open "SELECT CODICE_PRESTAZIONE, COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q, SUM(IMPORTO*QUANTITA) AS TOTALE_LORDO, SUM(IMPORTO_SCONTATO*QUANTITA) AS TOTALE_SCONTATO  FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 GROUP BY CODICE_PRESTAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            If primo Then
                .Fields("NOME_ASL") = "FUORI REGIONE CAMPANIA"
            Else
                .Fields("NOME_ASL") = ""
            End If
            primo = False
            .Fields("PRESTAZIONI") = rsDataset("TOTALE_Q")
            .Fields("RICETTE") = rsDataset("TOTALE_R")
            .Fields("TOTALE_LORDO") = rsDataset("TOTALE_LORDO")
            .Fields("TOTALE_SCONTATO") = rsDataset("TOTALE_SCONTATO")
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 AND (R.CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND NOT CODICE_REGIONE=16 AND NOT FLAG=3 AND R.CODICE_ESENZIONE=-1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND NOT CODICE_REGIONE=16 AND (NOT FLAG=3) AND (NOT R.CODICE_ESENZIONE=-1) AND ESENZIONE_QUOTA=FALSE AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("SCONTO") = .Fields("TOTALE_LORDO") - .Fields("TOTALE_SCONTATO")
            .Fields("TOTALE_NETTO") = .Fields("TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")

            rsAppo.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE KEY=" & rsDataset("CODICE_PRESTAZIONE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("IMPORTO") = rsAppo("IMPORTO")
            .Fields("IMPORTO_SCONTATO") = rsAppo("IMPORTO_SCONTATO")
            .Fields("CODICE_PRESTAZIONE") = rsAppo("CODICE")
            rsAppo.Close
            
            .Update
            totaleLordo = totaleLordo + .Fields("TOTALE_LORDO")
            totaleScontato = totaleScontato + .Fields("TOTALE_SCONTATO")
            totaleNetto = totaleNetto + .Fields("TOTALE_NETTO")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            totalePrestazioni = totalePrestazioni + .Fields("PRESTAZIONI")
            totaleTicket = totaleTicket + .Fields("TOTALE_TICKET")
            totaleQuotaAggiuntiva = totaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
            totaleQuotaNazionale = totaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE"
        .Fields("PRESTAZIONI") = totalePrestazioni
        .Fields("RICETTE") = totaleRicette
        .Fields("TOTALE_LORDO") = Round(totaleLordo, 2)
        .Fields("TOTALE_SCONTATO") = Round(totaleScontato, 2)
        .Fields("TOTALE_TICKET") = Round(totaleTicket, 2)
        .Fields("TOTALE_QUOTA_AGGIUNTIVA") = Round(totaleQuotaAggiuntiva, 2)
        .Fields("TOTALE_QUOTA_NAZIONALE") = Round(totaleQuotaNazionale, 2)
        .Fields("TOTALE_NETTO") = Round(totaleNetto, 2)
        .Fields("SCONTO") = Round(totaleLordo, 2) - Round(totaleScontato, 2)
        .Fields("IMPORTO") = 0
        .Fields("IMPORTO_SCONTATO") = 0
        .Fields("CODICE_PRESTAZIONE") = ""
        .Update
    End With
    
    If totaleRicette = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.NOME AS ASLNOME, COMUNI.NOME AS COMUNINOME " & _
                "FROM      ((INTESTAZIONE_FATTURA " & _
                "          INNER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
                "          INNER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE) "
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            rptFattura2.Sections("intestazione").Controls("lblAsl").Caption = "ASL " & rsDataset("ASLNOME")
            rptFattura2.Sections("intestazione").Controls("lblIndirizzo").Caption = rsDataset("INDIRIZZO")
            rptFattura2.Sections("intestazione").Controls("lblCap").Caption = rsDataset("CAP")
            rptFattura2.Sections("intestazione").Controls("lblProvincia").Caption = rsDataset("COMUNINOME") & " (" & rsDataset("PROV") & ")"
            rptFattura2.Sections("intestazione").Controls("lblIva").Caption = rsDataset("P_IVA")
            rptFattura2.Sections("pie").Controls("lblDicitura").Caption = rsDataset("DICITURA")
            rptFattura2.Sections("pie").Controls("lblIntestatario").Caption = rsDataset("INTESTATARIO_CC")
            rptFattura2.Sections("pie").Controls("lblIban").Caption = rsDataset("IBAN")
            
            If chkNumeroAutorizzazione.Value = Checked Then
                rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = "Autorizzazione N° " & rsDataset("NUMERO_AUTORIZZAZIONE")
            Else
                rptFattura2.Sections("pie").Controls("lblAutorizzazione").Caption = ""
            End If
            
            If chkBollo.Value = Checked Then
                rptFattura2.Sections("pie").Controls("lblBollo").Caption = "Obbligo bollo assolto in maniera virtuale"
            Else
                rptFattura2.Sections("pie").Controls("lblBollo").Caption = ""
            End If
            
        End If
        rsDataset.Close
        If napoli3 And Option1 Then
            numfat = txtNumFattura + 1
        ElseIf napoli3 And Option2 Then
            numfat = txtNumFattura + 2
        Else
            numfat = txtNumFattura
        End If
        rptFattura2.Sections("intestazione").Controls("lblNumeroFattura").Caption = numfat & " / " & cboAnno.Text
        rptFattura2.Sections("intestazione").Controls("lblData").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptFattura2.Sections("intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        rptFattura2.Sections("intestazione").Controls("lblTicket").Caption = "Ticket    " & Format(ticket, "###.00") & " €"
        Set rptFattura2.DataSource = rsMain
        rptFattura2.RightMargin = 0
        rptFattura2.LeftMargin = 0
        rptFattura2.PrintReport (Not (napoli3 And Option1 Or Option2)), rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing

    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub StampaPerMazzetteDistretti()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim nomeAsl As String
    Dim i As Integer
    
    Dim totaleRicette As Integer
    Dim totaleQuantita As Integer
    Dim totaleTotaleRicette As Integer
    Dim totaleTotaleQuantita As Integer
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter

    If cboMese.ListIndex >= 5 And cboMese.ListIndex <= 8 And cboAnno.Text = 2010 Then
        MsgBox "Impossibile stampare solo per il mese di " & cboMese.Text & " " & cboAnno.Text & vbCrLf & "Selezionare l'opzione Stampa Giu. - Sett. 2010 in Stampa Mazzette Mensili", vbCritical, "Attenzione"
        Exit Sub
    End If
        

    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(20) AS MAZZETTE, " & _
                "       NEW adVarChar(35) AS NOME_ASL, " & _
                "       NEW adVarChar(4) AS NOME_DISTRETTO, " & _
                "       NEW adInteger  AS RICETTE, " & _
                "       NEW adInteger AS QUANTITA"

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT CODICE_ASL, NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close

    ' asl del centro
    rsDataset.Open "SELECT DISTINCT D.KEY, A.NOME, D.NOME FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN ASL A ON A.KEY=P.CODICE_ASL) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND D.CODICE_ASL=" & structIntestazione.sCodiceAsl, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("NOME_ASL") = nomeAsl
            .Fields("NOME_DISTRETTO") = rsDataset("D.NOME")
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=" & rsDataset("KEY") & " AND NOT FLAG=3 ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("RICETTE") = rsAppo("TOTALE_R")
            .Fields("QUANTITA") = rsAppo("TOTALE_Q")
            For i = 0 To rsAppo("TOTALE_R") \ 51
                .Fields("MAZZETTE") = .Fields("MAZZETTE") & i + 1 & " - "
            Next
            .Fields("MAZZETTE") = Left(.Fields("MAZZETTE"), Len(.Fields("MAZZETTE")) - 3)
            rsAppo.Close
            totaleQuantita = totaleQuantita + .Fields("QUANTITA")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            .Update
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE ASL " & nomeAsl
        .Fields("NOME_DISTRETTO") = ""
        .Fields("RICETTE") = totaleRicette
        .Fields("QUANTITA") = totaleQuantita
        .Fields("MAZZETTE") = ""
        .AddNew
        .Update
    End With
    totaleTotaleQuantita = totaleTotaleQuantita + totaleQuantita
    totaleTotaleRicette = totaleTotaleRicette + totaleRicette
    totaleQuantita = 0
    totaleRicette = 0

    ' asl in campania non del centro
    rsDataset.Open "SELECT DISTINCT D.KEY, A.NOME, D.NOME FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN ASL A ON A.KEY=P.CODICE_ASL) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND NOT D.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND A.CODICE_REGIONE=16", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("NOME_ASL") = rsDataset("A.NOME")
            .Fields("NOME_DISTRETTO") = rsDataset("D.NOME")
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_DISTRETTO=" & rsDataset("KEY") & " AND NOT FLAG=3 ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("RICETTE") = rsAppo("TOTALE_R")
            .Fields("QUANTITA") = rsAppo("TOTALE_Q")
            For i = 0 To rsAppo("TOTALE_R") / 51
                .Fields("MAZZETTE") = .Fields("MAZZETTE") & i + 1 & " - "
            Next
            .Fields("MAZZETTE") = Left(.Fields("MAZZETTE"), Len(.Fields("MAZZETTE")) - 3)
            rsAppo.Close
            totaleQuantita = totaleQuantita + .Fields("QUANTITA")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            .Update
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE REGIONE CAMPANIA FUORI ASL"
        .Fields("NOME_DISTRETTO") = ""
        .Fields("RICETTE") = totaleRicette
        .Fields("QUANTITA") = totaleQuantita
        .Fields("MAZZETTE") = ""
        .AddNew
        .Update
    End With
    totaleTotaleQuantita = totaleTotaleQuantita + totaleQuantita
    totaleTotaleRicette = totaleTotaleRicette + totaleRicette
    totaleQuantita = 0
    totaleRicette = 0
    
    ' asl fuori campania
    rsDataset.Open "SELECT DISTINCT D.KEY, A.NOME, D.NOME, A.KEY FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) LEFT OUTER JOIN ASL A ON A.KEY=P.CODICE_ASL) LEFT OUTER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 AND (NOT A.CODICE_REGIONE=16 OR ISNULL(A.CODICE_REGIONE)) ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("NOME_ASL") = rsDataset("A.NOME") & ""
            .Fields("NOME_DISTRETTO") = rsDataset("D.NOME") & ""
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R, SUM(QUANTITA) AS TOTALE_Q FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND CODICE_ASL=" & IIf(IsNull(rsDataset("A.KEY")), -1, rsDataset("A.KEY")) & IIf(IsNull(rsDataset("D.KEY")), " ", " AND CODICE_DISTRETTO=" & rsDataset("D.KEY")) & " AND NOT FLAG=3 ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("RICETTE") = rsAppo("TOTALE_R")
            .Fields("QUANTITA") = rsAppo("TOTALE_Q")
            For i = 0 To rsAppo("TOTALE_R") / 51
                .Fields("MAZZETTE") = .Fields("MAZZETTE") & i + 1 & " - "
            Next
            .Fields("MAZZETTE") = Left(.Fields("MAZZETTE"), Len(.Fields("MAZZETTE")) - 3)
            rsAppo.Close
            totaleQuantita = totaleQuantita + .Fields("QUANTITA")
            totaleRicette = totaleRicette + .Fields("RICETTE")
            .Update
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE FUORI REGIONE CAMPANIA"
        .Fields("NOME_DISTRETTO") = ""
        .Fields("RICETTE") = totaleRicette
        .Fields("QUANTITA") = totaleQuantita
        .Fields("MAZZETTE") = ""
        .AddNew
        .Update
    End With
    totaleTotaleQuantita = totaleTotaleQuantita + totaleQuantita
    totaleTotaleRicette = totaleTotaleRicette + totaleRicette
    
    ' totali dei totali
    With rsMain
        .AddNew
        .Update
        .AddNew
        .Fields("NOME_ASL") = "TOTALE GENERALE"
        .Fields("NOME_DISTRETTO") = ""
        .Fields("RICETTE") = totaleTotaleRicette
        .Fields("QUANTITA") = totaleTotaleQuantita
        .Fields("MAZZETTE") = ""
        .Update
    End With
    
    
    If rsMain.RecordCount = 2 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        rptRiepilogoMazzettePerDistretto.Sections("Intestazione").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptRiepilogoMazzettePerDistretto.DataSource = rsMain
        rptRiepilogoMazzettePerDistretto.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Sub StampaPerMazzettaSingola()
    On Error GoTo gestione
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio1 As Recordset
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rsRicette As Recordset
    Dim i As Integer
    Dim k As Integer
    Dim vecchioNumeroRicette As Integer
    Dim vecchioNumeroPrestazioni As Integer
    Const numRecordInPag As Integer = 24
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    If cboMese.ListIndex >= 5 And cboMese.ListIndex <= 8 And cboAnno.Text = 2010 Then
        MsgBox "Impossibile stampare solo per il mese di " & cboMese.Text & " " & cboAnno.Text & vbCrLf & "Selezionare l'opzione Stampa Giu. - Sett. 2010 in Stampa Mazzette Mensili", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adInteger AS NUMEROMAZZETTA1, " & _
                "       NEW adVarChar(5) AS NOME_DISTRETTO, " & _
                "       NEW adVarChar(35) AS NOME_ASL, " & _
                "       NEW adInteger AS RICETTE, " & _
                "       NEW adInteger AS PRESTAZIONI, " & _
                "       NEW adVarChar(5) AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adVarChar(5) AS LINK1, " & _
                "           NEW adInteger AS NUMEROMAZZETTA2, " & _
                "           NEW adVarChar(5) AS NOME_DISTRETTO, " & _
                "           NEW adVarChar(25) AS NOME_ASL, " & _
                "           NEW adInteger AS PROGRESSIVO_RICETTA, " & _
                "           NEW adVarChar(16) AS NUMERO_RICETTA, " & _
                "           NEW adInteger AS PRESTAZIONI, " & _
                "           NEW adVarChar(35) AS COGNOME, " & _
                "           NEW adVarChar(35) AS NOME " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    Set rsRicette = New Recordset
    
    With rsMain
        rsDataset.Open "SELECT DISTINCT MAZZETTA1 FROM RICETTE WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 ORDER BY MAZZETTA1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        i = 0
        Do While Not rsDataset.EOF
            rsAppo.Open "SELECT DISTINCT D.KEY, D.NOME, A.NOME, A.KEY FROM (((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) LEFT OUTER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) LEFT OUTER JOIN ASL A ON A.KEY=P.CODICE_ASL) WHERE MAZZETTA1=" & rsDataset("MAZZETTA1") & " AND MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsAppo.EOF
                .AddNew
                .Fields("NUMEROMAZZETTA1") = rsDataset("MAZZETTA1")
                .Fields("NOME_ASL") = "ASL " & rsAppo("A.NOME") & ""
                .Fields("NOME_DISTRETTO") = rsAppo("D.NOME") & ""
                rsRicette.Open "SELECT COUNT(R.KEY) AS TOTALER, SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MAZZETTA1=" & rsDataset("MAZZETTA1") & " AND CODICE_ASL=" & IIf(IsNull(rsAppo("A.KEY")), -1, rsAppo("A.KEY")) & IIf(IsNull(rsAppo("D.KEY")), " ", " AND CODICE_DISTRETTO=" & rsAppo("D.KEY")) & " AND MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                .Fields("RICETTE") = rsRicette("TOTALER")
                .Fields("PRESTAZIONI") = rsRicette("TOTALEQ")
                rsRicette.Close
                .Fields("LINK1") = i
                rsRicette.Open "SELECT MAZZETTA2, P.NOME, P.COGNOME, QUANTITA, NUMERO_RICETTA, PROGRESSIVO_RICETTA FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE MAZZETTA1=" & rsDataset("MAZZETTA1") & " AND CODICE_ASL=" & IIf(IsNull(rsAppo("A.KEY")), -1, rsAppo("A.KEY")) & IIf(IsNull(rsAppo("D.KEY")), " ", " AND CODICE_DISTRETTO=" & rsAppo("D.KEY")) & " AND MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3 ORDER BY MAZZETTA2", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                k = 0
                Do While Not rsRicette.EOF
                    Set rsFiglio1 = .Fields("Res1").Value
                    With rsFiglio1
                        .AddNew
                        .Fields("LINK1") = i
                        .Fields("NUMEROMAZZETTA2") = rsRicette("MAZZETTA2")
                        .Fields("NOME_ASL") = rsAppo("A.NOME")
                        .Fields("NOME_DISTRETTO") = rsAppo("D.NOME")
                        .Fields("PROGRESSIVO_RICETTA") = rsRicette("PROGRESSIVO_RICETTA")
                        .Fields("NUMERO_RICETTA") = rsRicette("NUMERO_RICETTA")
                        .Fields("PRESTAZIONI") = rsRicette("QUANTITA")
                        .Fields("COGNOME") = rsRicette("COGNOME")
                        .Fields("NOME") = rsRicette("NOME")
                        .Update
                    End With
                    k = k + 1
                    If k = numRecordInPag Then
                        vecchioNumeroRicette = rsMain("RICETTE")
                        vecchioNumeroPrestazioni = rsMain("PRESTAZIONI")
                        rsMain.Update
                        rsMain.AddNew
                        rsMain.Fields("NUMEROMAZZETTA1") = rsDataset("MAZZETTA1")
                        rsMain.Fields("NOME_ASL") = "ASL " & rsAppo("A.NOME")
                        rsMain.Fields("NOME_DISTRETTO") = rsAppo("D.NOME")
                        rsMain.Fields("RICETTE") = vecchioNumeroRicette
                        rsMain.Fields("PRESTAZIONI") = vecchioNumeroPrestazioni
                        i = i + 1
                        k = 0
                        rsMain.Fields("LINK1") = i
                    End If
                    rsRicette.MoveNext
                Loop
                rsRicette.Close
                rsAppo.MoveNext
                i = i + 1
            Loop
            rsAppo.Close
            rsDataset.MoveNext
        Loop
        rsDataset.Close
    End With
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, Me.Caption
    Else
        rptRiepilogoMazzettaSingola.Sections("livello1").Controls("lblMese").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptRiepilogoMazzettaSingola.DataSource = rsMain
        rptRiepilogoMazzettaSingola.LeftMargin = 0
        rptRiepilogoMazzettaSingola.RightMargin = 0
        rptRiepilogoMazzettaSingola.PrintReport False, rptRangeAllPages
    End If
    
    Set rsRicette = Nothing
    Set rsAppo = Nothing
    Set rsDataset = Nothing
    
gestione:
    If Err.Number = cdlCancel Then      ' se clicco ANNULLA nella finestra di scelta Stampante
        Exit Sub
    End If
End Sub

Private Function Completo() As Boolean
    Completo = False
        If cboMese.ListIndex = -1 Then
            MsgBox "Selezionare il mese", vbCritical, "Attenzione"
            Exit Function
        End If
        If cboAnno.ListIndex = -1 Then
            MsgBox "Selezionare l'anno", vbCritical, "Attenzione"
            Exit Function
        End If
        If tStampeRiepilogo = tpFATTURA Then
            If txtNumFattura = "" Then
                MsgBox "Inserire il numero di fattura", vbCritical, "Attenzione"
                Exit Function
            End If
        End If
    Completo = True
End Function

Private Sub SceltaFattura()
    ' se asl di caserta, avellino o benevento stampa fattura 1
    ' se asl na1 e na3sud stampa fattura 2
    ' Asl Napoli 2 Nord codice asl 5
    
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT CODICE_ASL FROM INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset("CODICE_ASL") = 5 Then
        Call StampaPerAslNapoli2Nord
    Else
        If rsDataset("CODICE_ASL") >= 4 And rsDataset("CODICE_ASL") <= 6 Then
            Call StampaFattura2
        Else
            Call StampaFattura
        End If
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    If Completo Then
        Select Case tStampeRiepilogo
            Case TPXMAZZETTEMENSILI: Call StampaPerMazzetteMensili
            Case tpXMAZZETTASINGOLA: Call StampaPerMazzettaSingola
            Case tpXMAZZETTEDISTRETTI: Call StampaPerMazzetteDistretti
            Case tpXPAZIENTE: Call StampaPerPaziente
            Case tpFATTURA: Call SceltaFattura
            Case tpXTOTALIPERPRESTAZIONE: Call StampaPerTotaliPerPrestazione
            Case tpXASLDISTRETTI: Call StampaPerAslDistretto
            Case tpXTOTALIPERASL: Call StampaTotaliPerAsl
            Case tpXIMPEGNATIVE: Call StampaRiepilogoImpegnative
        End Select
    End If
End Sub

Private Sub txtNumFattura_GotFocus()
    txtNumFattura.BackColor = colArancione
End Sub

Private Sub txtNumFattura_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumFattura_LostFocus()
    txtNumFattura.BackColor = vbWhite
End Sub

'' Scrive la data secondo lo standard del file XML
'
' @param data data da modificare
Private Function sistemaData(data As Date) As String
    sistemaData = Year(data) & "-" & Format(Month(data), "00") & "-" & Format(Day(data), "00")
End Function

'' Crea un singolo nodo del file XML
'
' @param nome nome del nodo
' @param valore valore da inserire nel nodo
' @return nodo da aggiungere al documento XML
Private Function CreaNodo(nome As String, valore As String) As IXMLDOMNode
    Dim nodo As IXMLDOMNode
    Set nodo = doc.createElement(nome)
    nodo.Text = valore
    Set CreaNodo = nodo
End Function

Private Sub fattelettr_Click()
        Dim rsDataset As New Recordset
        rsDataset.Open "SELECT * FROM RICETTE  WHERE (ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsDataset.EOF And rsDataset.BOF Then
            MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, "Genera file XML"
            rsDataset.Close
            Exit Sub
        Else
            Set rsDataset = Nothing
        End If
 
    If Completo = False Then          ' controlla che ci sia il n° della fattura
        Exit Sub
    End If
    
    frmFatEle.Show 1
 
    If OKGeneraFE = False Then     ' Esce se si preme il tasto CHIUDI nel formFatEle
        Exit Sub
    End If
 
    OKGeneraFE = False
 
    Dim proc As IXMLDOMProcessingInstruction
    Dim nodo0 As IXMLDOMNode
    Dim nodo1 As IXMLDOMNode
    Dim nodo2 As IXMLDOMNode
    Dim nodo3 As IXMLDOMNode
    Dim attr As IXMLDOMAttribute
    Dim root As IXMLDOMElement
    Dim frag As IXMLDOMDocumentFragment
    
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsAppo As New Recordset
    Dim rsFE As New Recordset
    Dim ticket As Currency
    Dim quotaAggiuntiva As Currency
    Dim quotaNazionale As Currency
    Dim totaleRicette As Integer
    Dim coeffTicket As Integer
    Dim coeffQuota As Single
    Dim coeffQuotaNazionale As Single
    
  ' totale dei totali
    Dim totaleAsl As Integer
    Dim totaleRegione As Integer
    Dim totaleFuoriRegione As Integer
    Dim importoTotale As Currency
    Dim importoTotaleScontato As Currency
    Dim importoTotaleTicket As Currency
    Dim importoTotaleQuotaAggiuntiva As Currency
    Dim importoTotaleQuotaNazionale As Currency
    Dim importoTotaleNetto As Currency

    Dim MCodFisc As String
    Dim MCodiceAsl As Integer
    Dim MCodiceComune As Integer
    Dim MCod_Destinatario As String
    Dim MProgr_Invio As String
    Dim NameXML As String
    Dim NameExtXML As String
  '  Dim ret As Integer
           
  '  Dim rsDataset As New Recordset
  '  Dim rsAppo As New Recordset
  '  Dim codiceSTS As String
  '  Dim codiceAsl As String
  '  Dim totaleAssistito As Single
  '  Dim totaleScontato As Single
  '  Dim coefficienteQuotaAggiuntiva As Single
    
    ' verifica se ci sono ricette
'    rsDataset.Open "SELECT * FROM RICETTE  WHERE (ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If rsDataset.EOF And rsDataset.BOF Then
'        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, "Genera file XML"
'        rsDataset.Close
'        Exit Sub
'    Else
'        Set rsDataset = Nothing
'    End If
    
    Set doc = Nothing
    ' versione
    ' ?xml version="1.0" encoding="UTF-8"?
    ' ?xml-stylesheet type="text/xsl" href="fatturapa_v1.1xsl"?
    Set proc = doc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    doc.appendChild proc
    Set proc = doc.createProcessingInstruction("xml-stylesheet", "type='text/xsl' href='fatturapa_v1.1.xsl'")
    doc.appendChild proc
           
    '<p:FatturaElettronica versione="1.1"
    'xmlns:ds="http://www.w3.org/2000/09/xmldsig#"
    'xmlns:p="http://www.fatturapa.gov.it/sdi/fatturapa/v1.1"
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    
    ' root
    Set root = doc.createElement("p:FatturaElettronica")
    doc.appendChild root
    
    Set attr = doc.createAttribute("versione")
    attr.Value = "1.1"
    root.setAttributeNode attr
 
    Set attr = doc.createAttribute("xmlns:ds")
    attr.Value = "http://www.w3.org/2000/09/xmldsig#"
    root.setAttributeNode attr
    
    Set attr = doc.createAttribute("xmlns:p")
    attr.Value = "http://www.fatturapa.gov.it/sdi/fatturapa/v1.1"
    root.setAttributeNode attr
    
    Set attr = doc.createAttribute("xmlns:xsi")
    attr.Value = "http://www.w3.org/2001/XMLSchema-instance"
    root.setAttributeNode attr
    'doc.appendChild root

'-------------------------------------------------------------
    'crea 1° macroblocco
    Set nodo0 = doc.createElement("FatturaElettronicaHeader")

  ' DATI TRASMISSIONE punto 1.1
    rsDataset.Open "SELECT COD_DESTINATARIO,PROGR_INVIO,PROGR_INVIO_RIGEN FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
    MCod_Destinatario = rsDataset("COD_DESTINATARIO")
    
    If rsDataset("PROGR_INVIO_RIGEN") = 0 Then   ' controlla se va rigenerata la FE attraverso il progressivo d'invio
        MProgr_Invio = rsDataset("PROGR_INVIO")  ' se progr_invio_rigen = 0 incrementa il n° progressivo d'invio
        rsDataset("PROGR_INVIO") = MProgr_Invio + 1 'incrementa il numero progresso
    Else
        MProgr_Invio = rsDataset("PROGR_INVIO_RIGEN") 'annulla la rigenerazione della FE ponendo il valore = 0
        rsDataset("PROGR_INVIO_RIGEN") = 0
    End If
    rsDataset.Update
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    MCodFisc = rsDataset("CODICE_FISCALE")

    'crea 1° elemento
    Set nodo1 = doc.createElement("DatiTrasmissione")
        'aggiunge al 1° macroblocco il 1° elemento
        nodo0.appendChild nodo1
    
    'crea 2° elemento e nodi
    Set nodo2 = doc.createElement("IdTrasmittente")
        nodo2.appendChild CreaNodo("IdPaese", "IT")
        nodo2.appendChild CreaNodo("IdCodice", MCodFisc)
        'aggiunge al 1° elemento l'elemento i nodi del 2° elemento
        nodo1.appendChild nodo2
    
    'aggiunge al 1° elemento i nodi
    nodo1.appendChild CreaNodo("ProgressivoInvio", MProgr_Invio & "-" & cboAnno.Text)
    nodo1.appendChild CreaNodo("FormatoTrasmissione", "SDI11")
    nodo1.appendChild CreaNodo("CodiceDestinatario", MCod_Destinatario)
    
    'crea 2° elemento e nodi
    Set nodo2 = doc.createElement("ContattiTrasmittente")
       ' nodo2.appendChild CreaNodo("Telefono", "0000000")
        nodo2.appendChild CreaNodo("Email", rsDataset("MAIL"))
        'aggiunge al 1° elemento l'elemento i nodi del 2° elemento
        nodo1.appendChild nodo2
    
'..................................................
    'CEDENTE PRESTATORE punto 1.2
    
    'crea 1° elemento
    Set nodo1 = doc.createElement("CedentePrestatore")
        'aggiunge al 1° macroblocco il 1° elemento
        nodo0.appendChild nodo1
    
    'crea 2° elemento punto 1.2.1
    Set nodo2 = doc.createElement("DatiAnagrafici")
        'aggiunge al 1° elemento il 2° elemento
        nodo1.appendChild nodo2
        
    'crea 3° elemento e nodi punto 1.2.1.1
    Set nodo3 = doc.createElement("IdFiscaleIVA")
        nodo3.appendChild CreaNodo("IdPaese", "IT")
        nodo3.appendChild CreaNodo("IdCodice", rsDataset("IVA"))
        'aggiunge al 2° elemento l'elemento e i nodi del 3° elemento
        nodo2.appendChild nodo3
        
    ' punto 1.2.1.2
        nodo2.appendChild CreaNodo("CodiceFiscale", rsDataset("CODICE_FISCALE"))
        'aggiunge al 1° elemento i nodi
        nodo1.appendChild nodo2
            
    ' punto 1.2.1.3
    Set nodo3 = doc.createElement("Anagrafica")
        nodo3.appendChild CreaNodo("Denominazione", rsDataset("RAGIONE_SOCIALE"))
        'aggiunge al 2° elemento l'elemento e i nodi del 3° elemento
        nodo2.appendChild nodo3
    
    ' punto 1.2.1.8
    nodo2.appendChild CreaNodo("RegimeFiscale", rsDataset("REG_FISCALE"))
    'aggiunge al 1° elemento i nodi
    nodo1.appendChild nodo2
  
    ' punto 1.2.2
    Set nodo2 = doc.createElement("Sede")
        nodo2.appendChild CreaNodo("Indirizzo", rsDataset("INDIRIZZO"))
        nodo2.appendChild CreaNodo("CAP", rsDataset("CAP"))
        nodo2.appendChild CreaNodo("Comune", rsDataset("CITTA"))
        nodo2.appendChild CreaNodo("Provincia", rsDataset("PROV"))
        nodo2.appendChild CreaNodo("Nazione", "IT")
        'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
        nodo1.appendChild nodo2
    
    ' punto 1.2.4
    Set nodo2 = doc.createElement("IscrizioneREA")
        nodo2.appendChild CreaNodo("Ufficio", rsDataset("PR_UFF_REG"))
        nodo2.appendChild CreaNodo("NumeroREA", rsDataset("NUM_REA"))
        nodo2.appendChild CreaNodo("CapitaleSociale", VirgolaOrPunto(Format(rsDataset("CAP_SOCIALE"), "#####.00"), ","))
        If rsDataset("SRL") = True Then 'campo srl = si
            nodo2.appendChild CreaNodo("SocioUnico", rsDataset("SOCIO"))
        End If
        nodo2.appendChild CreaNodo("StatoLiquidazione", rsDataset("LIQUIDAZIONE"))
        'aggiunge al 1° elemento l'elementi e i nodi del 2° elemento
        nodo1.appendChild nodo2
   
    rsDataset.Close
'..................................................
    'CESSIONARIO COMMITTENTE punto 1.4
    
    rsDataset.Open "SELECT * FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'crea 1° elemento
    Set nodo1 = doc.createElement("CessionarioCommittente")
        'aggiunge al 1° macroblocco il 1° elemento
        nodo0.appendChild nodo1
        
    'crea 2° elemento punto 1.4.1
    Set nodo2 = doc.createElement("DatiAnagrafici")
        'aggiunge al 1° elemento il 2° elemento
        nodo1.appendChild nodo2
        
    'crea 3° elemento e nodi punto 1.4.1.1
    Set nodo3 = doc.createElement("IdFiscaleIVA")
        nodo3.appendChild CreaNodo("IdPaese", "IT")
        nodo3.appendChild CreaNodo("IdCodice", rsDataset("P_IVA"))
        'aggiunge al 2° elemento l'elemento e i nodi del 3° elemento
        nodo2.appendChild nodo3
    
    ' punto 1.4.1.3
    MCodiceAsl = rsDataset("CODICE_ASL")
    rsAppo.Open "SELECT NOME FROM ASL WHERE KEY=" & MCodiceAsl, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set nodo3 = doc.createElement("Anagrafica")
        nodo3.appendChild CreaNodo("Denominazione", "ASL " & rsAppo("NOME"))
        'aggiunge al 2° elemento l'elemento e i nodi del 3° elemento
        nodo2.appendChild nodo3
    rsAppo.Close
  
    ' punto 1.4.2
    MCodiceComune = rsDataset("CODICE_COMUNE")
    rsAppo.Open "SELECT NOME FROM COMUNI WHERE KEY=" & MCodiceComune, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set nodo2 = doc.createElement("Sede")
        nodo2.appendChild CreaNodo("Indirizzo", rsDataset("INDIRIZZO"))
        nodo2.appendChild CreaNodo("CAP", rsDataset("CAP"))
        nodo2.appendChild CreaNodo("Comune", rsAppo("NOME"))
        nodo2.appendChild CreaNodo("Provincia", rsDataset("PROV"))
        nodo2.appendChild CreaNodo("Nazione", "IT")
        'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
        nodo1.appendChild nodo2
            
    root.appendChild nodo0
    
    rsAppo.Close
    rsDataset.Close
'------------------------------------------------------------
    'crea 2° macroblocco
    Set nodo0 = doc.createElement("FatturaElettronicaBody")

   'DATI GENERALI punto 2.1

    'crea 1° elemento punto 2.1
    Set nodo1 = doc.createElement("DatiGenerali")
        'aggiunge al 2° elemento radice il 1° elemento
        nodo0.appendChild nodo1
        
    'crea 2° elemento e nodi punto 2.1.1
    Set nodo2 = doc.createElement("DatiGeneraliDocumento")
        nodo2.appendChild CreaNodo("TipoDocumento", "TD01")
        nodo2.appendChild CreaNodo("Divisa", "EUR")
        nodo2.appendChild CreaNodo("Data", sistemaData(GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)))
        nodo2.appendChild CreaNodo("Numero", txtNumFattura)
        'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
        nodo1.appendChild nodo2
    
    rsDataset.Open "SELECT NUMERO_AUTORIZZAZIONE,IMPORTO_BOLLO FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'crea 3° elemento e nodi punto 2.2.1.6
    Set nodo3 = doc.createElement("DatiBollo")
       nodo3.appendChild CreaNodo("BolloVirtuale", "SI") 'versione 1.1
'       nodo3.appendChild CreaNodo("NumeroBollo", "DM-17-GIU-2014") 'dicitura temporanea
        nodo3.appendChild CreaNodo("ImportoBollo", VirgolaOrPunto(Format(rsDataset("IMPORTO_BOLLO"), "#####.00"), ","))
        'aggiunge al 2° elemento l'elemento e i nodi del 3° elemento
        nodo2.appendChild nodo3

    rsDataset.Close
'..................................................
    'DATI BENI SERVIZI punto 2.2
    
   ' rsDataset.Open "SELECT CODICE_ASL,P_IVA,CODICE_FISCALE,INDIRIZZO,CAP,CODICE_COMUNE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'crea 1° elemento
    Set nodo1 = doc.createElement("DatiBeniServizi")
        'aggiunge al 2° macroblocco il 1° elemento
        nodo0.appendChild nodo1

    importoTotale = 0
    importoTotaleNetto = 0
    importoTotaleScontato = 0
    importoTotaleTicket = 0
    importoTotaleQuotaAggiuntiva = 0
    importoTotaleQuotaNazionale = 0
        
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(10) AS CODICE_PRESTAZIONE, " & _
                "       NEW adInteger AS TOTALE_ASL, " & _
                "       NEW adInteger AS TOTALE_REGIONE, " & _
                "       NEW adInteger AS TOTALE_FUORI_REGIONE, " & _
                "       NEW adCurrency AS IMPORTO_UNITARIO, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE, " & _
                "       NEW adCurrency AS IMPORTO_SCONTATO, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE_SCONTATO, " & _
                "       NEW adCurrency AS TOTALE_TICKET, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_AGGIUNTIVA, " & _
                "       NEW adCurrency AS TOTALE_QUOTA_NAZIONALE, " & _
                "       NEW adCurrency AS IMPORTO_NETTO"
      
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    rsDataset.Open "SELECT DISTINCT PR.CODICE_PRESTAZIONE, CODICE FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN NOMENCLATORE_TARIFFARIO N ON N.KEY=PR.CODICE_PRESTAZIONE) WHERE MESE=" & cboMese.ListIndex + 1 & " AND ANNO=" & cboAnno.Text & " AND NOT FLAG=3", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
   'INIZIO BLOCCO DETTAGLIO - QUESTO BLOCCO VA RIPETUTO PER TUTTE LE LINEE DI DETTAGLIO

    FE_NumLinea = 1
    
    Do While Not rsDataset.EOF
    
        'crea 2° elemento e nodi punto 2.2.1
        Set nodo2 = doc.createElement("DettaglioLinee")
        nodo2.appendChild CreaNodo("NumeroLinea", Str(FE_NumLinea))
       'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
        nodo1.appendChild nodo2

        With rsMain
            .AddNew
            .Fields("CODICE_PRESTAZIONE") = rsDataset("CODICE")
            FE_Descrizione = rsDataset("CODICE")
            rsAppo.Open "SELECT  SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_ASL") = 0
            Else
                .Fields("TOTALE_ASL") = rsAppo("TOTALEQ")

            End If
            rsAppo.Close
            totaleAsl = totaleAsl + .Fields("TOTALE_ASL")
        
            rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND P.CODICE_ASL IN (SELECT KEY FROM ASL WHERE CODICE_REGIONE=16) AND NOT P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_REGIONE") = 0
            Else
                .Fields("TOTALE_REGIONE") = rsAppo("TOTALEQ")
            End If
            rsAppo.Close
            totaleRegione = totaleRegione + .Fields("TOTALE_REGIONE")
            
            rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ FROM ((RICETTE R INNER JOIN PAZIENTI P ON P.KEY=R.CODICE_PAZIENTE) INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT P.CODICE_ASL IN (SELECT KEY FROM ASL WHERE CODICE_REGIONE=16) AND NOT P.CODICE_ASL=" & structIntestazione.sCodiceAsl & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If IsNull(rsAppo("TOTALEQ")) Then
                .Fields("TOTALE_FUORI_REGIONE") = 0
            Else
                .Fields("TOTALE_FUORI_REGIONE") = rsAppo("TOTALEQ")
            End If
            rsAppo.Close
            totaleFuoriRegione = totaleFuoriRegione + .Fields("TOTALE_FUORI_REGIONE")
            
            'Somma le dialisi x il codice prestazione
            FE_Quantita = totaleAsl + totaleRegione + totaleFuoriRegione
        
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND T.ESENZIONE_QUOTA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            totaleRicette = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE E ON E.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND (CODICE_ESENZIONE=-1 OR CODICE='E05')", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffTicket = rsAppo("TOTALE_R")
            coeffQuotaNazionale = rsAppo("TOTALE_R")
            rsAppo.Close
            
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND CODICE_ESENZIONE=-1", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = rsAppo("TOTALE_R")
            rsAppo.Close
            rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALE_R FROM ((RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND (NOT FLAG=3) AND (NOT CODICE_ESENZIONE=-1) AND T.ESENZIONE_QUOTA=FALSE  AND ESENZIONE_DOPPIA=FALSE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            coeffQuota = coeffQuota + rsAppo("TOTALE_R") / 2
            rsAppo.Close
            
            rsAppo.Open "SELECT DISTINCT IMPORTO, IMPORTO_SCONTATO FROM (RICETTE R INNER JOIN PRESCRIZIONI PR ON PR.CODICE_RICETTA=R.KEY) WHERE CODICE_PRESTAZIONE=" & rsDataset("CODICE_PRESTAZIONE") & " AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND NOT FLAG=3", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            .Fields("IMPORTO_UNITARIO") = rsAppo("IMPORTO")
            .Fields("IMPORTO_TOTALE") = .Fields("IMPORTO_UNITARIO") * (.Fields("TOTALE_ASL") + .Fields("TOTALE_REGIONE") + .Fields("TOTALE_FUORI_REGIONE"))
            
            'Riporta il costo ed il totale della specifica prestazione
            FE_PrezzoUnit = rsAppo("IMPORTO")
            FE_PrezzoTot = .Fields("IMPORTO_UNITARIO") * (.Fields("TOTALE_ASL") + .Fields("TOTALE_REGIONE") + .Fields("TOTALE_FUORI_REGIONE"))
            
        'aggiunge al 2° elemento i nodi
    nodo2.appendChild CreaNodo("Descrizione", FE_Descrizione)
    nodo2.appendChild CreaNodo("Quantita", VirgolaOrPunto(Format(FE_Quantita, "#####.00"), ","))
    nodo2.appendChild CreaNodo("UnitaMisura", "NR")
    nodo2.appendChild CreaNodo("DataInizioPeriodo", sistemaData("01-" & cboMese.ListIndex + 1 & "-" & cboAnno.Text))
    nodo2.appendChild CreaNodo("DataFinePeriodo", sistemaData(GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)))
    nodo2.appendChild CreaNodo("PrezzoUnitario", VirgolaOrPunto(Format(FE_PrezzoUnit, "#####.00"), ","))
    nodo2.appendChild CreaNodo("PrezzoTotale", VirgolaOrPunto(Format(FE_PrezzoTot, "#####.00"), ","))
    nodo2.appendChild CreaNodo("AliquotaIVA", "0.00")
    nodo2.appendChild CreaNodo("Natura", "N4")
        
            
            importoTotale = importoTotale + .Fields("IMPORTO_TOTALE")
            .Fields("IMPORTO_SCONTATO") = rsAppo("IMPORTO_SCONTATO")
            .Fields("IMPORTO_TOTALE_SCONTATO") = .Fields("IMPORTO_SCONTATO") * (.Fields("TOTALE_ASL") + .Fields("TOTALE_REGIONE") + .Fields("TOTALE_FUORI_REGIONE"))
            
            .Fields("TOTALE_QUOTA_AGGIUNTIVA") = quotaAggiuntiva * coeffQuota
            .Fields("TOTALE_QUOTA_NAZIONALE") = quotaNazionale * coeffQuotaNazionale
            .Fields("TOTALE_TICKET") = ticket * coeffTicket
            .Fields("IMPORTO_NETTO") = .Fields("IMPORTO_TOTALE_SCONTATO") - .Fields("TOTALE_TICKET") - .Fields("TOTALE_QUOTA_AGGIUNTIVA") - .Fields("TOTALE_QUOTA_NAZIONALE")
            
            importoTotaleQuotaAggiuntiva = importoTotaleQuotaAggiuntiva + .Fields("TOTALE_QUOTA_AGGIUNTIVA")
            importoTotaleQuotaNazionale = importoTotaleQuotaNazionale + .Fields("TOTALE_QUOTA_NAZIONALE")
            importoTotaleTicket = importoTotaleTicket + .Fields("TOTALE_TICKET")
            importoTotaleNetto = importoTotaleNetto + .Fields("IMPORTO_NETTO")
            importoTotaleScontato = importoTotaleScontato + .Fields("IMPORTO_TOTALE_SCONTATO")
            totaleRicette = 0
            totaleAsl = 0
            totaleRegione = 0
            totaleFuoriRegione = 0
            rsAppo.Close
            .Update
            FE_NumLinea = FE_NumLinea + 1
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close

    Set rsDataset = Nothing
    Set rsAppo = Nothing
    Set rsMain = Nothing
    
'FINE BLOCCO DETTAGLIO
    
    'crea 2° elemento e nodi punto 2.2.2
    Set nodo2 = doc.createElement("DatiRiepilogo")
        nodo2.appendChild CreaNodo("AliquotaIVA", "0.00")
        nodo2.appendChild CreaNodo("Natura", "N4")
        nodo2.appendChild CreaNodo("ImponibileImporto", VirgolaOrPunto(Format(importoTotale, "#####.00"), ","))
        nodo2.appendChild CreaNodo("Imposta", "0.00")
       'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
        nodo1.appendChild nodo2
'..................................................
    'DATI PAGAMENTO punto 2.4
    
    rsDataset.Open "SELECT INTESTATARIO_CC,CODICE_ASL,IBAN FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    'crea 1° elemento
    Set nodo1 = doc.createElement("DatiPagamento")
        'aggiunge al 2° macroblocco il 1° elemento
        nodo0.appendChild nodo1

    'crea nodo punto 2.4.1
    nodo1.appendChild CreaNodo("CondizioniPagamento", "TP02")
        
    'crea 2° elemento e nodi punto 2.4.2
    Set nodo2 = doc.createElement("DettaglioPagamento")
        
    'aggiunge al 2° elemento i nodi
    nodo2.appendChild CreaNodo("Beneficiario", rsDataset("INTESTATARIO_CC"))
    nodo2.appendChild CreaNodo("ModalitaPagamento", "MP05")
    nodo2.appendChild CreaNodo("ImportoPagamento", VirgolaOrPunto(Format(importoTotale, "#####.00"), ","))
    nodo2.appendChild CreaNodo("IBAN", rsDataset("IBAN"))
    'aggiunge al 1° elemento l'elemento e i nodi del 2° elemento
    nodo1.appendChild nodo2
    
    root.appendChild nodo0
    rsDataset.Close
    
    ' Salva la fattura XML
    NameXML = "\IT" & MCodFisc & "_" & Format(MProgr_Invio, "00000")
    NameExtXML = NameXML & ".xml"
    doc.Save structApri.pathExe & "\FE\" & NameXML & ".xml"
    
    ' Salva le info della fattura nella tabella FE
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    v_Nomi = Array("KEY", "N_FATTURA", "TIPO_DOC", "PROGR_INVIO", "DATA_INVIO", "NOME_FILE")
    v_Val = Array(GetNumero("FE"), txtNumFattura, "Fattura", MProgr_Invio, sistemaData(GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)), Mid(NameXML, 2, Len(NameXML) - 1))
    Set rsFE = New Recordset
    rsFE.Open "FE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsFE.AddNew v_Nomi, v_Val
    rsFE.Update
    rsFE.Close
    Set rsFE = Nothing
            
    ' Copia il file XML dalla cartella FE al percorso selezionato
    ' Call FileCopyEx(structApri.pathExe & "\*.xsl", txtPercorso & "\*.xsl")
    Call FileCopyEx(structApri.pathExe & "\FE\" & NameExtXML, txtPercorso & "\" & NameExtXML)
    
    'Visualizza nel browser la fattura dal file XML della cartella FE
    'SHOW_SHOWNORMAL = 1
    'SHOW_SHOWMAXIMIZED = 3
    ret = ShellExecute(Me.hWnd, "open", structApri.pathExe & "\FE\" & NameExtXML, vbNullString, vbNullString, 1)
    If ret < 32 Then MsgBox "Si è verificato un errore aprendo il browser di default", vbCritical, "ATTENZIONE!!!"

End Sub

Private Sub VisualizzaFE_Click()
    frmVisualizzaFattureElettroniche.Show 1
End Sub
