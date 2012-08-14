VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRimborsiSpese 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rimborsi Spese"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSchede 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Stampa per Paziente "
      TabPicture(0)   =   "frmRimborsiSpese.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblNome"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCognome"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdTrova"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Stampa per Distretti"
      TabPicture(1)   =   "frmRimborsiSpese.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboDa"
      Tab(1).Control(1)=   "cboA"
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(3)=   "Label1(0)"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   120
         Picture         =   "frmRimborsiSpese.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   450
      End
      Begin VB.ComboBox cboDa 
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
         ItemData        =   "frmRimborsiSpese.frx":0491
         Left            =   -73320
         List            =   "frmRimborsiSpese.frx":0493
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cboA 
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
         ItemData        =   "frmRimborsiSpese.frx":0495
         Left            =   -71520
         List            =   "frmRimborsiSpese.frx":0497
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblCognome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   17
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   16
         Top             =   960
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
         Index           =   6
         Left            =   720
         TabIndex        =   13
         Top             =   930
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
         Index           =   5
         Left            =   720
         TabIndex        =   12
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Distretti:    da"
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
         Left            =   -74880
         TabIndex        =   11
         Top             =   510
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a"
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
         Left            =   -71880
         TabIndex        =   10
         Top             =   510
         Width           =   150
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5775
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
         ItemData        =   "frmRimborsiSpese.frx":0499
         Left            =   4680
         List            =   "frmRimborsiSpese.frx":049B
         Style           =   2  'Dropdown List
         TabIndex        =   14
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
         ItemData        =   "frmRimborsiSpese.frx":049D
         Left            =   840
         List            =   "frmRimborsiSpese.frx":04C5
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   265
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
         Left            =   3960
         TabIndex        =   8
         Top             =   265
         Width           =   540
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
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
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRielabora 
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
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   -120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRimborsiSpese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intPazientiKey As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    Call RicaricaComboBox("SELECT NOME, D.KEY FROM (INTESTAZIONE_STAMPA I INNER JOIN DISTRETTI D ON D.CODICE_ASL=I.CODICE_ASL) ORDER BY NOME", "NOME", cboDa)
    Call RicaricaComboBox("SELECT NOME, D.KEY FROM (INTESTAZIONE_STAMPA I INNER JOIN DISTRETTI D ON D.CODICE_ASL=I.CODICE_ASL) ORDER BY NOME", "NOME", cboA)
    cboDa.ListIndex = 0
    cboA.ListIndex = cboA.ListCount - 1
    
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    cboMese.ListIndex = Month(Now) - 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Sub StampaFoglioViaggioHelios(codPaziente As Integer, traspAmbulanza As Boolean)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset

    Dim nomeAsl As String
    Dim giorni As String
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(30) AS ILLAPAZIENTE, " & _
                "       NEW adVarChar(30) AS AFFETTODA, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adInteger AS NUM_GIORNI, " & _
                "       NEW adVarChar(100) as GIORNI"
                
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    rsDataset.Open "SELECT DISTINCT P.KEY, SESSO, COGNOME, P.NOME, D.NOME FROM ((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE (YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & ") AND ERRATA=FALSE AND TRASPORTO_IN_AMBULANZA=" & IIf(traspAmbulanza, "TRUE", "FALSE") & " AND P.KEY=" & codPaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            .Fields("ILLAPAZIENTE") = IIf(rsDataset("SESSO") = "M", "il paziente", "la paziente")
            .Fields("AFFETTODA") = IIf(rsDataset("SESSO") = "M", "affetto da", "affetta da")
            .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("P.NOME")
             
            rsAppo.Open "SELECT DATA FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & codPaziente & " AND YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & " AND ERRATA=FALSE ORDER BY DATA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            giorni = ""
            Do While Not rsAppo.EOF
                giorni = giorni & Format(Day(rsAppo("DATA")), "00") & " - "
                rsAppo.MoveNext
            Loop
            .Fields("NUM_GIORNI") = rsAppo.RecordCount
            .Fields("GIORNI") = Mid(giorni, 1, Len(giorni) - 3)
            rsAppo.Close

            .Update
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
       Exit Sub
    Else
        rptViaggiHelios.Sections("corpo").Controls("lblFineMese").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptViaggiHelios.Sections("corpo").Controls("lblMeseAnno").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptViaggiHelios.DataSource = rsMain
        rptViaggiHelios.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
End Sub

Private Sub StampaDichiarazioneResponsabilita_Bartoli(codPaziente As Integer)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim nomeAsl As String

        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(60) AS ACCOMPAGNATORE, " & _
                "       NEW adDate AS DATA_NASCITA, " & _
                "       NEW adVarChar(50) as CITTA_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_RESIDENZA, " & _
                "       NEW adVarChar(50) AS INDIRIZZO, " & _
                "       NEW adVarChar(50) AS CODICE_FISCALE, " & _
                "       NEW adVarChar(15) as TIPO_AUTO, " & _
                "       NEW adVarChar(15) as TARGA, " & _
                "       NEW adVarChar(15) as SIGNOR, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adVarChar(6) as NAT, " & _
                "       NEW adVarChar(60) AS CITTA_NASCITA_PAZIENTE, " & _
                "       NEW adDate AS DATA_NASCITA_PAZIENTE, " & _
                "       NEW adVarChar(60) AS COMUNE_RESIDENZA_PAZIENTE "

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close

    rsDataset.Open "SELECT DISTINCT P.KEY, A.COGNOME, A.CODICE_FISCALE, A.NOME, A.DATA_NASCITA, A.CITTA_NASCITA, CITTA, PROV, A.INDIRIZZO, TIPO, TARGA, C.NOME, PROV_RESIDENZA, P.DATA_NASCITA, P.CITTA_NASCITA, P.INDIRIZZO, P.COGNOME, P.NOME, SESSO " & _
                   "FROM (((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN COMUNI C ON C.KEY=P.CODICE_COMUNE_RESIDENZA) LEFT OUTER JOIN ACCOMPAGNATORI A ON A.KEY=P.CODICE_ACCOMPAGNATORE) WHERE (YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & ") AND ERRATA=FALSE AND NOT CODICE_ACCOMPAGNATORE=0 AND TRASPORTO_IN_AMBULANZA=FALSE AND P.KEY=" & codPaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            If IsNull(rsDataset("A.COGNOME")) Then
                Exit Sub
            Else
            .Fields("ACCOMPAGNATORE") = rsDataset("A.COGNOME") & "   " & rsDataset("A.NOME")
            .Fields("DATA_NASCITA") = rsDataset("A.DATA_NASCITA")
            .Fields("CITTA_NASCITA") = rsDataset("A.CITTA_NASCITA")
            .Fields("COMUNE_PROV_RESIDENZA") = rsDataset("CITTA") & " (" & rsDataset("PROV") & ")"
            .Fields("INDIRIZZO") = rsDataset("A.INDIRIZZO")
            .Fields("CODICE_FISCALE") = rsDataset("CODICE_FISCALE")
            .Fields("TIPO_AUTO") = rsDataset("TIPO")
            .Fields("TARGA") = rsDataset("TARGA")
            .Fields("CITTA_NASCITA_PAZIENTE") = rsDataset("P.CITTA_NASCITA")
            .Fields("COMUNE_RESIDENZA_PAZIENTE") = rsDataset("C.NOME") & " (" & rsDataset("PROV_RESIDENZA") & "), " & rsDataset("P.INDIRIZZO")
            .Fields("DATA_NASCITA_PAZIENTE") = rsDataset("P.DATA_NASCITA")
            .Fields("PAZIENTE") = rsDataset("P.COGNOME") & " " & rsDataset("P.NOME")
            .Fields("NAT") = IIf(rsDataset("SESSO") = "M", "nato a", "nata a")
            .Fields("SIGNOR") = IIf(rsDataset("SESSO") = "M", "il signore ", "la signora")
            .Update
            End If
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
       Exit Sub
    Else
        rptDichiarazioneDiResponsabilitàBartoli.Sections("intestazione").Controls("lblAsl").Caption = "SPETT/LE ASL DI " & UCase(nomeAsl)
        rptDichiarazioneDiResponsabilitàBartoli.Sections("corpo").Controls("lblFineMese").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptDichiarazioneDiResponsabilitàBartoli.Sections("corpo").Controls("lblMeseAnno").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptDichiarazioneDiResponsabilitàBartoli.DataSource = rsMain
        rptDichiarazioneDiResponsabilitàBartoli.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
End Sub

Private Sub StampaRimborsoSpeseTrasporto(codPaziente As Integer)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rimborso As Currency
    Dim nomeAsl As String
    Dim giorni As String
    Dim condizioneHelios As String
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(20) AS SOTTOSCRITT, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adVarChar(10) AS NAT, " & _
                "       NEW adDate AS DATA_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_RESIDENZA, " & _
                "       NEW adVarChar(50) AS INDIRIZZO, " & _
                "       NEW adVarChar(50) as COMUNE_RESIDENZA, " & _
                "       NEW adSingle AS KM, " & _
                "       NEW adInteger AS NUM_GIORNI, " & _
                "       NEW adVarChar(100) as GIORNI, " & _
                "       NEW adInteger AS NUM_VIAGGI, " & _
                "       NEW adSingle AS TOTALE_KM, " & _
                "       NEW adCurrency AS TOTALE"

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT RIMBORSO_SPESE_VIAGGIO FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rimborso = VirgolaOrPunto(rsDataset("RIMBORSO_SPESE_VIAGGIO"), ".")
    rsDataset.Close
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Then
        condizioneHelios = " AND NOT D.NOME='22'"
    End If
    
    rsDataset.Open "SELECT DISTINCT P.KEY, SESSO, COGNOME, P.NOME, C.NOME, D.NOME, KM, PROV_RESIDENZA, PROV_NASCITA, CITTA_NASCITA, DATA_NASCITA, CODICE_FISCALE, INDIRIZZO, PROV_RESIDENZA FROM (((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN COMUNI C ON C.KEY=P.CODICE_COMUNE_RESIDENZA) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE (YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & ") AND ERRATA=FALSE AND TRASPORTO_IN_AMBULANZA=FALSE AND P.KEY=" & codPaziente & condizioneHelios, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            .Fields("SOTTOSCRITT") = IIf(rsDataset("SESSO") = "M", "Il sottoscritto", "La sottoscritta")
            .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("P.NOME")
            .Fields("NAT") = IIf(rsDataset("SESSO") = "M", "nato a", "nata a")
            .Fields("DATA_NASCITA") = rsDataset("DATA_NASCITA")
            .Fields("COMUNE_PROV_NASCITA") = rsDataset("CITTA_NASCITA") & " (" & rsDataset("PROV_NASCITA") & ")"
            .Fields("COMUNE_PROV_RESIDENZA") = rsDataset("C.NOME") & " (" & rsDataset("PROV_RESIDENZA") & ")"
            .Fields("INDIRIZZO") = rsDataset("INDIRIZZO")
            .Fields("COMUNE_RESIDENZA") = rsDataset("C.NOME")
            .Fields("KM") = rsDataset("KM")
            
            rsAppo.Open "SELECT DATA FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & codPaziente & " AND YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & " AND ERRATA=FALSE ORDER BY DATA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            giorni = ""
            Do While Not rsAppo.EOF
                giorni = giorni & Format(Day(rsAppo("DATA")), "00") & " - "
                rsAppo.MoveNext
            Loop
            .Fields("NUM_GIORNI") = rsAppo.RecordCount
            .Fields("GIORNI") = Mid(giorni, 1, Len(giorni) - 3)
            rsAppo.Close
            
            .Fields("NUM_VIAGGI") = .Fields("NUM_GIORNI") * 4
            .Fields("TOTALE_KM") = .Fields("KM") * .Fields("NUM_VIAGGI")
            .Fields("TOTALE") = .Fields("NUM_GIORNI") * rimborso
            .Update
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
       Exit Sub
    Else
        rptRimborsoSpeseTrasporto.Sections("intestazione").Controls("lblAsl").Caption = "SPETT/LE ASL DI " & UCase(nomeAsl)
        rptRimborsoSpeseTrasporto.Sections("corpo").Controls("lblFineMese").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptRimborsoSpeseTrasporto.Sections("corpo").Controls("lblMeseAnno").Caption = cboMese.Text & " " & cboAnno.Text
        rptRimborsoSpeseTrasporto.Sections("corpo").Controls("lblRimborso").Caption = rimborso
        Set rptRimborsoSpeseTrasporto.DataSource = rsMain
        rptRimborsoSpeseTrasporto.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
End Sub

Private Sub StampaRimborsoSpeseTrasportoBartoli(codPaziente As Integer)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim nomeAsl As String
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(20) AS SOTTOSCRITT, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adVarChar(10) AS NAT, " & _
                "       NEW adDate AS DATA_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_RESIDENZA, " & _
                "       NEW adVarChar(50) AS INDIRIZZO, " & _
                "       NEW adVarChar(16) AS CODICE_FISCALE, " & _
                "       NEW adVarChar(30) AS TELEFONO, " & _
                "       NEW adInteger AS TOTALE_DIALISI "

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close
    
    rsDataset.Open "SELECT DISTINCT P.KEY, SESSO, COGNOME, CODICE_FISCALE, TELEFONO, P.NOME, C.NOME, PROV_RESIDENZA, PROV_NASCITA, CITTA_NASCITA, DATA_NASCITA, INDIRIZZO, PROV_RESIDENZA FROM ((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN COMUNI C ON C.KEY=P.CODICE_COMUNE_RESIDENZA) WHERE (YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & ") AND ERRATA=FALSE AND TRASPORTO_IN_AMBULANZA=FALSE AND P.KEY=" & codPaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            .Fields("SOTTOSCRITT") = IIf(rsDataset("SESSO") = "M", "Il sottoscritto", "La sottoscritta")
            .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("P.NOME")
            .Fields("NAT") = IIf(rsDataset("SESSO") = "M", "nato a", "nata a")
            .Fields("DATA_NASCITA") = rsDataset("DATA_NASCITA")
            .Fields("COMUNE_PROV_NASCITA") = rsDataset("CITTA_NASCITA") & " (" & rsDataset("PROV_NASCITA") & ")"
            .Fields("COMUNE_PROV_RESIDENZA") = rsDataset("C.NOME") & " (" & rsDataset("PROV_RESIDENZA") & ")"
            .Fields("INDIRIZZO") = rsDataset("INDIRIZZO")
            .Fields("CODICE_FISCALE") = rsDataset("CODICE_FISCALE")
            .Fields("TELEFONO") = rsDataset("TELEFONO")
            
            rsAppo.Open "SELECT COUNT(KEY) AS TOTALE FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & codPaziente & " AND YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & " AND ERRATA=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("TOTALE_DIALISI") = rsAppo("TOTALE")
            rsAppo.Close
            
            .Update
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
       Exit Sub
    Else
        rptRimborsoSpeseTrasportoBartoli.Sections("corpo").Controls("lblFineMese").Caption = GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
        rptRimborsoSpeseTrasportoBartoli.Sections("corpo").Controls("lblMeseAnno").Caption = cboMese.Text & " " & cboAnno.Text
        Set rptRimborsoSpeseTrasportoBartoli.DataSource = rsMain
        rptRimborsoSpeseTrasportoBartoli.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdRielabora_Click()
'    On Error GoTo gestione
    Dim condizione As String
    Dim rsDataset As New Recordset
    If intPazientiKey = 0 And tabSchede.Tab = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If tabSchede.Tab = 1 Then
        condizione = " AND D.NOME>='" & cboDa.Text & "' AND D.NOME<='" & cboA.Text & "' ORDER BY P.COGNOME, P.NOME"
    Else
        condizione = " AND P.KEY=" & intPazientiKey
    End If

    rsDataset.Open "SELECT DISTINCT P.KEY, D.NOME, P.NOME,TRASPORTO_IN_AMBULANZA, P.COGNOME FROM ((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE (YEAR([DATA])=" & cboAnno.Text & " AND MONTH([DATA])=" & cboMese.ListIndex + 1 & ") " & condizione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsDataset.Sort = "D.NOME"
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        If CBool(rsDataset("TRASPORTO_IN_AMBULANZA")) = True And tabSchede.Tab <> 1 Then
            If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Then
                Call StampaRimborsoSpese(rsDataset("KEY"), cboAnno.Text, cboMese.ListIndex + 1, True)
                Call StampaFoglioViaggioHelios(rsDataset("KEY"), True)
            Else
                MsgBox "Nessun rimborso. Trasporto effettuato in ambulanza", vbInformation, "Rimborso spese"
            End If
        Else
            cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
            cdlStampa.CancelError = True
            cdlStampa.ShowPrinter
            If tabSchede.Tab = 1 Then
                Call StartProgressBar(rsDataset.RecordCount, 0, Me)
            End If
            
            Do While Not rsDataset.EOF
                If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Then
                    Call StampaRimborsoSpese(rsDataset("KEY"), cboAnno.Text, cboMese.ListIndex + 1, False)
                    Call StampaRimborsoSpeseTrasporto(rsDataset("KEY"))
                    Call StampaDichiarazioneResponsabilita(rsDataset("KEY"), cboAnno.Text, cboMese.ListIndex + 1)
                    Call StampaFoglioViaggioHelios(rsDataset("KEY"), False)
                ElseIf structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Then
                    Call StampaRimborsoSpeseTrasportoBartoli(rsDataset("KEY"))
                    Call StampaDichiarazioneResponsabilita_Bartoli(rsDataset("KEY"))
                End If
                rsDataset.MoveNext
                frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
            Loop
            
            Call StopProgressBar(Me)
        End If
    Else
        MsgBox "Nessuna seduta dialitica per il mese di " & cboMese.Text, vbInformation, Me.Caption
    End If
    rsDataset.Close
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

''
' Lancia la stampa delle dichiarazioni di responsabilità
'
' @param codPaziente codice del paziente
' @param anno anno per la query
' @param mese mese per la query
Private Sub StampaDichiarazioneResponsabilita(codPaziente As Integer, anno As Integer, mese As Integer)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rimborso As Currency
    Dim nomeAsl As String
    Dim giorni As String
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(60) AS ACCOMPAGNATORE, " & _
                "       NEW adDate AS DATA_NASCITA, " & _
                "       NEW adVarChar(50) as CITTA_NASCITA, " & _
                "       NEW adVarChar(50) as COMUNE_PROV_RESIDENZA, " & _
                "       NEW adVarChar(50) AS INDIRIZZO, " & _
                "       NEW adVarChar(15) as TIPO_AUTO, " & _
                "       NEW adVarChar(15) as TARGA, " & _
                "       NEW adVarChar(15) as SIGNOR, " & _
                "       NEW adVarChar(15) as PATENTE, " & _
                "       NEW adDate  as DATA_RILASCIO, " & _
                "       NEW adVarChar(30) as ENTE_RILASCIO, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adVarChar(150) AS INDIRIZZO_PAZIENTE, " & _
                "       NEW adSingle AS KM, " & _
                "       NEW adInteger AS NUM_GIORNI, " & _
                "       NEW adVarChar(100) as GIORNI, " & _
                "       NEW adCurrency AS TOTALE"

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT RIMBORSO_SPESE_VIAGGIO FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rimborso = VirgolaOrPunto(rsDataset("RIMBORSO_SPESE_VIAGGIO"), ".")
    rsDataset.Close
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close

    rsDataset.Open "SELECT DISTINCT P.KEY, A.COGNOME, A.NOME, A.DATA_NASCITA, A.CITTA_NASCITA, CITTA, PROV, A.INDIRIZZO, PATENTE, ENTE_EMITTENTE, A.DATA_RILASCIO, TIPO, TARGA, C.NOME, PROV_RESIDENZA, P.INDIRIZZO, P.COGNOME, P.NOME, SESSO, KM " & _
                   "FROM (((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN COMUNI C ON C.KEY=P.CODICE_COMUNE_RESIDENZA) LEFT OUTER JOIN ACCOMPAGNATORI A ON A.KEY=P.CODICE_ACCOMPAGNATORE) WHERE (YEAR([DATA])=" & anno & " AND MONTH([DATA])=" & mese & ") AND ERRATA=FALSE AND NOT CODICE_ACCOMPAGNATORE=0 AND TRASPORTO_IN_AMBULANZA=FALSE AND P.KEY=" & codPaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            .Fields("ACCOMPAGNATORE") = rsDataset("A.COGNOME") & "   " & rsDataset("A.NOME")
            .Fields("DATA_NASCITA") = rsDataset("DATA_NASCITA")
            .Fields("CITTA_NASCITA") = rsDataset("CITTA_NASCITA")
            .Fields("COMUNE_PROV_RESIDENZA") = rsDataset("CITTA") & " (" & rsDataset("PROV") & ")"
            .Fields("INDIRIZZO") = rsDataset("A.INDIRIZZO")
            .Fields("PATENTE") = rsDataset("PATENTE")
            .Fields("ENTE_RILASCIO") = rsDataset("ENTE_EMITTENTE")
            .Fields("DATA_RILASCIO") = rsDataset("DATA_RILASCIO")
            .Fields("TIPO_AUTO") = rsDataset("TIPO")
            .Fields("TARGA") = rsDataset("TARGA")
            .Fields("INDIRIZZO_PAZIENTE") = rsDataset("C.NOME") & " (" & rsDataset("PROV_RESIDENZA") & "), " & rsDataset("P.INDIRIZZO")
            .Fields("PAZIENTE") = rsDataset("P.COGNOME") & " " & rsDataset("P.NOME")
            .Fields("SIGNOR") = IIf(rsDataset("SESSO") = "M", "il signore ", "la signora")
            .Fields("KM") = rsDataset("KM")
            
            rsAppo.Open "SELECT DATA FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & codPaziente & " AND YEAR([DATA])=" & anno & " AND MONTH([DATA])=" & mese & " AND ERRATA=FALSE ORDER BY DATA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            giorni = ""
            Do While Not rsAppo.EOF
                giorni = giorni & Format(Day(rsAppo("DATA")), "00") & " - "
                rsAppo.MoveNext
            Loop
            .Fields("NUM_GIORNI") = rsAppo.RecordCount
            .Fields("GIORNI") = Mid(giorni, 1, Len(giorni) - 3)
            rsAppo.Close
            
            .Fields("TOTALE") = .Fields("NUM_GIORNI") * rimborso
            .Update
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
       Exit Sub
    Else
        rptDichiarazioneDiResponsabilità.Sections("intestazione").Controls("lblAsl").Caption = "SPETT/LE ASL DI " & UCase(nomeAsl)
        rptDichiarazioneDiResponsabilità.Sections("corpo").Controls("lblFineMese").Caption = GetUltimoGiorno(mese, anno)
        rptDichiarazioneDiResponsabilità.Sections("corpo").Controls("lblMeseAnno").Caption = MonthName(mese) & " " & anno
        rptDichiarazioneDiResponsabilità.Sections("corpo").Controls("lblRimborso").Caption = rimborso
        Set rptDichiarazioneDiResponsabilità.DataSource = rsMain
        rptDichiarazioneDiResponsabilità.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
End Sub


''
' Lancia la stampa dei rimborsi
'
' @param codPaziente codice del paziente
' @param anno anno per la query
' @param mese mese per la query
Private Sub StampaRimborsoSpese(codPaziente As Integer, anno As Integer, mese As Integer, traspAmbulanza As Boolean)
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rimborso As Currency
    Dim nomeAsl As String
        
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(50) AS ASL_DISTRETTO, " & _
                "       NEW adVarChar(20) AS SOTTOSCRITT, " & _
                "       NEW adVarChar(60) AS PAZIENTE, " & _
                "       NEW adVarChar(10) AS NAT, " & _
                "       NEW adDate AS DATA_NASCITA, " & _
                "       NEW adVarChar(150) AS INDIRIZZO, " & _
                "       NEW adVarChar(30) AS TESSERA_SANITARIA, " & _
                "       NEW adVarChar(16) AS CODICE_FISCALE, " & _
                "       NEW adCurrency AS IMPORTO_TOTALE, " & _
                "       NEW adVarChar(30) AS ASSISTITO, " & _
                "       NEW adInteger AS NUM_GIORNI, " & _
                "       NEW adDate AS DATA_INIZIO, " & _
                "       NEW adDate AS DATA_FINE "

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset
    
    rsDataset.Open "SELECT RIMBORSO_SPESE_VIAGGIO FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rimborso = VirgolaOrPunto(rsDataset("RIMBORSO_SPESE_VIAGGIO"), ".")
    rsDataset.Close
    rsDataset.Open "SELECT NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomeAsl = rsDataset("NOME")
    rsDataset.Close

    rsDataset.Open "SELECT DISTINCT P.KEY, SESSO, COGNOME, P.NOME, C.NOME, D.NOME, DATA_NASCITA, CODICE_FISCALE, INDIRIZZO,TESSERA_SANITARIA, PROV_RESIDENZA FROM (((PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) INNER JOIN COMUNI C ON C.KEY=P.CODICE_COMUNE_RESIDENZA) INNER JOIN DISTRETTI D ON D.KEY=P.CODICE_DISTRETTO) WHERE (YEAR([DATA])=" & anno & " AND MONTH([DATA])=" & mese & ") AND ERRATA=FALSE AND TRASPORTO_IN_AMBULANZA=" & IIf(traspAmbulanza, "TRUE", "FALSE") & " AND P.KEY=" & codPaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            .AddNew
            .Fields("ASL_DISTRETTO") = UCase(nomeAsl) & "- distr. " & rsDataset("D.NOME")
            .Fields("SOTTOSCRITT") = IIf(rsDataset("SESSO") = "M", "Il sottoscritto", "La sottoscritta")
            .Fields("ASSISTITO") = "L'assistito"
            .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("P.NOME")
            .Fields("NAT") = IIf(rsDataset("SESSO") = "M", "nato il", "nata il")
            .Fields("DATA_NASCITA") = rsDataset("DATA_NASCITA")
            .Fields("INDIRIZZO") = rsDataset("C.NOME") & " (" & rsDataset("PROV_RESIDENZA") & ")" & ", " & rsDataset("INDIRIZZO")
            .Fields("TESSERA_SANITARIA") = rsDataset("TESSERA_SANITARIA")
            .Fields("CODICE_FISCALE") = rsDataset("CODICE_FISCALE")
            .Fields("ASSISTITO") = IIf(rsDataset("SESSO") = "M", "L'Assistito", "L'Assistita")
            
            rsAppo.Open "SELECT DATA FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & codPaziente & " AND YEAR([DATA])=" & anno & " AND MONTH([DATA])=" & mese & " AND ERRATA=FALSE ORDER BY DATA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            .Fields("DATA_INIZIO") = rsAppo("DATA")
            rsAppo.MoveLast
            .Fields("DATA_FINE") = rsAppo("DATA")
            .Fields("NUM_GIORNI") = rsAppo.RecordCount
            .Fields("IMPORTO_TOTALE") = rsAppo.RecordCount * rimborso
            rsAppo.Close
            
            .Update
        End With
    End If
    rsDataset.Close
    
    If rsMain.RecordCount = 0 Then
        Exit Sub
    Else
        Set rptRimborsoSpese.DataSource = rsMain
        rptRimborsoSpese.PrintReport False, rptRangeAllPages
    End If
    
    Set rsDataset = Nothing
    Set rsAppo = Nothing
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then Exit Sub
    ' carica i dati del paziente
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognome = rsDataset("COGNOME")
    lblNome = rsDataset("NOME")
    Set rsDataset = Nothing
End Sub

Private Sub cmdTrova_Click()
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    Call CaricaPaziente
End Sub


