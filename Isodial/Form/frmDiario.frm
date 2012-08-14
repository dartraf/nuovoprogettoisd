VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmDiario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DIARIO CLINICO"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmDiario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   450
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
         Left            =   2160
         TabIndex        =   26
         Top             =   360
         Width           =   3255
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
         Left            =   6720
         TabIndex        =   25
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblEta 
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
         Left            =   11040
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anni"
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
         Left            =   10440
         TabIndex        =   17
         Top             =   360
         Width           =   465
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
         Index           =   1
         Left            =   6000
         TabIndex        =   16
         Top             =   360
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
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   12015
      Begin VB.CheckBox chkFiltra 
         Height          =   270
         Left            =   2040
         Picture         =   "frmDiario.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Filtra titoli diario effettuati"
         Top             =   290
         Width           =   375
      End
      Begin VB.CheckBox chkStampa 
         Caption         =   "Stampa Cartella Clinica"
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
         Left            =   7920
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cboTitolo 
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
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   285
         Width           =   8175
      End
      Begin DataTimeBox.uDataTimeBox oDataTimeBox 
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   760
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
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
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titolo Diario"
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
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   12015
      Begin VB.TextBox txtDati 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   11655
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   12015
      Begin VB.Label lblNomeUtente 
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
         Left            =   8400
         TabIndex        =   23
         Top             =   430
         Width           =   3375
      End
      Begin VB.Label lblCognomeUtente 
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
         Left            =   3600
         TabIndex        =   22
         Top             =   430
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultima modifica effettuata da"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   39
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1680
         WordWrap        =   -1  'True
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
         Index           =   40
         Left            =   2400
         TabIndex        =   20
         Top             =   480
         Width           =   1005
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
         Index           =   41
         Left            =   7560
         TabIndex        =   19
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   12015
      Begin VB.CommandButton cmdElimina 
         Caption         =   "&Elimina"
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
         Left            =   6960
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdStampa 
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
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   6
         Top             =   240
         Width           =   1575
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
         Left            =   10560
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmDiario.frm
'
' <b>Descrizione</b>: Scheda Diario Clinico associata alla tab DIARI_CLINICI
'
' @remarks
'
' @author
'
' @date 04/02/2011 19.52
Option Explicit

'' indica se si è in fase di modifica
Dim modifica As Boolean
'' rs associato alla schefa
Dim rsDiario As Recordset
'' key del record caricato
Dim keyId As Integer
'' rs della tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim blnModificato As Boolean

'' Ricarica le cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("TITOLI_DIARIO", "NOME", cboTitolo)
        
    Select Case CaricaPazienteInAperturaForm(Me.Caption, blnModificato, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
    End Select
    
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    modifica = False

    oDataTimeBox.EnableElenca (False)
    oDataTimeBox.VisibleElenca = True
    Call ApriRsDisconnesso
    oDataTimeBox.ConnectionString = strConnectionStringCentro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "DIARI_CLINICI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Set rsDisco = New ADODB.Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsDisco.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsDisco.CursorLocation = adUseClient
    rsDisco.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

'' Confronta i campi per rilevare le eventuali modifiche
' e le salva nella relativa tabella delle modifiche
'
' @param rs rs che contiene lo stato del record che si è memorizzato
Private Sub Confronta(rs As Recordset)
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim v_modifiche() As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nome_campi As String
    Dim valori As String
    Dim trovato As Boolean
    
    ReDim v_modifiche(0)
    For i = 0 To rsDisco.Fields.count - 1
        trovato = False
        If IsNull(rsDisco(i)) Or IsNull(rs(i)) Then
            If Not (IsNull(rsDisco(i)) And IsNull(rs(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(i) <> rs(i) Then
                trovato = True
            End If
        End If
        If trovato Then
            ReDim Preserve v_modifiche(UBound(v_modifiche) + 1)
            v_modifiche(UBound(v_modifiche)) = i
        End If
    Next i
    If UBound(v_modifiche) <> 0 Then
        For i = 1 To UBound(v_modifiche)
            nome_campi = nome_campi & rsDisco.Fields((v_modifiche(i))).Name & "&-&"
            valori = valori & IIf(IsNull(rsDisco.Fields((v_modifiche(i)))) Or rsDisco.Fields((v_modifiche(i))) = "", "NULL", rsDisco.Fields((v_modifiche(i)))) & "&-&"
            ' aggiorna il rsDisco
            rsDisco(v_modifiche(i)) = rs(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "DATA_RECORD", "CODICE_TITOLO", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), oDataTimeBox.data, cboTitolo.ItemData(cboTitolo.ListIndex), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_DIARI_CLINICI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Carica i dati dell'ultimo utente modificatore
Private Sub CaricaUtenteModificatore(key As Integer)
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT COGNOME, NOME FROM LOGIN WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        lblCognomeUtente = rsDataset("COGNOME")
        lblNomeUtente = rsDataset("NOME")
    Else
        lblCognomeUtente = ""
        lblNomeUtente = ""
    End If
    rsDataset.Close
End Sub

'' Carica la scheda nel form e nel rsDisco
Private Sub CaricaScheda()
    Dim i As Integer
    Dim strSql As String
    
    If intPazientiKey <> 0 And oDataTimeBox.data <> "" And cboTitolo.ListIndex <> -1 Then
        Call Pulisci
        
        strSql = "SELECT    DIARI_CLINICI.* " & _
                "FROM       (DIARI_CLINICI " & _
                "           INNER JOIN TITOLI_DIARIO ON TITOLI_DIARIO.KEY=DIARI_CLINICI.CODICE_TITOLO) " & _
                "WHERE      (CODICE_PAZIENTE=" & intPazientiKey & ") AND " & _
                "           (DATA=#" & oDataTimeBox.DataAmericana & "#) AND " & _
                "           (TITOLI_DIARIO.KEY=" & cboTitolo.ItemData(cboTitolo.ListIndex) & ")"
        Set rsDiario = New Recordset
        rsDiario.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDiario.EOF And rsDiario.BOF) Then
            txtDati = rsDiario("DATI")
            chkStampa.Value = IIf(CBool(rsDiario("STAMPA")), Checked, Unchecked)
            keyId = rsDiario("KEY")
            Call CaricaUtenteModificatore(rsDiario("UTENTE_MODIFICATORE"))
            
            ' aggiorna i dati nel rsDisco
            Do While Not rsDisco.EOF
                rsDisco.Delete
                rsDisco.MoveNext
            Loop
            rsDisco.AddNew
            For i = 0 To rsDisco.Fields.count - 1
                rsDisco.Fields(i) = rsDiario.Fields(i)
            Next i
            rsDisco.Update
            
            modifica = True
        Else
            modifica = False
        End If
        Set rsDiario = Nothing
    End If
    blnModificato = False
End Sub

'' Verifica prima di memorizzare che tutti i dati siano inseriti
Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If oDataTimeBox.data = "" Then
        MsgBox "La data inserita non è corretta", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboTitolo.ListIndex = -1 Then
        MsgBox "Selezionare il titolo", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    modifica = False
    intPazientiKey = 0
    oDataTimeBox.Pulisci
    Call PulisciForm(Me)
    Call Pulisci
End Sub

Private Sub Pulisci()
    chkStampa.Value = Checked
    chkFiltra.Value = Unchecked
    txtDati.Text = ""
    lblCognomeUtente.Caption = ""
    lblNomeUtente.Caption = ""
    blnModificato = False
End Sub

'' Filtra solo i titoli diario che il paziente ha effettuato
Private Sub Filtra()
    Dim strSql As String
    If chkFiltra.Value = Checked Then
        strSql = "SELECT    DISTINCT TITOLI_DIARIO.KEY, TITOLI_DIARIO.NOME " & _
                 "FROM      (DIARI_CLINICI " & _
                 "          INNER JOIN TITOLI_DIARIO ON DIARI_CLINICI.CODICE_TITOLO=TITOLI_DIARIO.KEY) " & _
                 "WHERE      CODICE_PAZIENTE=" & intPazientiKey
        Call RicaricaComboBox(strSql, "NOME", cboTitolo)
    Else
        Call RicaricaComboBox("TITOLI_DIARIO", "NOME", cboTitolo)
    End If
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then Exit Sub
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA,CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognome = rsDataset("COGNOME")
    lblNome = rsDataset("NOME")
    Dim somma As Integer
    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
    Set rsDataset = Nothing
    
    Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
    blnModificato = False
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "La scheda deve essere prima memorizzata", vbCritical, "Attenzione"
        Exit Sub
    End If
      
    Set rsDiario = New Recordset
    rsDiario.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDiario("COGNOME") & " " & rsDiario("NOME")
    structIntestazione.sDataPaziente = rsDiario("DATA_NASCITA")
    rsDiario.Close
    Set rsDiario = New Recordset
    
    Call StampaDecimaParte(False, intPazientiKey)
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    
    If Completo Then
        If Not modifica Then
            keyId = GetNumero("DIARI_CLINICI")
        End If
        v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "DATI", "STAMPA", "UTENTE_MODIFICATORE", "CODICE_TITOLO")
        v_Val = Array(keyId, intPazientiKey, oDataTimeBox.data, txtDati, IIf(chkStampa.Value = Checked, True, False), tAccesso.key, cboTitolo.ItemData(cboTitolo.ListIndex))
        
        Set rsDiario = New Recordset
        If modifica Then
            rsDiario.Open "SELECT * FROM DIARI_CLINICI WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDiario.Update v_Nomi, v_Val
            If TRACCIATO Then
                Call Confronta(rsDiario)
            End If
        Else
            rsDiario.Open "DIARI_CLINICI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDiario.AddNew v_Nomi, v_Val
            rsDiario.Update
        End If
        Set rsDiario = Nothing
        
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        blnModificato = False
        modifica = True
    End If
End Sub

Private Sub cmdTrova_Click()
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        ' pulisce per evitare problemi
        Call PulisciTutto
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = ""
        tTrova.condStato = ""
        frmTrova.Show 1
        If tTrova.keyReturn = 0 Then
            Unload Me
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub cmdElimina_Click()
    If intPazientiKey <> 0 Then
        If modifica Then
            If MsgBox("Sei sicuro di voler eliminare la scheda di: " & UCase(lblCognome) & " " & UCase(lblNome) & "?", vbQuestion & vbYesNo, "Eliminazione") = vbYes Then
                Set rsDiario = New Recordset
                rsDiario.Open "SELECT * FROM DIARI_CLINICI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & oDataTimeBox.DataAmericana & "# AND CODICE_TITOLO=" & cboTitolo.ItemData(cboTitolo.ListIndex), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                If Not (rsDiario.BOF And rsDiario.EOF) Then
                    rsDiario.Delete
                End If
                rsDiario.Close
                
                Call PulisciTutto
                MsgBox "Eliminazione effettuata con successo", vbInformation, "Eliminazione"
                Call TrovaPaziente
            End If
        End If
    End If
End Sub

Private Sub chkFiltra_Click()
    If intPazientiKey <> 0 Then
        txtDati = ""
        Call Filtra
    End If
End Sub

'******** Gestione Modificato

Private Sub chkStampa_Click()
    blnModificato = True
End Sub

Private Sub cboTitolo_Click()
    ' puo elencare solo se il titolo è stato selezionato
    If cboTitolo.ListIndex = -1 Then
        oDataTimeBox.EnableElenca (False)
    Else
        oDataTimeBox.EnableElenca (True)
        Call Pulisci
        Call CaricaScheda
    End If
End Sub

Private Sub oDataTimeBox_OnCalendarClick(blnProsegui As Boolean)
    blnProsegui = ControlloChiusuraForm(blnModificato, Me.Caption)
End Sub

Private Sub oDataTimeBox_OnDataChange()
    If IsDate(oDataTimeBox.data) Then
        Call CaricaScheda
    Else
        If oDataTimeBox.data = "" Then Pulisci
    End If
End Sub

Private Sub oDataTimeBox_OnElencaClick()
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        tElenca.Tipo = tpDIARIO
        tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_TITOLO=" & cboTitolo.ItemData(cboTitolo.ListIndex)
        frmElencaDate.Show 1
        If laData <> "" Then oDataTimeBox.data = laData
    End If
End Sub

Private Sub txtDati_GotFocus()
    txtDati.BackColor = colArancione
End Sub

Private Sub txtDati_LostFocus()
    txtDati.BackColor = vbWhite
End Sub

Private Sub txtDati_Change()
    blnModificato = True
End Sub

