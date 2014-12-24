VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmDiario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DIARIO CLINICO"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmDiario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Seleziona il paziente"
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   12015
      Begin VB.CheckBox chkStampa 
         Caption         =   "Stampa In Cartella Clinica"
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
         Left            =   3360
         TabIndex        =   4
         Top             =   300
         Width           =   3075
      End
      Begin DataTimeBox.uDataTimeBox oDataTimeBox 
         Height          =   372
         Left            =   1080
         TabIndex        =   3
         Top             =   230
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin VB.Label lblNomeUtente 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   9720
         TabIndex        =   23
         Top             =   445
         Width           =   2172
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo aggiornamento del dr./dr.ssa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Index           =   39
         Left            =   7120
         TabIndex        =   22
         Top             =   165
         Width           =   2490
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCognomeUtente 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   9720
         TabIndex        =   21
         Top             =   205
         Width           =   2172
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
         TabIndex        =   15
         Top             =   260
         Width           =   516
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   12015
      Begin VB.TextBox txtDati 
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
         Height          =   2445
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   210
         Width           =   11655
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   12015
      Begin VB.OptionButton OptStDiario 
         Caption         =   "Stampa &TUTTO"
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
         Left            =   240
         TabIndex        =   7
         Top             =   530
         Width           =   3160
      End
      Begin VB.OptionButton OptStDiario 
         Caption         =   "Stampa &VISTA CORRENTE"
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
         Left            =   240
         TabIndex        =   6
         Top             =   200
         Value           =   -1  'True
         Width           =   3160
      End
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
         Left            =   7920
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
         Left            =   6600
         TabIndex        =   8
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
         Left            =   9220
         TabIndex        =   10
         Top             =   240
         Width           =   1270
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
         Left            =   10570
         TabIndex        =   11
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

Private Sub cmdStampa_Click()
    If Not modifica Then
  '      MsgBox "Memorizzare prima la scheda", vbCritical, "ATTENZIONE"
        Exit Sub
    End If
    
    ' STAMPA TUTTO
    If OptStDiario(1).Value = True Then
        Set rsDiario = New Recordset
        rsDiario.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        structIntestazione.sPaziente = rsDiario("COGNOME") & " " & rsDiario("NOME")
        structIntestazione.sDataPaziente = rsDiario("DATA_NASCITA")
        rsDiario.Close
        Set rsDiario = New Recordset
        Call StampaDecimaParte(False, intPazientiKey)
        Exit Sub
    End If
         
     ' STAMPA QUELLA CORRENTE
     Dim codiceId As Integer
     Set rsDiario = New Recordset
     rsDiario.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
     structIntestazione.sPaziente = rsDiario("COGNOME") & " " & rsDiario("NOME")
     structIntestazione.sDataPaziente = rsDiario("DATA_NASCITA")
     codiceId = rsDiario("CODICE_ID")
     rsDiario.Close
    
     Dim SQLString As String
     Dim cnConn As Connection        ' connessione per lo shape
     Dim rsMain As Recordset         ' recordset padre per lo shape
        
     SQLString = "SHAPE APPEND " & _
                 "   NEW adDate AS DATA, " & _
                 "   NEW adLongVarChar AS DATI "
            
     ' apre la connessione per lo shape
     Set cnConn = New ADODB.Connection
     cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
     Set rsMain = New ADODB.Recordset
     rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
            
     ' carica il recordset padre
     Set rsDiario = New Recordset
     rsDiario.Open "SELECT * FROM DIARI_CLINICI WHERE KEY=" & keyId & "  AND STAMPA=TRUE ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
     If Not (rsDiario.EOF And rsDiario.BOF) Then
        With rsMain
        Do While Not rsDiario.EOF
            .AddNew
            .Fields("DATA") = rsDiario("DATA")
            .Fields("DATI") = rsDiario("DATI") & vbCrLf & vbCrLf & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsDiario("UTENTE_MODIFICATORE"))
            rsDiario.MoveNext
            Loop
        End With
     End If
        
     ' azzero tutto
     Set rptCartellaClinica_10 = Nothing
     Set rptCartellaClinica_10.DataSource = rsMain
     rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
     rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
     rptCartellaClinica_10.PrintReport True, rptRangeAllPages

End Sub

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
            
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
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "DATA_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), oDataTimeBox.data, nome_campi, valori)
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
    Dim data As Date
    Dim i As Integer
    
    If intPazientiKey <> 0 And oDataTimeBox.data <> "" Then
        Call Pulisci
        ' la data americana
        data = Month(oDataTimeBox.data) & "/" & Day(oDataTimeBox.data) & "/" & Year(oDataTimeBox.data)
        Set rsDiario = New Recordset
        rsDiario.Open "SELECT * FROM DIARI_CLINICI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND (DATA=#" & data & "#)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDiario.EOF And rsDiario.BOF) Then
            txtDati = rsDiario("DATI")
            chkStampa.Value = IIf(CBool(rsDiario("STAMPA")), Checked, Unchecked)
            keyId = rsDiario("KEY")
            Call CaricaUtenteModificatore(rsDiario("UTENTE_MODIFICATORE"))
            Call Upd_rsDisco
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
    txtDati.Text = ""
    lblCognomeUtente.Caption = ""
    lblNomeUtente.Caption = ""
    blnModificato = False
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

Private Sub cmdChiudi_Click()
    Unload frmDiario
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    
    If Completo Then
        If Not modifica Then
            keyId = GetNumero("DIARI_CLINICI")
        End If
        v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "DATI", "STAMPA", "UTENTE_MODIFICATORE")
        v_Val = Array(keyId, intPazientiKey, oDataTimeBox.data, txtDati, IIf(chkStampa.Value = Checked, True, False), tAccesso.key)
        
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
            Call Upd_rsDisco
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
                rsDiario.Open "SELECT * FROM DIARI_CLINICI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & oDataTimeBox.DataAmericana & "#", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
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

'******** Gestione Modificato

Private Sub chkStampa_Click()
    blnModificato = True
End Sub

Private Sub oDataTimeBox_OnDataChange()
    If IsDate(oDataTimeBox.data) Then
        Call CaricaScheda
        txtDati.Enabled = True
    Else
        If oDataTimeBox.data = "" Then Pulisci
    End If
End Sub

Private Sub oDataTimeBox_OnDataClick()
    Call Pulisci
    oDataTimeBox.Pulisci
End Sub

Private Sub oDataTimeBox_OnElencaClick()
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        tElenca.Tipo = tpDIARIO
        tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey
        frmElencaDate.Show 1
        If laData <> "" Then oDataTimeBox.data = laData
    Else
        txtDati.Enabled = True
    End If
End Sub

Private Sub txtDati_GotFocus()
 ' se la data non è presente NON abilita a scrivere nel txtdati
   If oDataTimeBox.data = "" Then
       txtDati.Enabled = False
       Exit Sub
   Else
    txtDati.BackColor = colArancione
   End If
End Sub

Private Sub txtDati_LostFocus()
    txtDati.BackColor = vbWhite
End Sub

Private Sub txtDati_Change()
    blnModificato = True
End Sub

Private Sub Upd_rsDisco()
' aggiorna i dati nel rsDisco
    Dim i As Integer
    Do While Not rsDisco.EOF
        rsDisco.Delete
        rsDisco.MoveNext
    Loop
        rsDisco.AddNew
        For i = 0 To rsDisco.Fields.count - 1
            rsDisco.Fields(i) = rsDiario.Fields(i)
        Next i
        rsDisco.Update
End Sub
