VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScanDocumenti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scansione documenti paziente"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmScanDocumenti.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         Left            =   11160
         TabIndex        =   13
         Top             =   360
         Width           =   615
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
         TabIndex        =   11
         Top             =   360
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
         Index           =   1
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Width           =   630
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
         Left            =   10560
         TabIndex        =   9
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7695
      Begin VB.TextBox txtAppo 
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   35
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   3120
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   99
         FormatString    =   "| Nome documento                                                                                     | Data                "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmScanDocumenti.frx":0459
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   7800
      TabIndex        =   6
      Top             =   720
      Width           =   4455
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Documenti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   1280
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
         Height          =   615
         Left            =   1560
         TabIndex        =   16
         Top             =   1280
         Width           =   1335
      End
      Begin VB.CommandButton cmdVisualizza 
         Caption         =   "&Visualizza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1335
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
         Height          =   615
         Left            =   3000
         TabIndex        =   2
         Top             =   360
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
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   1280
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserisci 
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
         Height          =   615
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlApri 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   12120
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmScanDocumenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vRow As Integer
Dim vCol As Integer
Dim rsScan As Recordset
Dim intPazientiKey As Integer

Private Aggiorna As Boolean

Public Property Let LetAggiorna(ByVal vaggiorna As Boolean)
    Aggiorna = vaggiorna
    Call CaricaFlx
End Property

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 2
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
        .MousePointer = flexCustom
    End With
    flxGriglia.ColAlignment(1) = vbLeftJustify
    flxGriglia.Rows = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'------------------------


Private Sub CaricaFlx()
    Dim v_appo() As String
    
    flxGriglia.Rows = 1
    Set rsScan = New Recordset
    rsScan.Open "SELECT * FROM SCAN_DOCUMENTI_PAZIENTI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsScan.EOF And rsScan.BOF) Then
        With flxGriglia
            Do While Not rsScan.EOF
                If Right(rsScan("NOME_FILE"), 2) = "01" Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsScan("KEY")
                    v_appo = Split(rsScan("NOME_FILE"), " ", 2)
                    .TextMatrix(.Rows - 1, 1) = Left(v_appo(1), Len(v_appo(1)) - 2)
                    .TextMatrix(.Rows - 1, 2) = rsScan("DATA")
                End If
                rsScan.MoveNext
            Loop
        End With
    End If
    Set rsScan = Nothing
    flxGriglia.Row = 0
End Sub

Private Sub SalvaModifiche(vNome As String)
    Dim valore As Variant
    Dim nome As Variant
    Dim num As Integer
    
    nome = "NOME_FILE"
    valore = S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(vRow, 1)
    
    Set rsScan = New Recordset
    rsScan.Open "SELECT * FROM SCAN_DOCUMENTI_PAZIENTI WHERE CODICE_SCHEDA=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    Do While Not rsScan.EOF
        num = Right(rsScan("NOME_FILE"), 2)
        rsScan.Update nome, valore & num
        If Dir(structApri.pathDB & "\" & S_DP & intPazientiKey & " " & vNome & num & ".jpg") <> "" Then
            Name structApri.pathDB & "\" & S_DP & intPazientiKey & " " & vNome & num & ".jpg" As structApri.pathDB & "\" & valore & num & ".jpg"
        ElseIf Dir(structApri.pathDB & "\" & S_DP & intPazientiKey & " " & vNome & num & ".pdf") <> "" Then
            Name structApri.pathDB & "\" & S_DP & intPazientiKey & " " & vNome & num & ".pdf" As structApri.pathDB & "\" & valore & num & ".pdf"
        End If
        rsScan.MoveNext
    Loop
    rsScan.Close
    Set rsScan = Nothing
End Sub

Private Sub Visualizza(riga As Integer)
    Dim nome As String
    Dim Visualizza As Boolean
    Dim keyScan As Integer
    Dim numPag As Integer
    
    Set rsScan = New Recordset
    rsScan.Open "SELECT * FROM SCAN_DOCUMENTI_PAZIENTI WHERE CODICE_SCHEDA=" & flxGriglia.TextMatrix(riga, 0), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsScan.EOF And rsScan.BOF) Then
        nome = rsScan("NOME_FILE")
        keyScan = rsScan("KEY")
        numPag = rsScan.RecordCount
        Visualizza = True
    Else
        Visualizza = False
    End If
    rsScan.Close
    
    If Visualizza Then
        Unload frmVisualizzaScansione
        Load frmVisualizzaScansione
        frmVisualizzaScansione.LetNomeFile = nome
        frmVisualizzaScansione.letcodiceRecord = flxGriglia.TextMatrix(riga, 0)
        frmVisualizzaScansione.letNumPag = numPag
        frmVisualizzaScansione.Show 1
    Else
        MsgBox "Nessun documento da visualizzare", vbCritical, "Attenzione"
    End If
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    
    If intPazientiKey = 0 Then
        ' pulisce la griglia
        ' pulisce la flx azzerando le righe
        flxGriglia.Rows = 1
    Else
        ' carica i dati del paziente
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
        ' cerca i riferimenti al paziente
        Call CaricaFlx
    End If
End Sub

Private Sub cmdVisualizza_Click()
    If flxGriglia.Row <> 0 And intPazientiKey <> 0 Then
        If Dir(structApri.pathDB & "\" & S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(vRow, 1) & "01.jpg") <> "" Then
            Call Visualizza(vRow)
        ElseIf Dir(structApri.pathDB & "\" & S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(vRow, 1) & "01.pdf") <> "" Then
            ShellExecute Me.hWnd, "open", structApri.pathDB & "\" & S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(vRow, 1) & "01.pdf", "", "", 5
        Else
            MsgBox "File non trovato", vbInformation, "Visualizza referto"
        End If
    Else
        MsgBox "Selezionare il documento da visualizzare", vbCritical, "Attenzione"
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdElimina_Click()
    Dim nomeFile As String
    
    If intPazientiKey = 0 Then Exit Sub
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare il documento da eliminare", vbCritical, "Attenzione"
    Else
        If MsgBox("Sicuri di eliminare questo documento?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminazione") = vbYes Then
            Set rsScan = New Recordset
            rsScan.Open "SELECT * FROM SCAN_DOCUMENTI_PAZIENTI WHERE CODICE_SCHEDA=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            Do While Not rsScan.EOF
                nomeFile = rsScan("NOME_FILE")
                rsScan.Delete
                ' elimina anche il file
                If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                    Kill structApri.pathDB & "\" & nomeFile & ".jpg"
                ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                    Kill structApri.pathDB & "\" & nomeFile & ".pdf"
                End If
                rsScan.MoveNext
            Loop
            rsScan.Close
            Set rsScan = Nothing
            
            If flxGriglia.Rows = 2 Then
                flxGriglia.Rows = 1
            Else
                flxGriglia.RemoveItem (vRow)
            End If
        End If
    End If
End Sub

Private Sub cmdGestioneReferti_Click()
    If flxGriglia.Row <> 0 Then
        If intPazientiKey <> 0 Then
            Unload frmGestioneDocumentiEsterni
            Load frmGestioneDocumentiEsterni
            frmGestioneDocumentiEsterni.LetCodicePaziente = intPazientiKey
            frmGestioneDocumentiEsterni.letcodiceRecord = flxGriglia.TextMatrix(flxGriglia.Row, 0)
            frmGestioneDocumentiEsterni.LetNomeFile = S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(flxGriglia.Row, 1)
            tDocumentiEsterni = tpSCANDOCPAZIENTI
            frmGestioneDocumentiEsterni.Show 1
        Else
            MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        End If
    Else
        MsgBox "Selezionare il documento", vbCritical, "Attenzione"
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    Dim i As Integer
    
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then Exit Sub
    If flxGriglia.Rows <= 1 Then Exit Sub
     
    Set rsScan = New Recordset
    rsScan.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsScan("COGNOME") & " " & rsScan("NOME")
    structIntestazione.sDataPaziente = rsScan("DATA_NASCITA")
    codiceId = rsScan("CODICE_ID")
    Set rsScan = Nothing

    
    strSql = "SHAPE APPEND  NEW adVarChar (40) as NOME_DOCUMENTO, " & _
                    "       NEW adVarChar (10) as DATA_DOCUMENTO "
                                              
                                              
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
       
    With rsMain
        For i = 1 To flxGriglia.Rows - 1
            .AddNew
            .Fields("NOME_DOCUMENTO") = flxGriglia.TextMatrix(i, 1)
            .Fields("DATA_DOCUMENTO") = flxGriglia.TextMatrix(i, 2)
        Next i
    End With

    Set rptScansioneDocumentiPazienti.DataSource = rsMain
    rptScansioneDocumentiPazienti.TopMargin = 0
    rptScansioneDocumentiPazienti.BottomMargin = 0
    rptScansioneDocumentiPazienti.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptScansioneDocumentiPazienti.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptScansioneDocumentiPazienti.PrintReport True, rptRangeAllPages
    
    Set rsScan = Nothing
End Sub

Private Sub cmdInserisci_Click()
    Dim nome As String
    Dim keyId As Integer
    Dim nomeFile As String
    Dim caratteriSpeciali As Boolean
    
    nomeFile = "Nuovo documento"
    If intPazientiKey <> 0 Then
        Do
            caratteriSpeciali = False
            nomeFile = InputBox("Inserire il nome del documento", "Scansione documento", nomeFile)
            If nomeFile = "" Then
                Exit Sub
            End If
            If InStr(1, nomeFile, "\") Or InStr(1, nomeFile, "/") Or InStr(1, nomeFile, "?") Or InStr(1, nomeFile, "*") Or InStr(1, nomeFile, ":") Or InStr(1, nomeFile, "''") Or InStr(1, nomeFile, "|") Or InStr(1, nomeFile, ">") Or InStr(1, nomeFile, "<") Then
                MsgBox "Il nome del file non può contenere i seguenti caratteri  \/*:?''|<>", vbCritical, "Attenzione"
                caratteriSpeciali = True
            End If
            If Esiste(flxGriglia, 1, vRow, nomeFile) Then
                MsgBox "Il nome inserito è già presente", vbCritical, "Attenzione"
            End If
            If Len(nomeFile) > 50 Then
                MsgBox "Il nome non può essere piu lungo di 50 caratteri", vbCritical, "Attenzione"
            End If
        Loop While Esiste(flxGriglia, 1, vRow, nomeFile) Or Len(nomeFile) > 50 Or caratteriSpeciali = True

        nome = S_DP & intPazientiKey & " " & nomeFile & "01"
        keyId = GetNumero("SCAN_DOCUMENTI_PAZIENTI")
        Set rsScan = New Recordset
        rsScan.Open "SCAN_DOCUMENTI_PAZIENTI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsScan.AddNew
        rsScan("KEY") = keyId
        rsScan("CODICE_PAZIENTE") = intPazientiKey
        rsScan("NOME_FILE") = nome
        rsScan("DATA") = date
        rsScan("CODICE_SCHEDA") = keyId
        rsScan.Update
        Set rsScan = Nothing
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = keyId
            .TextMatrix(.Rows - 1, 1) = nomeFile
            .TextMatrix(.Rows - 1, 2) = date
            .Row = .Rows - 1
        End With
        Unload frmGestioneDocumentiEsterni
        Load frmGestioneDocumentiEsterni
        frmGestioneDocumentiEsterni.LetCodicePaziente = intPazientiKey
        frmGestioneDocumentiEsterni.letcodiceRecord = 0
        frmGestioneDocumentiEsterni.LetNomeFile = S_DP & intPazientiKey & " " & flxGriglia.TextMatrix(flxGriglia.Row, 1)
        tDocumentiEsterni = tpSCANDOCPAZIENTI
        frmGestioneDocumentiEsterni.Show 1

    End If
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    flxGriglia.Rows = 1
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        If intPazientiKey = tTrova.keyReturn Then
            intPazientiKey = 0
            Call CaricaPaziente
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        vRow = flxGriglia.Row
        vCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        .SetFocus
        If .Col = 1 Then
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        End If
    End With
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim vecchioNome As String
    txtAppo.Visible = False
    If UCase(flxGriglia.TextMatrix(vRow, 1)) <> UCase(txtAppo) Then
        If Trim(txtAppo) = "" Then
            MsgBox "Impossibile memorizzare nomi vuoti", vbCritical, "Attenzione"
            flxGriglia.Row = vRow
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Sub
        End If
        If InStr(1, txtAppo, "\") Or InStr(1, txtAppo, "/") Or InStr(1, txtAppo, "?") Or InStr(1, txtAppo, "*") Or InStr(1, txtAppo, ":") Or InStr(1, txtAppo, "''") Or InStr(1, txtAppo, "|") Or InStr(1, txtAppo, ">") Or InStr(1, txtAppo, "<") Then
            MsgBox "Il nome del file non può contenere i seguenti caratteri  \/*:?''|<>", vbCritical, "Attenzione"
            flxGriglia.Row = vRow
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            txtAppo.Visible = True
            txtAppo.SetFocus
            Exit Sub
        End If
        If Esiste(flxGriglia, 1, vRow, txtAppo) Then
            MsgBox "Il nome inserito è già presente", vbCritical, "Attenzione"
        Else
            vecchioNome = flxGriglia.TextMatrix(vRow, 1)
            flxGriglia.TextMatrix(vRow, 1) = UCase(txtAppo.Text)
            Call SalvaModifiche(vecchioNome)
        End If
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

