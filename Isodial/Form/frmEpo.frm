VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmEpo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dose Eritropoietina per Paziente"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabScheda 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tabella"
      TabPicture(0)   =   "frmEpo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Grafico 2D"
      TabPicture(1)   =   "frmEpo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grafico(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Grafico 3D"
      TabPicture(2)   =   "frmEpo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grafico(1)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
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
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11415
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
            ItemData        =   "frmEpo.frx":0054
            Left            =   960
            List            =   "frmEpo.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   855
         End
         Begin VB.ComboBox cboEPO 
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
            ItemData        =   "frmEpo.frx":0058
            Left            =   3840
            List            =   "frmEpo.frx":006B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   300
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   3495
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   15
            ScrollTrack     =   -1  'True
            FormatString    =   $"frmEpo.frx":0091
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblUnitaMisura 
            AutoSize        =   -1  'True
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
            Left            =   5280
            TabIndex        =   12
            Top             =   360
            Width           =   75
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Eritropoietina"
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
            Left            =   2280
            TabIndex        =   8
            Top             =   345
            Width           =   1410
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
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   345
            Width           =   540
         End
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   5055
         Index           =   0
         Left            =   -74880
         OleObjectBlob   =   "frmEpo.frx":014D
         TabIndex        =   4
         Top             =   480
         Width           =   11175
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Width           =   11415
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
            Left            =   8520
            TabIndex        =   10
            Top             =   240
            Width           =   1215
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
            Left            =   10080
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Pazienti in elenco:"
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
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label lblTotale 
            AutoSize        =   -1  'True
            Caption         =   "0"
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
            Left            =   2330
            TabIndex        =   11
            Top             =   375
            Width           =   135
         End
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   5055
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmEpo.frx":2DD5
         TabIndex        =   9
         Top             =   480
         Width           =   11175
      End
   End
End
Attribute VB_Name = "frmEpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmEpo.frm
'
' <b>Descrizione</b>: Scheda EPO i cui dati sono caricati della schede dialitiche giornaliere
'
' @remarks
'
' @author
'
' @date 05/02/2011 16.43
Option Explicit

'' Evita l'evento click della cbo quando è in fase di caricamento dati
Dim stoCaricando As Boolean
' intervalli
' alfa 1000-8000
' beta 1000-10000
' darbo 10-500

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim k As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft

    With flxGriglia
        .ColWidth(1) = 0
        .Row = 0
        For i = 0 To 14
            .Col = i
            .CellFontBold = True
        Next i
        .Col = 0
    End With
    
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    For k = 0 To 1
        For i = 1 To 12
            grafico(k).Column = 1
            grafico(k).Row = i
            grafico(k).data = 0
            grafico(k).RowLabel = UCase(MonthName(i, True))
        Next i
    Next k
    tabScheda.Tab = 0
    Call CaricaScheda
End Sub

Private Sub Stampa()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim i As Integer
    Dim k As Integer
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adInteger AS CODICE, " & _
                "       NEW adVarChar(50) AS PAZIENTE, " & _
                "       NEW adInteger AS MESE1, " & _
                "       NEW adInteger AS MESE2, " & _
                "       NEW adInteger AS MESE3, " & _
                "       NEW adInteger AS MESE4, " & _
                "       NEW adInteger AS MESE5, " & _
                "       NEW adInteger AS MESE6, " & _
                "       NEW adInteger AS MESE7, " & _
                "       NEW adInteger AS MESE8, " & _
                "       NEW adInteger AS MESE9, " & _
                "       NEW adInteger AS MESE10, " & _
                "       NEW adInteger AS MESE11, " & _
                "       NEW adInteger AS MESE12, " & _
                "       NEW adInteger AS TOTALE "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        
    With rsMain
        For i = 1 To flxGriglia.Rows - 1
            .AddNew
            .Fields("PAZIENTE") = " " & flxGriglia.TextMatrix(i, 0)
            For k = 2 To 13
                .Fields("MESE" & k - 1) = "     " & flxGriglia.TextMatrix(i, k)
            Next k
            .Fields("TOTALE") = flxGriglia.TextMatrix(i, 14)
        Next i
    End With
    
    Set rptEpo.DataSource = rsMain
    rptEpo.Orientation = rptOrientLandscape
    rptEpo.Sections("intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    rptEpo.Sections("intestazione").Controls.Item("lblEpo").Caption = cboEPO.Text & lblUnitaMisura
    rptEpo.PrintReport True, rptRangeAllPages
End Sub

'' Aggiorna il grafico con il paziente selezionato
Private Sub AggiornaGrafico(Index As Integer)
     Dim i As Integer
     Dim max As Long
     With flxGriglia
        If .Row = 0 Then
            For i = 1 To 12
                grafico(Index).Column = 1
                grafico(Index).Row = i
                grafico(Index).data = 0
            Next i
        Else
            ' trova il massimo relativo
            max = .TextMatrix(.Row, 2)
            For i = 3 To 13
                If .TextMatrix(.Row, i) > max Then
                    max = .TextMatrix(.Row, i)
                End If
            Next i
            ' imposta il massimo
            grafico(Index).Plot.Axis(VtChAxisIdY).ValueScale.Maximum = max
            ' disegna tutti i mesi
            For i = 1 To 12
                grafico(Index).Column = 1
                grafico(Index).Row = i
                grafico(Index).data = .TextMatrix(.Row, i + 1)
            Next i
        End If
     End With
End Sub

'' Calcola il totale di EPO
Private Sub CalcolaTotale()
    Dim i As Integer
    Dim k As Integer
    Dim somma As Long
    i = 1
    Do While i <> flxGriglia.Rows
        somma = 0
        For k = 2 To 13
            somma = somma + flxGriglia.TextMatrix(i, k)
        Next k
        If somma = 0 Then
            If flxGriglia.Rows = 2 Then
                flxGriglia.Rows = 1
                If cboEPO.ListIndex <> -1 Then
                    MsgBox "Nessun paziente utilizza l'eritropoietina selezionata", vbInformation, "Calcola eritropoietina"
                End If
            Else
                flxGriglia.RemoveItem (i)
            End If
        Else
            flxGriglia.TextMatrix(i, 14) = somma
            i = i + 1
        End If
    Loop
End Sub

'' Carica i dati dell'EPO nella flx
Private Sub CaricaScheda()
    Dim keyId As Integer        ' key del paziente
    Dim i As Integer
    Dim k As Integer
    Dim somma As Long
    Dim rsDataset As Recordset
    somma = 0
    Call CaricaPazienti
    With flxGriglia
        Set rsDataset = New Recordset
        ' POWER
        rsDataset.Open "SELECT PAZIENTI.KEY, SCHEDE_DIALISI.CODICE_PAZIENTE, SCHEDE_DIALISI.CONFERMA_SOMM, SCHEDE_DIALISI.ERRATA,  Month([DATA]) AS MESE, Year([DATA]) AS ANNO, " & _
                       "SCHEDE_DIALISI.CODICE_STORICO_DIALISI, STORICO_DIALISI_GIORNALIERA.KEY, " & _
                       "STORICO_DIALISI_GIORNALIERA.UI, STORICO_DIALISI_GIORNALIERA.EPO FROM PAZIENTI, " & _
                       "SCHEDE_DIALISI, STORICO_DIALISI_GIORNALIERA " & _
                       "WHERE (((SCHEDE_DIALISI.CODICE_PAZIENTE)=[PAZIENTI].[KEY]) AND " & _
                       "((STORICO_DIALISI_GIORNALIERA.KEY)=[SCHEDE_DIALISI].[CODICE_STORICO_DIALISI]) " & _
                       "AND ((STORICO_DIALISI_GIORNALIERA.EPO)=" & cboEPO.ListIndex & ") AND (Year([DATA])=" & cboAnno.Text & ")  AND ((SCHEDE_DIALISI.CONFERMA_SOMM)=TRUE)) AND ((SCHEDE_DIALISI.ERRATA)=FALSE)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            For i = 1 To flxGriglia.Rows - 1
                keyId = .TextMatrix(i, 1)
                For k = 1 To 12
                    rsDataset.Filter = ("PAZIENTI.KEY=" & keyId & " AND MESE=" & k)
                    Do While Not rsDataset.EOF
                        somma = somma + rsDataset("UI")
                        rsDataset.MoveNext
                    Loop
                    .TextMatrix(i, k + 1) = somma
                    somma = 0
                Next k
            Next i
        Set rsDataset = Nothing
    End With
    Call CalcolaTotale
    flxGriglia.Row = 0
    lblTotale = flxGriglia.Rows - 1
End Sub

'' Carica la lista dei pazienti nella flx
Private Sub CaricaPazienti()
    Dim aggiungi As Boolean
    Dim rsDataset As Recordset
    With flxGriglia
        .Rows = 1
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT KEY, COGNOME, NOME, STATO, STATODATA FROM PAZIENTI ORDER BY COGNOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            Do While Not rsDataset.EOF
                If rsDataset("STATO") <> 1 And rsDataset("STATO") <> 2 Then
                    aggiungi = True
                Else
                    If IsNull(rsDataset("STATODATA")) Then
                        aggiungi = True
                    ElseIf Year(rsDataset("STATODATA")) >= cboAnno.Text Then
                        aggiungi = True
                    Else
                        aggiungi = False
                    End If
                End If
                    
                If aggiungi Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .TextMatrix(.Rows - 1, 0) = UCase(rsDataset("COGNOME") & " " & rsDataset("NOME"))
                    .Col = 0
                    .CellBackColor = RGB(231, 255, 255)
                    .TextMatrix(.Rows - 1, 1) = rsDataset("KEY")
                End If
                rsDataset.MoveNext
            Loop
        End If
        Set rsDataset = Nothing
    End With
End Sub

Private Sub cmdStampa_Click()
    If cboEPO.ListIndex = -1 Then
        MsgBox "Selezionare il tipo di eritropoietina", vbCritical, "Stampa"
    Else
        Call Stampa
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
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
'------------------------------

Private Sub cboAnno_Click()
    If stoCaricando Then Exit Sub
    Call CaricaScheda
End Sub

Private Sub cboEPO_Click()
    If cboEPO.ListIndex = 2 Or cboEPO.ListIndex = 3 Then
        lblUnitaMisura = "Valori espressi in mcg"
    Else
        lblUnitaMisura = "Valori espressi in UI"
    End If
    Call CaricaScheda
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, 14, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, 14)
    End If
End Sub

'' Aggiorna il relativo grafico (2D o 3D)
Private Sub tabScheda_Click(PreviousTab As Integer)
    If tabScheda.Tab = 1 Or tabScheda.Tab = 2 Then
        Call AggiornaGrafico(tabScheda.Tab - 1)
    End If
End Sub
