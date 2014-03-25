VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmKtv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calcolo Kt/V"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabScheda 
      Height          =   4680
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8255
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
      TabPicture(0)   =   "frmKtv.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Grafico 2D"
      TabPicture(1)   =   "frmKtv.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grafico(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Grafico 3D"
      TabPicture(2)   =   "frmKtv.frx":0038
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
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   12735
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
            ItemData        =   "frmKtv.frx":0054
            Left            =   960
            List            =   "frmKtv.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtAppo 
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
            Left            =   2880
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   720
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   2535
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   8
            Cols            =   13
            ScrollTrack     =   -1  'True
            MousePointer    =   15
            FormatString    =   $"frmKtv.frx":0058
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmKtv.frx":01A1
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
            TabIndex        =   8
            Top             =   320
            Width           =   540
         End
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   4800
         Index           =   0
         Left            =   -74880
         OleObjectBlob   =   "frmKtv.frx":02FB
         TabIndex        =   1
         Top             =   480
         Width           =   12735
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   5200
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmKtv.frx":2F8E
         TabIndex        =   9
         Top             =   360
         Width           =   12735
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   12735
         Begin VB.CheckBox Check2 
            Caption         =   "Stampa Grafico 3D"
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
            TabIndex        =   23
            Top             =   520
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Stampa Grafico 2D"
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
            TabIndex        =   22
            Top             =   180
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CommandButton cmdEsportaEsame 
            Caption         =   "&Esporta Kt/V"
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
            Left            =   4800
            TabIndex        =   21
            Top             =   240
            Width           =   1815
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
            Left            =   3240
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdImportaEsami 
            Caption         =   "&Importa esami"
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
            Left            =   6840
            TabIndex        =   10
            Top             =   240
            Width           =   1815
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
            Height          =   480
            Left            =   11400
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnnulla 
            Caption         =   "&Annulla digitazione"
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
            Height          =   480
            Left            =   8880
            TabIndex        =   5
            Top             =   240
            Width           =   2295
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmKtv.frx":5C3D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   450
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
         Left            =   12120
         TabIndex        =   19
         Top             =   360
         Width           =   615
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
         Left            =   7200
         TabIndex        =   18
         Top             =   360
         Width           =   3135
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
         Left            =   2280
         TabIndex        =   17
         Top             =   360
         Width           =   3255
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
         Left            =   6360
         TabIndex        =   14
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
         Left            =   11400
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmKtv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmKtv.frm
'
' <b>Descrizione</b>: Scheda KTV associata alla tab KTV
'
' @remarks
'
' @author
'
' @date 07/02/2011 18.41
Option Explicit

Dim lettera As String
'' rs della scheda
Dim rsKTV As Recordset
Dim vCol As Integer
Dim vRow As Integer
'' oggetto che gestisce la lista CAnnulla
Dim objAnnulla As CAnnulla
'' evita di lanciare l'evento click delle cbo quando si sta caricando
Dim stoCaricando As Boolean
Dim intPazientiKey As Integer

Dim rsDialisi As Recordset
Private Peso_Post As Single
Private diff_peso As Single
Private durata As Single
Private attivaPass As Boolean
Private cod_paz As Integer
Private data As Date

Public Property Get getData() As Date
    getData = data
End Property

Public Property Let LetData(ByVal vdata As Date)
    data = vdata
End Property

Public Property Get getCod_Paz() As Integer
    getCod_Paz = cod_paz
End Property

Public Property Let LetCod_paz(ByVal vcod_paz As Integer)
    cod_paz = vcod_paz
End Property

Public Property Get getAttiva() As Boolean
    getAttiva = attivaPass
End Property

Public Property Let LetAttiva(ByVal attiva As Boolean)
    attivaPass = attiva
End Property

Public Property Get GetPeso_Post() As Single
    GetPeso_Post = Peso_Post
End Property

Public Property Let LetPeso_Post(ByVal vPeso_Post As Single)
    Peso_Post = vPeso_Post
End Property

Public Property Get GetDiff_Peso() As Single
    GetDiff_Peso = diff_peso
End Property

Public Property Let LetDiff_Peso(ByVal vDiff_Peso As Single)
    diff_peso = vDiff_Peso
End Property

Public Property Get GetDurata() As Single
    GetDurata = durata
End Property

Public Property Let LetDurata(ByVal vDurata)
    durata = vDurata
End Property

'' Se la scheda è stata richiamata dal form schede dialisi giornaliere carica i dati del paziente
' altrimenti richiama il form Trova
Private Sub Form_Activate()
    Dim i As Integer
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    If attivaPass Then
        cboAnno.ListIndex = IIf(cboAnno.List(0) = Year(data), 0, 1)
        intPazientiKey = cod_paz
        With flxGriglia
            ' verifica se esistono gia i dati
            If .TextMatrix(7, Month(data)) <> "" Then
                If MsgBox("La scheda del paziente " & UCase(lblCognome) & " " & UCase(lblNome) & vbCrLf & _
                          "contiene già dei dati sul calcolo del kt/v." & vbCrLf & _
                          "Sostituire le informazioni?", vbQuestion + vbYesNo, "Calcolo del Kt/v") = vbNo Then
                    Unload Me
                    Exit Sub
                End If
            End If
            ' carica i dati nella flx
            .TextMatrix(3, Month(data)) = durata
            .TextMatrix(4, Month(data)) = VirgolaOrPunto(CStr(diff_peso), ",")
            .TextMatrix(5, Month(data)) = VirgolaOrPunto(CStr(Peso_Post), ",")
            .TextMatrix(7, Month(data)) = Day(data)
            ' colora la colonna
            flxGriglia.Col = Month(data)
            Call ColoraColonna
            ' salva i dati nel db
            vCol = Month(data)
            For i = 3 To 5
                vRow = i
                Call SalvaModifiche
            Next i
            vRow = 7
            Call SalvaModifiche
            vRow = 0
            vCol = 0
            cmdImportaEsami.Enabled = False
        End With
    Else
        If intPazientiKey = 0 Then
            cmdTrova_Click
        End If
    End If
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
        .MousePointer = flexCustom
        .Row = 0
        For i = 0 To 12
            .Col = i
            .CellFontBold = True
        Next i
        .Col = 0
        For i = 1 To 7
            .Row = i
            .CellFontBold = True
            .CellBackColor = RGB(231, 255, 255)
        Next i
    End With
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    For k = 0 To 1
        ' valore massimo di ktv è 4
        grafico(k).Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 4
        For i = 1 To 12
            grafico(k).Column = 1
            grafico(k).Row = i
            grafico(k).data = 0
            grafico(k).RowLabel = UCase(MonthName(i, True))
        Next i
    Next k
    tabScheda.Tab = 0
    Set objAnnulla = New CAnnulla
    Peso_Post = 0
    diff_peso = 0
    durata = 0
    attivaPass = False
    cod_paz = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

'' Salva i dati di un singolo campo del record
Private Sub SalvaModifiche()
    Dim nome_campo As String
    With flxGriglia
        Select Case vRow
            Case 1: nome_campo = "UREA_POST"
            Case 2: nome_campo = "UREA_PRE"
            Case 3: nome_campo = "DURATA"
            Case 4: nome_campo = "VOLUME"
            Case 5: nome_campo = "PESO"
            Case 7: nome_campo = "GIORNO"
        End Select
        Set rsKTV = New Recordset
        rsKTV.Open "SELECT * FROM KTV WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND ANNO=" & cboAnno.Text & " AND MESE=" & vCol, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If rsKTV.EOF And rsKTV.BOF Then
            ' nn esiste e quindi lo aggiunge
            rsKTV.AddNew
            rsKTV("KEY") = GetNumero("KTV")
            rsKTV("CODICE_PAZIENTE") = intPazientiKey
            rsKTV("ANNO") = cboAnno.Text
            rsKTV("MESE") = vCol
            rsKTV(nome_campo) = IIf(.TextMatrix(vRow, vCol) = "", 0, VirgolaOrPunto(.TextMatrix(vRow, vCol), ","))
            rsKTV.Update
        Else
           '  esiste e lo modifica
            rsKTV(nome_campo) = IIf(.TextMatrix(vRow, vCol) = "", 0, VirgolaOrPunto(.TextMatrix(vRow, vCol), ","))
            rsKTV.Update
        End If
        Set rsKTV = Nothing
    End With
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    intPazientiKey = 0
    lblCognome = ""
    lblEta = ""
    lblNome = ""
    Call Pulisci
    objAnnulla.Refresh
End Sub

'' Pulisce solo la flx
Private Sub Pulisci()
    Dim i As Integer
    Dim k As Integer
    For i = 1 To 12
        For k = 1 To 7
            flxGriglia.TextMatrix(k, i) = ""
        Next k
        For k = 0 To 1
            grafico(k).Column = 1
            grafico(k).Row = i
            grafico(k).data = 0
        Next k
    Next i
End Sub

'' Calcola il KTV della colonna selezionata
'
' @param vCol colonna selezionata
' @return valore del KTV
Private Function CalcolaKtv(vCol As Integer) As Double
    On Error GoTo gestione
    Dim C(1 To 5) As Single
    Dim i As Integer
    '(4-3,5*C1/C2)*(C4/C5)-LN(C1/C2-0,008*C3)
    With flxGriglia
        For i = 1 To 5
            If .TextMatrix(i, vCol) = "" Then Exit Function
            C(i) = CSng(VirgolaOrPunto(CStr(.TextMatrix(i, vCol)), "."))
        Next i
        CalcolaKtv = Format((4 - 3.5 * C(1) / C(2)) * (C(4) / C(5)) - Log(C(1) / C(2) - 0.008 * C(3)), "##.##")
 
    End With
    Exit Function
gestione:
    If Err.Number = 13 Or Err.Number = 5 Then
        MsgBox "Impossibile calcolare il valore del Kt/V con i dati presenti", vbCritical, "Attenzione"
    Else
        MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
    End If
End Function

'' Carica la scheda nella flx
Private Sub CaricaScheda()
    Dim i As Integer
    Dim k As Integer
    Dim valore As Single
    vRow = 0
    vCol = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    With flxGriglia
        Set rsKTV = New Recordset
        For i = 1 To 12
            rsKTV.Open "SELECT * FROM KTV WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND ANNO=" & cboAnno.Text & " AND MESE=" & i, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsKTV.EOF And rsKTV.BOF) Then
                .TextMatrix(1, i) = VirgolaOrPunto(rsKTV("UREA_POST") & "", ",")
                .TextMatrix(2, i) = VirgolaOrPunto(rsKTV("UREA_PRE") & "", ",")
                .TextMatrix(3, i) = rsKTV("DURATA") & ""
                .TextMatrix(4, i) = VirgolaOrPunto(rsKTV("VOLUME") & "", ",")
                .TextMatrix(5, i) = VirgolaOrPunto(rsKTV("PESO") & "", ",")
                .TextMatrix(7, i) = rsKTV("GIORNO") & ""
                valore = CalcolaKtv(i)
                .TextMatrix(6, i) = VirgolaOrPunto(CStr(valore), ",")
                For k = 0 To 1
                    grafico(k).Column = 1
                    grafico(k).Row = i
                    grafico(k).data = valore
                Next k
            End If
            rsKTV.Close
        Next i
        Set rsKTV = Nothing
    End With
End Sub

'' Colora l'intera colonna di una flx
Private Sub ColoraColonna(Optional colore As ColorConstants = vbCyan)
    Dim i As Integer
    Dim k As Integer
    Dim Col As Integer
    Dim riga As Integer
    Dim colAppo As ColorConstants
    
    If flxGriglia.Row = 0 Or flxGriglia.Col = 0 Then Exit Sub
    
    riga = flxGriglia.Row
    Col = flxGriglia.Col
    ' discolora la colonna colorata
    For k = 1 To flxGriglia.Cols - 1
        flxGriglia.Col = k
        ' utilizzo un var di appoggio perche cosi funziona
        colAppo = flxGriglia.CellBackColor
        If colAppo <> vbWhite And colAppo <> 0 Then
            For i = flxGriglia.FixedRows To flxGriglia.Rows - 1
                flxGriglia.Row = i
                flxGriglia.CellBackColor = vbWhite
            Next i
            Exit For
        End If
    Next k
    flxGriglia.Col = Col
    ' cambia colore della riga
    For i = flxGriglia.FixedRows To flxGriglia.Rows - 1
        flxGriglia.Row = i
        flxGriglia.CellBackColor = colore
    Next i
    flxGriglia.Row = riga
End Sub

Public Sub DiscoloraColonna()
    ' Discolora la colonna di una flex
    Dim i As Integer
    Dim k As Integer
    Dim Col As Integer
    Dim riga As Integer
    Dim colAppo As ColorConstants
       
    For k = 1 To flxGriglia.Cols - 1
        flxGriglia.Col = k
        ' utilizzo un var di appoggio perche cosi funziona
        colAppo = flxGriglia.CellBackColor
        If colAppo <> vbWhite And colAppo <> 0 Then
            For i = flxGriglia.FixedRows To flxGriglia.Rows - 1
                flxGriglia.Row = i
                flxGriglia.CellBackColor = vbWhite
            Next i
            Exit For
        End If
    Next k
    
    ' Imposto la colonna uguale a 0 per evitare problemi nella Sub SalvaModifiche
    flxGriglia.Col = 0
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

'' Importa i valori dagli esami Azotemia post e pre nelle righe Urea post - pre
Private Sub cmdImportaEsami_Click()
    Dim rsDataset As Recordset
    Dim continua As Boolean
    Dim valore As Single
    Dim strSql As String
    
    If flxGriglia.Col = 0 Then
        MsgBox "Selezionare il mese degli esami da importare", vbInformation, "Informazione"
    Else
        Set rsDataset = New Recordset
        strSql = "SELECT    VOCI_ESAMI.NOME, VALORE " & _
                 "FROM      ((ANAMNESI_ESAMI " & _
                 "          INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                 "          INNER JOIN VOCI_ESAMI ON VOCI_ESAMI.KEY=ESAMI_LAB.CODICE_ESAME) " & _
                 "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                 "          (NOME like '%Azotemia Post%' OR NOME like '%Azotemia Pre%') AND " & _
                 "          MONTH([DATA])=" & flxGriglia.Col & " AND " & _
                 "          YEAR([DATA])=" & cboAnno.Text & " " & _
                 "ORDER BY  DATA DESC"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If flxGriglia.TextMatrix(1, flxGriglia.Col) <> "" Or flxGriglia.TextMatrix(2, flxGriglia.Col) <> "" Then
                If MsgBox("I valori degli esami sono già presenti" & vbCrLf & "Vuoi sovrascriverli?", vbQuestion + vbYesNo + vbDefaultButton2, "Importa esami") = vbYes Then
                    continua = True
                Else
                    continua = False
                End If
            Else
                continua = True
            End If
            If continua Then
                rsDataset.Filter = ("NOME like '%Azotemia Post%'")
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    flxGriglia.TextMatrix(1, flxGriglia.Col) = VirgolaOrPunto(rsDataset("VALORE"), ",")
                    vRow = 1
                    Call SalvaModifiche
                End If
                rsDataset.Filter = ("NOME like '%Azotemia Pre%'")
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    flxGriglia.TextMatrix(2, flxGriglia.Col) = VirgolaOrPunto(rsDataset("VALORE"), ",")
                    vRow = 2
                    Call SalvaModifiche
                End If
                valore = CalcolaKtv(flxGriglia.Col)
                flxGriglia.TextMatrix(6, flxGriglia.Col) = VirgolaOrPunto(CStr(valore), ",")
                Call CaricaScheda
                Call DiscoloraColonna
            End If
        Else
            MsgBox "Nessun esame per il mese selezionato", vbInformation, "Importa esami"
        End If
        Set rsDataset = Nothing
    End If
End Sub

Private Sub cmdEsportaEsame_Click()
    Dim rsDataset As Recordset
    Dim keyEsame As Integer
    Dim keyGruppo As Integer
    Dim keyAnamnesi As Long
    Dim keyRecord As Long
    Dim keyNuovo As Long
    
    If flxGriglia.TextMatrix(6, vCol) = "" Then
        MsgBox "IMPOSSIBILE ESPORTARE!!!" & vbCrLf & "Valori non definiti", vbCritical, "Attenzione"
        Exit Sub
    End If
        
    If flxGriglia.Col = 0 Then
        MsgBox "Selezionare il mese dell'esame da esportare", vbCritical, "Attenzione"
    Else
        If flxGriglia.TextMatrix(6, vCol) <> 0 Then
            Set rsDataset = New Recordset
            
            ' verifica se esiste l'esame kt/v
            rsDataset.Open "SELECT * FROM VOCI_ESAMI WHERE NOME like '%KT/V%' ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                keyEsame = rsDataset("KEY")
            Else
                MsgBox "La voce KT/V non è presente nella lista degli esami di laboratorio", vbCritical, "ATTENZIONE!!!"
            End If
            rsDataset.Close
            
            If keyEsame <> 0 Then
                ' verifica se esiste l'associazione con qualche gruppo esami lab
                rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_ESAME=" & keyEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    keyGruppo = rsDataset("CODICE_GRUPPO")
                Else
                    MsgBox "La voce KT/V è presente nella lista degli esami di laboratorio ma non è associata ad alcun gruppo", vbCritical, "ATTENZIONE!!!"
                End If
                rsDataset.Close
                
                If keyGruppo <> 0 Then
                    ' verifica se esiste un record anamnesi per il mese selezionato del gruppo
                    rsDataset.Open "SELECT * FROM ANAMNESI_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & keyGruppo & " AND DATA BETWEEN #" & Format(vCol, "00") & "/01/" & cboAnno.Text & "# AND #" & GetUltimoGiorno(vCol, cboAnno.Text, True) & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                    If Not (rsDataset.EOF And rsDataset.BOF) Then
                        If rsDataset.RecordCount > 1 Then
                            rsDataset.Filter = "DATA>=" & DateValue(flxGriglia.TextMatrix(7, vCol) & "/" & vCol & "/" & cboAnno.Text)
                            If rsDataset.RecordCount > 1 Then
                                ' mostra un pannello con le date filtrare e fa scegliere all'utente
                                tElenca.Tipo = tpESPORTAESAMI
                                tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & keyGruppo & " AND DATA>= #" & DateValue(vCol & "/" & flxGriglia.TextMatrix(7, vCol) & "/" & cboAnno.Text) & "# AND DATA<=#" & GetUltimoGiorno(vCol, cboAnno.Text, True) & "#"
                                frmElencaDate.Show 1
                                If laData = "" Then Exit Sub
                                rsDataset.Filter = "DATA=#" & laData & "#"
                                If rsDataset.RecordCount > 0 Then
                                    keyAnamnesi = rsDataset("KEY")
                                Else
                                    keyAnamnesi = 0
                                End If
                            Else
                                keyAnamnesi = rsDataset("KEY")
                            End If
                            rsDataset.Filter = ""
                        Else
                            keyAnamnesi = rsDataset("KEY")
                        End If
                    Else
                        keyAnamnesi = 0
                    End If
                    rsDataset.Close
                    
                    If keyAnamnesi <> 0 Then
                        ' verifica se esiste un record per l'esame per l'anamnesi trovata
                        rsDataset.Open "SELECT * FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & keyAnamnesi & " AND CODICE_ESAME=" & keyEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        If Not (rsDataset.EOF And rsDataset.BOF) Then
                            keyRecord = rsDataset("KEY")
                        End If
                        rsDataset.Close
                        
                        If keyRecord <> 0 Then
                            If MsgBox("Il valore del Kt/V è già presente" & vbCrLf & "Vuoi sovrascriverlo?", vbQuestion + vbYesNo + vbDefaultButton2, "Esporta esame") = vbYes Then
                                ' modifica il dato esistente
                                rsDataset.Open "SELECT * FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & keyAnamnesi & " AND CODICE_ESAME=" & keyEsame, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                                If Not (rsDataset.EOF And rsDataset.BOF) Then
                                   rsDataset("VALORE") = flxGriglia.TextMatrix(6, vCol)
                                   rsDataset.Update
                                End If
                                rsDataset.Close
                                MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                            End If
                        Else
                            ' inserisce un nuovo esame
                            keyNuovo = GetNumero("ESAMI_LAB")
                            rsDataset.Open "ESAMI_LAB", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
                            rsDataset.AddNew
                            rsDataset("KEY") = keyNuovo
                            rsDataset("CODICE_ESAME") = keyEsame
                            rsDataset("CODICE_ANAMNESI_ESAMI") = keyAnamnesi
                            rsDataset("VALORE") = flxGriglia.TextMatrix(6, vCol)
                            rsDataset.Update
                            rsDataset.Close
                            MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                        End If
                    Else
                        ' inserisce la nuova data del gruppo esami
                        keyAnamnesi = GetNumero("ANAMNESI_ESAMI")
                        rsDataset.Open "ANAMNESI_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                        rsDataset.AddNew
                        rsDataset("KEY") = keyAnamnesi
                        rsDataset("CODICE_PAZIENTE") = intPazientiKey
                        rsDataset("DATA") = DateValue(flxGriglia.TextMatrix(7, vCol) & "/" & vCol & "/" & cboAnno.Text)
                        rsDataset("CODICE_GRUPPO") = keyGruppo
                        rsDataset("UTENTE_MODIFICATORE") = tAccesso.key
                        rsDataset.Update
                        rsDataset.Close
                        
                        ' inserisce un nuovo esame
                        keyNuovo = GetNumero("ESAMI_LAB")
                        rsDataset.Open "ESAMI_LAB", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
                        rsDataset.AddNew
                        rsDataset("KEY") = keyNuovo
                        rsDataset("CODICE_ESAME") = keyEsame
                        rsDataset("CODICE_ANAMNESI_ESAMI") = keyAnamnesi
                        rsDataset("VALORE") = flxGriglia.TextMatrix(6, vCol)
                        rsDataset.Update
                        rsDataset.Close
                        MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                    End If
                End If
            End If
            Set rsDataset = Nothing
        Else
            MsgBox "Per il mese selezionato impossibile calcolare il KT/V", vbCritical, "Attenzione"
        End If
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If

    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDialisi("COGNOME") & " " & rsDialisi("NOME")
    structIntestazione.sDataPaziente = rsDialisi("DATA_NASCITA")
    codiceId = rsDialisi("CODICE_ID")
    Set rsDialisi = Nothing

    strSql = "SHAPE APPEND  NEW adVarChar (10) as UREA_POST_GEN, " & _
                    "       NEW adVarChar (10) as UREA_POST_FEB, " & _
                    "       NEW adVarChar (10) as UREA_POST_MAR, " & _
                    "       NEW adVarChar (10) as UREA_POST_APR, " & _
                    "       NEW adVarChar (10) as UREA_POST_MAG, " & _
                    "       NEW adVarChar (10) as UREA_POST_GIU, " & _
                    "       NEW adVarChar (10) as UREA_POST_LUG, " & _
                    "       NEW adVarChar (10) as UREA_POST_AGO, " & _
                    "       NEW adVarChar (10) as UREA_POST_SET, " & _
                    "       NEW adVarChar (10) as UREA_POST_OTT, " & _
                    "       NEW adVarChar (10) as UREA_POST_NOV, " & _
                    "       NEW adVarChar (10) as UREA_POST_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as UREA_PRE_GEN, " & _
                    "       NEW adVarChar (10) as UREA_PRE_FEB, " & _
                    "       NEW adVarChar (10) as UREA_PRE_MAR, " & _
                    "       NEW adVarChar (10) as UREA_PRE_APR, " & _
                    "       NEW adVarChar (10) as UREA_PRE_MAG, " & _
                    "       NEW adVarChar (10) as UREA_PRE_GIU, " & _
                    "       NEW adVarChar (10) as UREA_PRE_LUG, " & _
                    "       NEW adVarChar (10) as UREA_PRE_AGO, " & _
                    "       NEW adVarChar (10) as UREA_PRE_SET, " & _
                    "       NEW adVarChar (10) as UREA_PRE_OTT, " & _
                    "       NEW adVarChar (10) as UREA_PRE_NOV, " & _
                    "       NEW adVarChar (10) as UREA_PRE_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_GEN, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_FEB, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_MAR, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_APR, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_MAG, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_GIU, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_LUG, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_AGO, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_SET, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_OTT, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_NOV, " & _
                    "       NEW adVarChar (10) as DURATA_DIALISI_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as VOL_ULTRA_GEN, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_FEB, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_MAR, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_APR, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_MAG, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_GIU, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_LUG, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_AGO, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_SET, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_OTT, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_NOV, " & _
                    "       NEW adVarChar (10) as VOL_ULTRA_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as PESO_GEN, " & _
                    "       NEW adVarChar (10) as PESO_FEB, " & _
                    "       NEW adVarChar (10) as PESO_MAR, " & _
                    "       NEW adVarChar (10) as PESO_APR, " & _
                    "       NEW adVarChar (10) as PESO_MAG, " & _
                    "       NEW adVarChar (10) as PESO_GIU, " & _
                    "       NEW adVarChar (10) as PESO_LUG, " & _
                    "       NEW adVarChar (10) as PESO_AGO, " & _
                    "       NEW adVarChar (10) as PESO_SET, " & _
                    "       NEW adVarChar (10) as PESO_OTT, " & _
                    "       NEW adVarChar (10) as PESO_NOV, " & _
                    "       NEW adVarChar (10) as PESO_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as KTV_GEN, " & _
                    "       NEW adVarChar (10) as KTV_FEB, " & _
                    "       NEW adVarChar (10) as KTV_MAR, " & _
                    "       NEW adVarChar (10) as KTV_APR, " & _
                    "       NEW adVarChar (10) as KTV_MAG, " & _
                    "       NEW adVarChar (10) as KTV_GIU, " & _
                    "       NEW adVarChar (10) as KTV_LUG, " & _
                    "       NEW adVarChar (10) as KTV_AGO, " & _
                    "       NEW adVarChar (10) as KTV_SET, " & _
                    "       NEW adVarChar (10) as KTV_OTT, " & _
                    "       NEW adVarChar (10) as KTV_NOV, " & _
                    "       NEW adVarChar (10) as KTV_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_GEN, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_FEB, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_MAR, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_APR, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_MAG, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_GIU, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_LUG, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_AGO, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_SET, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_OTT, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_NOV, " & _
                    "       NEW adVarChar (10) as PRELIEVO_GIORNO_DIC "
              
              
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Dim vett(1 To 12) As String
    vett(1) = "GEN"
    vett(2) = "FEB"
    vett(3) = "MAR"
    vett(4) = "APR"
    vett(5) = "MAG"
    vett(6) = "GIU"
    vett(7) = "LUG"
    vett(8) = "AGO"
    vett(9) = "SET"
    vett(10) = "OTT"
    vett(11) = "NOV"
    vett(12) = "DIC"
    Dim valore As Double
    Dim i As Integer
    Dim C(1 To 5) As Double
    
    With rsMain
        .AddNew
        
        Set rsKTV = New Recordset
        For i = 1 To 12
            rsKTV.Open "SELECT * FROM KTV WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND ANNO=" & cboAnno.Text & " AND MESE=" & i, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsKTV.EOF And rsKTV.BOF) Then
                .Fields("UREA_POST_" & vett(i)) = VirgolaOrPunto(rsKTV("UREA_POST") & "", ",")
                .Fields("UREA_PRE_" & vett(i)) = VirgolaOrPunto(rsKTV("UREA_PRE") & "", ",")
                .Fields("DURATA_DIALISI_" & vett(i)) = rsKTV("DURATA") & ""
                .Fields("VOL_ULTRA_" & vett(i)) = VirgolaOrPunto(rsKTV("VOLUME") & "", ",")
                .Fields("PESO_" & vett(i)) = VirgolaOrPunto(rsKTV("PESO") & "", ",")
                .Fields("PRELIEVO_GIORNO_" & vett(i)) = rsKTV("GIORNO") & ""
                If .Fields("UREA_POST_" & vett(i)) <> "" And .Fields("UREA_PRE_" & vett(i)) <> "" And .Fields("DURATA_DIALISI_" & vett(i)) <> "" And .Fields("VOL_ULTRA_" & vett(i)) <> "" And .Fields("PESO_" & vett(i)) <> "" Then
                    C(1) = CSng(VirgolaOrPunto(CStr(.Fields("UREA_POST_" & vett(i))), "."))
                    C(2) = CSng(VirgolaOrPunto(CStr(.Fields("UREA_PRE_" & vett(i))), "."))
                    C(3) = CSng(VirgolaOrPunto(CStr(.Fields("DURATA_DIALISI_" & vett(i))), "."))
                    C(4) = CSng(VirgolaOrPunto(CStr(.Fields("VOL_ULTRA_" & vett(i))), "."))
                    C(5) = CSng(VirgolaOrPunto(CStr(.Fields("PESO_" & vett(i))), "."))
                    
                    valore = Format((4 - 3.5 * C(1) / C(2)) * (C(4) / C(5)) - Log(C(1) / C(2) - 0.008 * C(3)), "##.##")
                    .Fields("KTV_" & vett(i)) = VirgolaOrPunto(CStr(valore), ",")
                Else
                    .Fields("KTV_" & vett(i)) = ""
                End If
            Else
                .Fields("UREA_POST_" & vett(i)) = ""
                .Fields("UREA_PRE_" & vett(i)) = ""
                .Fields("DURATA_DIALISI_" & vett(i)) = ""
                .Fields("VOL_ULTRA_" & vett(i)) = ""
                .Fields("PESO_" & vett(i)) = ""
                .Fields("PRELIEVO_GIORNO_" & vett(i)) = ""
                .Fields("KTV_" & vett(i)) = ""
            End If
            rsKTV.Close
        Next i
        Set rsKTV = Nothing
    End With

    Set rptCalcoloKtV.DataSource = rsMain
    rptCalcoloKtV.TopMargin = 0
    rptCalcoloKtV.BottomMargin = 0
    rptCalcoloKtV.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptCalcoloKtV.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptCalcoloKtV.Sections("Intestazione").Controls.Item("lblEta").Caption = lblEta.Caption
    rptCalcoloKtV.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    
    'stampa i grafici
    
    If Check1 Then
        
        'Porta il Grafico nella pagina di stampa
        With rptCalcoloKtV.Sections("corpo")
        
        'Imposta se necessarie le dimensioni dell'immagine:
        With .Controls("Image2d")
   '     .Height = 15000
   '     .Left = 0
   '     .Top = 10
   '     .Width = 11200
   '     .PictureAlignment = rptPACenter
   '     .SizeMode = 0
        grafico(0).EditCopy
        Set .Picture = Clipboard.getData
        End With
        End With
    Else
   '    cancella l'immagine nella clipboard e nel campo image del report
        Clipboard.Clear
        With rptCalcoloKtV.Sections("corpo")
        With .Controls("Image2d")
        Set .Picture = LoadPicture()
        End With
        End With
    End If

    If Check2 Then
        With rptCalcoloKtV.Sections("corpo")
        With .Controls("Image3d")
  '      .Top = 10
        grafico(1).EditCopy
        Set .Picture = Clipboard.getData
        End With
        End With
    Else
        Clipboard.Clear
        With rptCalcoloKtV.Sections("corpo")
        With .Controls("Image3d")
        Set .Picture = LoadPicture()
        End With
        End With
    End If
    
    rptCalcoloKtV.PrintReport True, rptRangeAllPages
    Clipboard.Clear
End Sub

Private Sub cmdAnnulla_Click()
    Dim k As Integer
    Dim Dato As String
    Dim Col As Integer
    Dim Row As Integer
    Dim valore As Single
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    ' row identifica la riga e non il key
    Row = objAnnulla.Row
    With flxGriglia
        .TextMatrix(Row, Col) = Dato
        valore = CalcolaKtv(Col)
        .TextMatrix(6, Col) = VirgolaOrPunto(CStr(valore), ",")
        For k = 0 To 1
            grafico(k).Column = 1
            grafico(k).Row = Col
            grafico(k).data = valore
        Next k
        objAnnulla.Remove
        ' modifica anche il db
        vRow = Row
        Call SalvaModifiche
        If objAnnulla.Vuoto = True Then
            cmdAnnulla.Enabled = False
        End If
    End With
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    Call DiscoloraColonna
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    If tTrova.keyReturn = 0 Then
        Unload frmKtv
    Else
        Call CaricaPaziente
    End If
End Sub

Private Sub flxGriglia_Click()
    vCol = flxGriglia.Col
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraColonna(vbWhite)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraColonna
        vRow = flxGriglia.Row
    End If
End Sub

Private Sub flxGriglia_DblClick()
    With flxGriglia
        .SetFocus
        If .Col <> 0 And (.Row = 1 Or .Row = 2 Or .Row = 4) Then
            If .Row <> 4 And attivaPass Then Exit Sub
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

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        Exit Sub
    End If
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
    ' carica la scheda
    Call CaricaScheda
End Sub

'' Richiama CaricaScheda
Private Sub cboAnno_Click()
    If stoCaricando Then Exit Sub
    Call Pulisci
    Call DiscoloraColonna
    Call CaricaScheda
End Sub

Private Sub txtAppo_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtAppo, lettera)
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii <> vbKeyReturn Then
        If flxGriglia.Row = 8 Then Exit Sub
        If KeyAscii = 44 Then KeyAscii = 46
        lettera = Chr(KeyAscii)
    Else
        txtAppo_Validate (False)
    End If
End Sub

Private Sub Elimina()
    Set rsKTV = New Recordset
    rsKTV.Open "SELECT * FROM KTV WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND MESE=" & vCol & " AND ANNO=" & cboAnno.Text, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsKTV.EOF And rsKTV.BOF) Then
        rsKTV.Delete
        rsKTV.Update
    End If
    Set rsKTV = Nothing
End Sub

Private Sub txtAppo_LostFocus()
    Dim k As Integer
    Dim valore As Single
    
    txtAppo.Visible = False
    If txtAppo = "" And vRow <> 8 Then txtAppo = 0
    flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, vRow)
    cmdAnnulla.Enabled = True
    Call SalvaModifiche
    valore = CalcolaKtv(vCol)
    flxGriglia.TextMatrix(6, vCol) = VirgolaOrPunto(CStr(valore), ",")
    For k = 0 To 1
        grafico(k).Column = 1
        grafico(k).Row = vCol
        grafico(k).data = valore
    Next k

    With flxGriglia
        If (.TextMatrix(1, vCol) = "" Or .TextMatrix(1, vCol) = "0") And (.TextMatrix(2, vCol) = "" Or .TextMatrix(2, vCol) = "0") And (.TextMatrix(3, vCol) = "" Or .TextMatrix(3, vCol) = "0") And (.TextMatrix(4, vCol) = "" Or .TextMatrix(4, vCol) = "0") And (.TextMatrix(5, vCol) = "" Or .TextMatrix(5, vCol) = "0") And (.TextMatrix(7, vCol) = "" Or .TextMatrix(7, vCol) = "0") Then
            Call Elimina
            .TextMatrix(1, vCol) = ""
            .TextMatrix(2, vCol) = ""
            .TextMatrix(3, vCol) = ""
            .TextMatrix(4, vCol) = ""
            .TextMatrix(5, vCol) = ""
            .TextMatrix(6, vCol) = ""
            .TextMatrix(7, vCol) = ""
        End If
    End With
    txtAppo.MaxLength = 0
End Sub

Private Sub txtAppo_Validate(Cancel As Boolean)
    If txtAppo = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtAppo.Text)
    End If
    If Not Cancel Then
        flxGriglia.SetFocus
    End If
End Sub

