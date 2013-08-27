VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmPianoLavoro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Piano di lavoro"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   12015
      Begin VB.OptionButton optTurno 
         Caption         =   "SER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optTurno 
         Caption         =   "MAT"
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
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTurno 
         Caption         =   "POM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TURNO"
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
         Left            =   4560
         TabIndex        =   24
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblGiorno 
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9120
         TabIndex        =   23
         Top             =   300
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Selezionare la data"
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
         TabIndex        =   22
         Top             =   300
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmdSposta 
         Caption         =   "<"
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
         Height          =   375
         Index           =   1
         Left            =   5700
         TabIndex        =   7
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton cmdSposta 
         Caption         =   ">"
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
         Height          =   375
         Index           =   0
         Left            =   5700
         TabIndex        =   6
         Top             =   3720
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrigliaSx 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FormatString    =   "|  Cognome                            |  Nome                                 |  Data di nascita          "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrigliaDx 
         Height          =   2415
         Left            =   6360
         TabIndex        =   8
         Top             =   3000
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FormatString    =   "|  Cognome                            |  Nome                                 |  Data di nascita          "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfermieri 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FormatString    =   "| Cognome                                  | Nome                                  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblNonAssociati 
         AutoSize        =   -1  'True
         Caption         =   "Pazienti non associati:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6360
         TabIndex        =   17
         Top             =   2760
         Width           =   2340
      End
      Begin VB.Label lblAssociati 
         AutoSize        =   -1  'True
         Caption         =   "Pazienti associati:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1905
      End
      Begin VB.Label lblInfermieri 
         AutoSize        =   -1  'True
         Caption         =   "Infermieri:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   12015
      Begin VB.ComboBox cboMedico 
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
         ItemData        =   "frmPianoLavoro.frx":0000
         Left            =   7440
         List            =   "frmPianoLavoro.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox cboCaposala 
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
         ItemData        =   "frmPianoLavoro.frx":0004
         Left            =   1560
         List            =   "frmPianoLavoro.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Medico"
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
         Left            =   6360
         TabIndex        =   20
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Coordinatore"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   6720
      Width           =   12015
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
         Left            =   7200
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdPulisci 
         Caption         =   "&Pulisci"
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
         Left            =   8880
         TabIndex        =   12
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
         Left            =   10560
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPianoLavoro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPianoLavoro.frm
'
' <b>Descrizione</b>: Scheda Piano di lavoro associata alla tab PIANO_LAVORO e ASSOCIAZIONI_PIANO_LAVORO
'
' @remarks
'
' @author
'
' @date 19/03/2011 12.45
Option Explicit

' rs della scheda
Dim rsPiano As Recordset
' rs disconnesso di appoggio da ASSOCIAZIONI_PIANO_LAVORO
Dim rsAppo As Recordset
' valore del campo codice_piano di ASSOCIAZIONI_PIANO_LAVORO
Dim keyId As Integer
Dim modifica As Boolean

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
    
    cmdPulisci_Click
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxInfermieri
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 2
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
    End With
    With flxGrigliaSx
        .ColWidth(0) = 0
        .Row = 0
        .Rows = 1
        For i = 1 To 3
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
    End With
    With flxGrigliaDx
        .ColWidth(0) = 0
        .Row = 0
        .Rows = 1
        For i = 1 To 3
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
    End With
    modifica = False
    Call ApriRsAppo
End Sub

' Crea il rs disconnesso di appoggio da ASSOCIAZIONI_PIANO_LAVORO
Private Sub ApriRsAppo()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ASSOCIAZIONI_PIANO_LAVORO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Set rsAppo = New Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsAppo.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsAppo.CursorLocation = adUseServer
    rsAppo.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

' Trova un piano di lavoro con la data e il turno e restituisce il key
Private Function TrovaPiano() As Integer
    Dim rsDataset As Recordset
    Dim data As Date
    Dim turno As Integer
    
    If optTurno(0).Value Then
        turno = 1
    ElseIf optTurno(1).Value Then
        turno = 2
    Else
        turno = 3
    End If
    data = DateValue(Month(oData(0).data) & "/" & Day(oData(0).data) & "/" & Year(oData(0).data))
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM PIANO_LAVORO WHERE (DATA=#" & data & "# AND TURNO=" & turno & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        TrovaPiano = rsDataset("KEY")
    Else
        TrovaPiano = 0
    End If
    Set rsDataset = Nothing
End Function

' Crea la condizione per elenco pazienti del turno selezionato
'
' @param sottrai se true sottrae quelli presenti e associati gia in rsAppo
' @return stringa tipo (key in (1,2, 3))
Private Function ElencoPazientiInTurni(sottrai As Boolean) As String
    Dim tipostrTurno As String
    Dim giorno As Integer
    Dim turno As Integer
    Dim condizione As String
    Dim rsPazientiTurni As Recordset
    
    If optTurno(0).Value Then
        turno = 1
        tipostrTurno = "AM_INIZIO"
    ElseIf optTurno(1).Value Then
        turno = 2
        tipostrTurno = "PM_INIZIO"
    Else
        turno = 3
        tipostrTurno = "SR_INIZIO"
    End If
    giorno = Weekday(CDate(oData(0).data), vbMonday)
    
    Set rsPazientiTurni = New Recordset
    rsPazientiTurni.Open "SELECT    PAZIENTI.KEY, PAZIENTI.COGNOME, PAZIENTI.NOME, PAZIENTI.DATA_NASCITA, PAZIENTI.STATO, " & _
                         "          TURNI.AM_INIZIO" & giorno & ",TURNI.PM_INIZIO" & giorno & ",TURNI.SR_INIZIO" & giorno & " " & _
                         "FROM      PAZIENTI INNER JOIN TURNI ON PAZIENTI.KEY = TURNI.CODICE_PAZIENTE " & _
                         "WHERE     ( (PAZIENTI.STATO=0 OR PAZIENTI.STATO=4) AND " & _
                         "          TURNI." & tipostrTurno & giorno & "<>"""" )", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsPazientiTurni.EOF
        condizione = condizione & rsPazientiTurni("KEY") & ","
        rsPazientiTurni.MoveNext
    Loop
    Set rsPazientiTurni = Nothing
    
    If condizione = "" Then
        condizione = " KEY IN (-1)"
    Else
        condizione = Left(condizione, Len(condizione) - 1)
        condizione = " KEY IN (" & condizione & ")"
    End If

    If sottrai And condizione <> "" Then
        ' sottrae quelli gia associati
        condizione = "( " & condizione & " AND NOT " & CreaCondizione() & ")"
    End If
    
    ElencoPazientiInTurni = condizione
End Function

' Pulisce le flx
Private Sub Pulisci()
    Dim i As Integer
    flxGrigliaSx.Rows = 1
    flxGrigliaSx.Row = 0
    flxGrigliaDx.Rows = 1
    flxGrigliaDx.Row = 0
    For i = 0 To 1
        cmdSposta(i).Enabled = False
    Next i
End Sub

' Carica l'intero piano di lavoro
Private Sub CaricaScheda()
    modifica = False
    keyId = 0
    Call CaricaInfermieri
    Call RicaricaComboBox("SELECT KEY, (COGNOME + ' ' + NOME) AS MEDICO FROM MEDICI_DIALISI WHERE ELIMINATO=FALSE", "MEDICO", cboMedico)
    Call PulisciRsAppo
    Call Pulisci
    keyId = TrovaPiano
    If keyId Then
        Call CaricaRsAppo(keyId)
        ' carica il resto dei dati
        Set rsPiano = New Recordset
        rsPiano.Open "SELECT * FROM PIANO_LAVORO WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        cboCaposala.ListIndex = GetListIndexLocale("INFERMIERI", rsPiano("CODICE_CAPOSALA"), cboCaposala)
        cboMedico.ListIndex = GetListIndexLocale("MEDICI_DIALISI", rsPiano("CODICE_MEDICO"), cboMedico)
        Set rsPiano = Nothing
        modifica = True
    End If
    Call CaricaPazienti(flxGrigliaDx, ElencoPazientiInTurni(IIf(keyId, True, False)))
End Sub

' Restituisce il listindex della cbo al nome corrispondente dato il key
' nel caso non sia presente lo aggiunge alla cbo
Private Function GetListIndexLocale(nomeTabella As String, key As Integer, cbo As ComboBox) As Integer
    Dim rsDataset As Recordset
    Dim i As Integer
    Dim nome As String
    Dim trovato As Boolean
    
    If key = -1 Then
        GetListIndexLocale = key
        Exit Function
    End If
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        ' nn lo ha trovato
        GetListIndexLocale = -1
    Else
        nome = rsDataset("COGNOME") & " " & rsDataset("NOME")
        trovato = False
        If nome <> "" Then
            For i = 0 To cbo.ListCount - 1
                If UCase(nome) = UCase(cbo.List(i)) Then
                    GetListIndexLocale = i
                    trovato = True
                    Exit For
                End If
            Next i
        Else
            GetListIndexLocale = -1
        End If
        
        If Not trovato Then
            cbo.AddItem rsDataset("COGNOME") & " " & rsDataset("NOME")
            GetListIndexLocale = cbo.ListCount - 1
        End If
    End If
    Set rsDataset = Nothing
End Function

' Riempe il rs di appoggio disconnesso con i dati del piano
Private Sub CaricaRsAppo(codicePiano As Integer)
    Set rsPiano = New Recordset
    rsPiano.Open "SELECT * FROM ASSOCIAZIONI_PIANO_LAVORO WHERE CODICE_PIANO=" & codicePiano, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsPiano.EOF
        rsAppo.AddNew
        rsAppo("CODICE_PAZIENTE") = rsPiano("CODICE_PAZIENTE")
        rsAppo("CODICE_INFERMIERE") = rsPiano("CODICE_INFERMIERE")
        rsAppo.Update
        rsPiano.MoveNext
    Loop
    Set rsPiano = Nothing
End Sub

' Carica la flx degli infermieri
Private Sub CaricaInfermieri()
    Set rsPiano = New Recordset
    rsPiano.Open "SELECT * FROM INFERMIERI WHERE ELIMINATO=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    cboCaposala.Clear
    With flxInfermieri
        .Rows = 1
        Do While Not rsPiano.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsPiano("KEY")
            .TextMatrix(.Rows - 1, 1) = rsPiano("COGNOME")
            .TextMatrix(.Rows - 1, 2) = rsPiano("NOME")
            If rsPiano("MANSIONE") = 2 Then
                cboCaposala.AddItem rsPiano("COGNOME") & " " & rsPiano("NOME")
            End If
            rsPiano.MoveNext
        Loop
    End With
    Set rsPiano = Nothing
    flxInfermieri.Row = 0
End Sub

' Crea la condizione partendo dai dati di rsAppo disconnesso
Private Function CreaCondizione() As String
    Dim condizione As String
    Dim filtro As String

    ' filtri gli associati nel rsAppo
    If flxInfermieri.Row = 0 Then
        filtro = ""
    Else
        filtro = "CODICE_INFERMIERE=" & flxInfermieri.TextMatrix(flxInfermieri.Row, 0)
    End If
    rsAppo.Filter = filtro
    Do While Not rsAppo.EOF
        condizione = condizione & rsAppo("CODICE_PAZIENTE") & ","
        rsAppo.MoveNext
    Loop
    
    If condizione = "" Then
        condizione = " KEY IN (-1)"
    Else
        condizione = Left(condizione, Len(condizione) - 1)
        condizione = " KEY IN (" & condizione & ")"
    End If

    CreaCondizione = condizione
    rsAppo.Filter = ""
End Function

' Carica i pazienti nella flx
'
' @param flx flx dove caricare i pazienti
' @param condizione condizione per caricare i pazienti
Private Sub CaricaPazienti(flx As MSFlexGrid, condizione As String)
    Dim strSql As String
    Dim rsDatasetCerca As Recordset
    ' pulisce la flx azzerando le righe
    flx.Rows = 1
    strSql = "SELECT * FROM PAZIENTI WHERE " & condizione & " AND (STATO=0 OR STATO=4) ORDER BY COGNOME"
    Set rsDatasetCerca = New Recordset
    rsDatasetCerca.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDatasetCerca.EOF
        With flx
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDatasetCerca("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDatasetCerca("COGNOME")
            .TextMatrix(.Rows - 1, 2) = "" & rsDatasetCerca("NOME")
            .TextMatrix(.Rows - 1, 3) = rsDatasetCerca("DATA_NASCITA")
            rsDatasetCerca.MoveNext
        End With
    Loop
    Set rsDatasetCerca = Nothing
    flx.Row = 0
    lblInfermieri(8) = "Infermieri : " & flxInfermieri.Rows - 1
    lblAssociati = "Pazienti associati: " & flxGrigliaSx.Rows - 1
    lblNonAssociati = "Pazienti non associati: " & flxGrigliaDx.Rows - 1
End Sub

' Pulisce il rs disconnesso di appoggio
Private Sub PulisciRsAppo()
    If Not (rsAppo.EOF And rsAppo.BOF) And rsAppo.RecordCount <> 0 Then
        rsAppo.MoveFirst
        Do While Not rsAppo.EOF
            rsAppo.Delete
            rsAppo.MoveNext
        Loop
    End If
End Sub

' Verifica prima di memorizzare se tutti i dati necessari sono presenti
Private Function Completo() As Boolean
    Completo = False
    If oData(0).txtBox = "" Then
        MsgBox "La data inserita non è corretta", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

' Restituisce il codice dell'infermiere o del medico dalle relative cbo
Private Function GetCodice(nomeTabella As String, cbo As ComboBox) As Integer
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open nomeTabella, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    Do While Not rsDataset.EOF
        If UCase(rsDataset("COGNOME") & " " & rsDataset("NOME")) = UCase(cbo.Text) Then
            GetCodice = rsDataset("KEY")
            Exit Function
        End If
        rsDataset.MoveNext
    Loop
    Set rsDataset = Nothing
    GetCodice = -1
End Function

Private Sub Memorizza()
    Dim cmCommand As New Command
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim numKey As Integer
    Dim turno As Integer
    
    If modifica Then
        numKey = keyId
    Else
        numKey = GetNumero("PIANO_LAVORO")
    End If
    If optTurno(0).Value Then
        turno = 1
    ElseIf optTurno(1).Value Then
        turno = 2
    Else
        turno = 3
    End If
    v_Nomi = Array("KEY", "DATA", "TURNO", "CODICE_CAPOSALA", "CODICE_MEDICO")
    v_Val = Array(numKey, oData(0).txtBox, turno, GetCodice("INFERMIERI", cboCaposala), GetCodice("MEDICI_DIALISI", cboMedico))
    
    Set rsPiano = New Recordset
    If modifica Then
        rsPiano.Open "SELECT * FROM PIANO_LAVORO WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsPiano.Update v_Nomi, v_Val
    Else
        rsPiano.Open "PIANO_LAVORO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsPiano.AddNew v_Nomi, v_Val
        rsPiano.Update
    End If
    rsPiano.Close
    
    If modifica Then
        ' elimina le precedenti associazioni
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
        cmCommand.CommandText = "DELETE * FROM ASSOCIAZIONI_PIANO_LAVORO WHERE CODICE_PIANO=" & keyId
        cmCommand.Execute
    End If
    
    rsPiano.Open "ASSOCIAZIONI_PIANO_LAVORO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsAppo.MoveFirst
    Do While Not rsAppo.EOF
        rsPiano.AddNew
        rsPiano("CODICE_INFERMIERE") = rsAppo("CODICE_INFERMIERE")
        rsPiano("CODICE_PAZIENTE") = rsAppo("CODICE_PAZIENTE")
        rsPiano("CODICE_PIANO") = numKey
        rsPiano.Update
        rsAppo.MoveNext
    Loop
    Set rsPiano = Nothing
End Sub

Private Sub Stampa()
    On Error GoTo gestione
    Dim SQLString As String

    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim tipostrTurno As String
    Dim nomeTurno As String
    Dim giorno As Integer
    Dim dataAppo As Date
    
    If optTurno(0).Value Then
        tipostrTurno = "AM_INIZIO"
        nomeTurno = "MATTINA"
    ElseIf optTurno(1).Value Then
        tipostrTurno = "PM_INIZIO"
        nomeTurno = "POMERIGGIO"
    Else
        tipostrTurno = "SR_INIZIO"
        nomeTurno = "SERALE"
    End If
    giorno = Weekday(CDate(oData(0).data), vbMonday)
    
    SQLString = "SHAPE APPEND " & _
                "   NEW adVarChar(10) AS POSTAZIONE_RENE, " & _
                "   NEW adVarChar(70) AS PAZIENTE, " & _
                "   NEW adVarChar(70) AS INFERMIERE, " & _
                "   NEW adVarChar(50) AS FILTRO, " & _
                "   NEW adVarChar(150) AS ANTICOAGULANTEconDOSI, " & _
                "   NEW adVarChar(100) AS SOL_DIALITICA, " & _
                "   NEW adVarChar(15) AS EPOconUI, " & _
                "   NEW adVarChar(15) AS DURATA, " & _
                "   NEW adLongVarChar AS MEDICINALI, " & _
                "   NEW adLongVarChar AS SOMMINISTRAZIONE "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        

    Set rsPiano = New Recordset
    Set rsDataset = New Recordset
    rsPiano.Open "SELECT    TURNI.CODICE_PAZIENTE, TURNI.CODICE_RENE, TURNI.AM_INIZIO" & giorno & ", TURNI.PM_INIZIO" & giorno & ", TURNI.PM_INIZIO" & giorno & ", " & _
                 "          PAZIENTI.KEY, PAZIENTI.COGNOME, PAZIENTI.NOME, APPARATI.KEY, APPARATI.POSTAZIONE " & _
                 "FROM      ((TURNI " & _
                 "          INNER JOIN PAZIENTI ON TURNI.CODICE_PAZIENTE = PAZIENTI.KEY) " & _
                 "          INNER JOIN APPARATI ON TURNI.CODICE_RENE = APPARATI.KEY) " & _
                 "WHERE     ((PAZIENTI.STATO=0 OR PAZIENTI.STATO=4) AND " & _
                 "          TURNI." & tipostrTurno & giorno & "<>"""")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsPiano.EOF And rsPiano.BOF) Then
        With rsMain
            Do While Not rsPiano.EOF
                .AddNew
                .Fields("POSTAZIONE_RENE") = rsPiano("POSTAZIONE")
                .Fields("PAZIENTE") = rsPiano("COGNOME") & " " & rsPiano("NOME")
                rsAppo.Filter = "CODICE_PAZIENTE=" & rsPiano("CODICE_PAZIENTE")
                If Not (rsAppo.EOF And rsAppo.BOF) Then
                    rsDataset.Open "SELECT * FROM INFERMIERI WHERE KEY=" & rsAppo("CODICE_INFERMIERE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                    .Fields("INFERMIERE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    rsDataset.Close
                End If
                
                Dim strSql As String
                strSql = "SELECT UI, EPO, TIPO_FILTRO, ANT.NOME, DOSE1, DOSE2, DOSE3, SOL_D.NOME, ORE, MINUTI, F.NOME, F.KEY " & _
                            "FROM (((ANAMNESI_DIALITICHE AN " & _
                            "       LEFT OUTER JOIN FILTRI F ON AN.TIPO_FILTRO=F.KEY) " & _
                            "       LEFT OUTER JOIN ANTICOAGULANTI ANT ON AN.ANTICOAGULANTE1=ANT.KEY) " & _
                            "       LEFT OUTER JOIN SOL_DIALITICHE SOL_D ON AN.SOL_DIALITICA=SOL_D.KEY) " & _
                            "WHERE AN.CODICE_PAZIENTE=" & rsPiano("PAZIENTI.KEY")
                            
                rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    .Fields("FILTRO") = rsDataset("F.NOME")
                    .Fields("ANTICOAGULANTEconDOSI") = IIf(rsDataset("ANT.NOME") = Null, "", rsDataset("ANT.NOME")) & vbCrLf & "Dose iniziale: " & rsDataset("DOSE1") & vbCrLf & "Dose intermedia: " & rsDataset("DOSE2") & vbCrLf & "Dose finale: " & rsDataset("DOSE3")
                    .Fields("EPOconUI") = IIf(rsDataset("EPO") <> -1, rsDataset("EPO") & " " & rsDataset("UI"), "")
                    .Fields("SOL_DIALITICA") = IIf(rsDataset("SOL_D.NOME") = Null, "", rsDataset("SOL_D.NOME"))
                    .Fields("DURATA") = rsDataset("ORE") & ":" & rsDataset("MINUTI")
                Else
                    .Fields("FILTRO") = "N/A"
                    .Fields("ANTICOAGULANTEconDOSI") = "N/A"
                    .Fields("EPOconUI") = "N/A"
                    .Fields("SOL_DIALITICA") = "N/A"
                    .Fields("DURATA") = "N/A"
                End If
                rsDataset.Close
                
                strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
                         "FROM      (TERAPIE_DIALITICHE " & _
                         "          INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
                         "WHERE     TERAPIE_DIALITICHE.CODICE_PAZIENTE=" & rsPiano("PAZIENTI.KEY") & " AND " & _
                         "          SOSPESA=FALSE " & _
                         "ORDER BY  TERAPIE_DIALITICHE.DATA DESC"
                rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    dataAppo = rsDataset("DATA")
                    Do While Not rsDataset.EOF
                        If rsDataset("DATA") <> dataAppo Then Exit Do
                        .Fields("MEDICINALI") = .Fields("MEDICINALI") & vbCrLf & rsDataset("MEDICINALINOME")
                        .Fields("SOMMINISTRAZIONE") = .Fields("SOMMINISTRAZIONE") & vbCrLf & Choose(rsDataset("SOMMINISTRAZIONE") + 1, "", "Intradialitica", "Postdialitica")
                        rsDataset.MoveNext
                    Loop
                Else
                    .Fields("MEDICINALI") = "N/A"
                    .Fields("SOMMINISTRAZIONE") = "N/A"
                End If
                rsDataset.Close
                rsPiano.MoveNext
            Loop
        End With
    Else
        MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
        Exit Sub
    End If
    
    rsAppo.Filter = ""
    Set rsDataset = Nothing
    Set rsPiano = Nothing
    
    rsMain.Sort = "POSTAZIONE_RENE"
    Set rptPianoGiornaliero.DataSource = rsMain
    With rptPianoGiornaliero.Sections("intestazione").Controls
        .Item("lblRagione").Caption = structIntestazione.sRagione
        .Item("lblTitolo").Caption = "PIANO DI LAVORO DEL GIORNO" & " " & oData(0).txtBox
        .Item("lblTurno").Caption = "TURNO: " & nomeTurno
    End With
    With rptPianoGiornaliero.Sections("pie").Controls
        .Item("lblMedico").Caption = cboMedico.Text
        .Item("lblCaposala").Caption = cboCaposala.Text
        .Item("lblDirettore").Caption = GetDirettore
        .Item("lblCompilataDa").Caption = tAccesso.cognome & " " & tAccesso.nome
    End With
    rptPianoGiornaliero.Orientation = rptOrientLandscape
    rptPianoGiornaliero.PrintReport True, rptRangeAllPages
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

' Restituisce il cognome nome del direttore sanitario
Private Function GetDirettore() As String
    Dim rsDataset As New Recordset
    rsDataset.Open "DIRETTORE_SANITARIO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        GetDirettore = rsDataset("COGNOME") & " " & rsDataset("NOME")
    Else
        GetDirettore = ""
    End If
    Set rsDataset = Nothing
End Function

Private Sub cmdSposta_Click(Index As Integer)
    Dim keyInf As Integer
    Dim Row As Integer
    
    keyInf = flxInfermieri.TextMatrix(flxInfermieri.Row, 0)
    Select Case Index
        Case 0  ' sposta il selezionato a destra
            Row = flxGrigliaSx.Row
            If Row = 0 Then Exit Sub
            rsAppo.Filter = "CODICE_INFERMIERE=" & keyInf & " AND CODICE_PAZIENTE=" & flxGrigliaSx.TextMatrix(Row, 0)
            rsAppo.Delete
            rsAppo.Filter = ""
            With flxGrigliaDx
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = flxGrigliaSx.TextMatrix(Row, 0)
                .TextMatrix(.Rows - 1, 1) = flxGrigliaSx.TextMatrix(Row, 1)
                .TextMatrix(.Rows - 1, 2) = flxGrigliaSx.TextMatrix(Row, 2)
                .TextMatrix(.Rows - 1, 3) = flxGrigliaSx.TextMatrix(Row, 3)
            End With
            flxGrigliaDx.Col = 1
            flxGrigliaDx.Sort = SortSettings.flexSortGenericAscending
            flxGrigliaDx.Col = 0
            If flxGrigliaSx.Rows = 2 Then
                flxGrigliaSx.Rows = 1
            Else
                flxGrigliaSx.RemoveItem (Row)
            End If
            
        Case 1  ' sposta il selezionato a sinistra
            If flxGrigliaSx.Rows = 5 Then
                MsgBox "L'infermiere selezionato è già associato a 4 pazienti", vbCritical, "Attenzione"
                Exit Sub
            End If
            Row = flxGrigliaDx.Row
            If Row = 0 Then Exit Sub
            rsAppo.AddNew
            rsAppo("CODICE_INFERMIERE") = keyInf
            rsAppo("CODICE_PAZIENTE") = flxGrigliaDx.TextMatrix(Row, 0)
            rsAppo.Update
            With flxGrigliaSx
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = flxGrigliaDx.TextMatrix(Row, 0)
                .TextMatrix(.Rows - 1, 1) = flxGrigliaDx.TextMatrix(Row, 1)
                .TextMatrix(.Rows - 1, 2) = flxGrigliaDx.TextMatrix(Row, 2)
                .TextMatrix(.Rows - 1, 3) = flxGrigliaDx.TextMatrix(Row, 3)
            End With
            flxGrigliaSx.Col = 1
            flxGrigliaSx.Sort = SortSettings.flexSortGenericAscending
            flxGrigliaSx.Col = 0
            If flxGrigliaDx.Rows = 2 Then
                flxGrigliaDx.Rows = 1
            Else
                flxGrigliaDx.RemoveItem (Row)
            End If
            
    End Select
    Call ColoraFlx(flxGrigliaDx, flxGrigliaDx.Cols - 1, True)
    Call ColoraFlx(flxGrigliaSx, flxGrigliaSx.Cols - 1, True)
    flxGrigliaDx.Row = 0
    flxGrigliaSx.Row = 0
    lblAssociati = "Pazienti associati: " & flxGrigliaSx.Rows - 1
    lblNonAssociati = "Pazienti non associati: " & flxGrigliaDx.Rows - 1
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdPulisci_Click()
    Call Pulisci
    flxInfermieri.Rows = 1
    cboCaposala.Clear
    cboMedico.Clear
    oData(0).Pulisci
    lblGiorno = ""
    modifica = False
End Sub

' Se ci sono pazienti associati salva anche
Private Sub cmdStampa_Click()
    If MsgBox("Stampare il piano di lavoro?", vbYesNo + vbQuestion, "Stampa") = vbYes Then
        If Completo Then
            If rsAppo.RecordCount = 0 Then
                Call Stampa
            Else
                Call Memorizza
                Call Stampa
            End If
        End If
    End If
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    If oData(0).data <> "" Then
        lblGiorno = UCase(WeekdayName(Weekday(CDate(oData(0).data), vbMonday), False, vbMonday))
        Call CaricaScheda
    End If
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
    cmdPulisci_Click
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGrigliaDx.hWnd Then
'        If flxGrigliaDx.TopRow - Rotation > 0 Then
'            If flxGrigliaDx.TopRow - Rotation < flxGrigliaDx.Rows Then
'                flxGrigliaDx.TopRow = flxGrigliaDx.TopRow - Rotation
'            End If
'        End If
'    ElseIf ControlHWnd = flxGrigliaSx.hWnd Then
'        If flxGrigliaSx.TopRow - Rotation > 0 Then
'            If flxGrigliaSx.TopRow - Rotation < flxGrigliaSx.Rows Then
'                flxGrigliaSx.TopRow = flxGrigliaSx.TopRow - Rotation
'            End If
'        End If
'    ElseIf ControlHWnd = flxInfermieri.hWnd Then
'        If flxInfermieri.TopRow - Rotation > 0 Then
'            If flxInfermieri.TopRow - Rotation < flxInfermieri.Rows Then
'                flxInfermieri.TopRow = flxInfermieri.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'------------------------

Private Sub flxInfermieri_Click()
    Dim i As Integer
    flxInfermieri.SetFocus
    If VerificaClickFlx(flxInfermieri) = False Then
        ' discolora
        Call ColoraFlx(flxInfermieri, flxInfermieri.Cols - 1, True)
        ' annulla le row e col
        flxInfermieri.Row = 0
        flxInfermieri.Col = 0
        flxGrigliaSx.Rows = 1
        For i = 0 To 1
            cmdSposta(i).Enabled = False
        Next i
    Else
        Call ColoraFlx(flxInfermieri, flxInfermieri.Cols - 1)
        Call CaricaPazienti(flxGrigliaSx, CreaCondizione())
        For i = 0 To 1
            cmdSposta(i).Enabled = True
        Next i
    End If
End Sub

Private Sub flxGrigliaSx_Click()
    flxGrigliaSx.SetFocus
    If VerificaClickFlx(flxGrigliaSx) = False Then
        ' discolora
        Call ColoraFlx(flxGrigliaSx, flxGrigliaSx.Cols - 1, True)
        ' annulla le row e col
        flxGrigliaSx.Row = 0
        flxGrigliaSx.Col = 0
    Else
        Call ColoraFlx(flxGrigliaSx, flxGrigliaSx.Cols - 1)
    End If
End Sub

Private Sub flxGrigliaSx_DblClick()
    If Not VerificaClickFlx(flxGrigliaSx) Then Exit Sub
    If flxInfermieri.Row <> 0 Then
        cmdSposta_Click (0)
    End If
End Sub

Private Sub flxGrigliaDx_Click()
    flxGrigliaDx.SetFocus
    If VerificaClickFlx(flxGrigliaDx) = False Then
        ' discolora
        Call ColoraFlx(flxGrigliaDx, flxGrigliaDx.Cols - 1, True)
        ' annulla le row e col
        flxGrigliaDx.Row = 0
        flxGrigliaDx.Col = 0
    Else
        Call ColoraFlx(flxGrigliaDx, flxGrigliaDx.Cols - 1)
    End If
End Sub

Private Sub flxGrigliaDx_DblClick()
    If Not VerificaClickFlx(flxGrigliaDx) Then Exit Sub
    If flxInfermieri.Row <> 0 Then
        cmdSposta_Click (1)
    End If
End Sub

Private Sub optTurno_Click(Index As Integer)
    Call ColoraSel(optTurno, Index, 3)
    If oData(0).data <> "" Then
        Call CaricaScheda
    End If
End Sub

