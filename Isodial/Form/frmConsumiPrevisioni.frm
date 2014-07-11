VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsumiPrevisioni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consumi e Previsioni"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12855
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
         ItemData        =   "frmConsumiPrevisioni.frx":0000
         Left            =   6120
         List            =   "frmConsumiPrevisioni.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboProdotto 
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
         ItemData        =   "frmConsumiPrevisioni.frx":0004
         Left            =   1320
         List            =   "frmConsumiPrevisioni.frx":001D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3735
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
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prodotto"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   12855
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         BackColorSel    =   16776960
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         FormatString    =   $"frmConsumiPrevisioni.frx":0077
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
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2775
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         BackColorSel    =   16776960
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         FormatString    =   $"frmConsumiPrevisioni.frx":014F
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
      Begin VB.Label lblPrevisioni 
         AutoSize        =   -1  'True
         Caption         =   "Previsioni per l'anno"
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
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consumi"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   12855
      Begin VB.CheckBox ChkPrev 
         Caption         =   "Stampa &PREVISIONI"
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
         TabIndex        =   14
         Top             =   520
         Width           =   2535
      End
      Begin VB.CheckBox ChkCons 
         Caption         =   "Stampa C&ONSUMI"
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
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Value           =   1  'Checked
         Width           =   2415
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
         Left            =   11280
         TabIndex        =   12
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
         Left            =   9600
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConsumiPrevisioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim v_giorni(1 To 7) As Integer
Private Type structTurni
    codicePaziente As Integer
    data_inizio_emodialisi As String
    data_fine_emodialisi As String
    totaleDialisi As Integer
    codiceProdotto As Integer
End Type
Dim v_Turni() As structTurni
Dim strNomeTabella As String
Dim strNomeCampo As String

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    
    For i = 0 To 3
        cboAnno.AddItem Year(Now) - i
    Next
    cboAnno.ListIndex = 0
    
    For j = 0 To 1
        With flxGriglia(j)
            .ColWidth(0) = 0
            .Row = 0
            For i = 1 To 14
                .Col = i
                .CellFontBold = True
            Next i
            .ColAlignment(1) = vbLeftJustify
            .Col = 0
        End With
    Next j
    lblPrevisioni = "Previsioni per l'anno " & Year(date)
End Sub

'' Calcola quanti giorni (lun, mar, merc...) ci sono in un mese
Private Sub CalcolaGiorni(mese As Integer, anno As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tipoGiorno As Integer
    
    For i = 1 To 7
        tipoGiorno = Weekday(DateValue(i & "/" & mese & "/" & anno), vbMonday)
        j = i
        v_giorni(tipoGiorno) = 0
        Do While Not j > Day(GetUltimoGiorno(mese, anno))
            v_giorni(tipoGiorno) = v_giorni(tipoGiorno) + 1
            j = j + 7
        Loop
        'Debug.Print tipoGiorno & " - " & v_giorni(tipoGiorno)
    Next i
End Sub

'' Calcola il num di dialisi da effettuare nel mese selezionato
' andando a controllare i turni
' @param key indice del paziente
Private Function GetNumeroDialisiFuture(key As Integer) As Integer
    Dim rsDataset As New Recordset
    Dim totale As Integer
    Dim i As Integer
    
    rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        For i = 1 To 7
            If rsDataset("AM_INIZIO" & i) <> "" Or rsDataset("PM_INIZIO" & i) <> "" Or rsDataset("SR_INIZIO" & i) <> "" Then
                totale = totale + v_giorni(i)
            End If
        Next i
    End If
    rsDataset.Close
    Set rsDataset = Nothing
    
    GetNumeroDialisiFuture = totale
End Function

Private Sub CalcolaPrevisioni()
    Dim i As Integer
    Dim j As Integer
    Dim somma As Integer
    Dim strSql As String
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    
    Select Case cboProdotto.ListIndex
        Case 0
            strNomeTabella = "AGO"
            strNomeCampo = "AGO1"
        Case 1
            strNomeTabella = "AGO"
            strNomeCampo = "AGO2"
        Case 2
            strNomeTabella = "CARTUCCE"
            strNomeCampo = "CARTUCCIA"
        Case 3
            strNomeTabella = "FILTRI"
            strNomeCampo = "TIPO_FILTRO"
        Case 4
            strNomeTabella = "LINEE"
            strNomeCampo = "TIPO_LINEE"
        Case 5
            strNomeTabella = "SOL_DIALITICHE"
            strNomeCampo = "SOL_DIALITICA"
        Case 6
            strNomeTabella = "SOL_INFUSIONALI"
            strNomeCampo = "SOL_INFUSIONALE"
    End Select
        
    flxGriglia(1).Rows = 1
    strSql = "SELECT     DISTINCT T.KEY, T.NOME " & _
            "FROM       ANAMNESI_DIALITICHE A " & _
            "           INNER JOIN " & strNomeTabella & " T ON T.KEY=A." & strNomeCampo & " " & _
            "ORDER BY   T.NOME"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia(1)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("NOME")
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    With flxGriglia(1)
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = 0
        .TextMatrix(.Rows - 1, 1) = "TOTALE"
    End With
    
    strSql = "SELECT    TURNI.CODICE_PAZIENTE, " & strNomeCampo & ", DATA_INIZIO, DATA_FINE " & _
            "FROM       (((TURNI " & _
            "           INNER JOIN ANAMNESI_DIALITICHE ON ANAMNESI_DIALITICHE.CODICE_PAZIENTE=TURNI.CODICE_PAZIENTE) " & _
            "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=TURNI.CODICE_PAZIENTE) " & _
            "           INNER JOIN ANAMNESI_NEFROLOGICHE ON ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
            "WHERE      (STATO=0) AND " & _
            "           NOT " & strNomeCampo & "=-1"
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    ReDim v_Turni(0)
    Do While Not rsDataset.EOF
        ReDim Preserve v_Turni(UBound(v_Turni) + 1)
        v_Turni(UBound(v_Turni)).codicePaziente = rsDataset("CODICE_PAZIENTE")
        v_Turni(UBound(v_Turni)).data_inizio_emodialisi = IIf(IsNull(rsDataset("DATA_INIZIO")), "", rsDataset("DATA_INIZIO"))
        v_Turni(UBound(v_Turni)).data_fine_emodialisi = IIf(IsNull(rsDataset("DATA_FINE")), "", rsDataset("DATA_FINE"))
        v_Turni(UBound(v_Turni)).codiceProdotto = rsDataset(strNomeCampo)
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    For i = 2 To 13
        Call CalcolaGiorni(i - 1, Year(Now))
        For j = 1 To UBound(v_Turni) - 1
            If v_Turni(j).data_inizio_emodialisi <> "" Then
                If v_Turni(j).data_fine_emodialisi <> "" Then
                    If CDate(v_Turni(j).data_inizio_emodialisi) <= DateValue("01/" & Format(i - 1, "00") & "/" & Year(Now)) And CDate(v_Turni(j).data_fine_emodialisi) >= GetUltimoGiorno(i - 1, Year(Now)) Then
                        v_Turni(j).totaleDialisi = GetNumeroDialisiFuture(v_Turni(j).codicePaziente)
                    Else
                        v_Turni(j).totaleDialisi = 0
                    End If
                Else
                    If CDate(v_Turni(j).data_inizio_emodialisi) <= DateValue("01/" & Format(i - 1, "00") & "/" & Year(Now)) Then
                        v_Turni(j).totaleDialisi = GetNumeroDialisiFuture(v_Turni(j).codicePaziente)
                    Else
                        v_Turni(j).totaleDialisi = 0
                    End If
                End If
            Else
                 v_Turni(j).totaleDialisi = 0
            End If
            frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
        Next j
        With flxGriglia(1)
            For j = 1 To .Rows - 2
                .TextMatrix(j, i) = "0"
                .TextMatrix(j, i) = Int(.TextMatrix(j, i)) + getNumeroPrevisioni(Int(.TextMatrix(j, 0)))
            Next j
        End With
    Next i
    
    With flxGriglia(1)
        For i = 1 To .Rows - 1
            .TextMatrix(i, 14) = 0
            For j = 2 To 13
                If .TextMatrix(i, j) <> "" Then
                    somma = .TextMatrix(i, j)
                Else
                    somma = 0
                End If
                .TextMatrix(i, 14) = Int(.TextMatrix(i, 14)) + somma
            Next j
        Next i
   
        For i = 2 To 14
            .TextMatrix(.Rows - 1, i) = 0
            For j = 1 To .Rows - 2
                If .TextMatrix(j, i) <> "" Then
                    somma = .TextMatrix(j, i)
                Else
                    somma = 0
                End If
                .TextMatrix(.Rows - 1, i) = Int(.TextMatrix(.Rows - 1, i)) + somma
            Next j
        Next i
    End With
End Sub

Private Function getNumeroPrevisioni(codiceProdotto As Integer) As Integer
    Dim i As Integer
    Dim somma As Integer
    
    For i = 1 To UBound(v_Turni)
        If v_Turni(i).codiceProdotto = codiceProdotto Then
            somma = somma + v_Turni(i).totaleDialisi
        End If
    Next i
    getNumeroPrevisioni = somma
End Function

Private Sub CalcolaConsumi()
    Dim i As Integer
    Dim j As Integer
    Dim somma As Integer
    Dim riga As Integer
    Dim strSql As String
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    
    Select Case cboProdotto.ListIndex
        Case 0
            strNomeTabella = "AGO"
            strNomeCampo = "TIPO_AGO1"
        Case 1
            strNomeTabella = "AGO"
            strNomeCampo = "TIPO_AGO2"
        Case 2
            strNomeTabella = "CARTUCCE"
            strNomeCampo = "CARTUCCIA"
        Case 3
            strNomeTabella = "FILTRI"
            strNomeCampo = "TIPO_FILTRO"
        Case 4
            strNomeTabella = "LINEE"
            strNomeCampo = "TIPO_LINEE"
        Case 5
            strNomeTabella = "SOL_DIALITICHE"
            strNomeCampo = "SOLUZIONE_DIALITICA"
        Case 6
            strNomeTabella = "SOL_INFUSIONALI"
            strNomeCampo = "SOLUZIONE_INFUSIONALE"
    End Select
    
    flxGriglia(0).Rows = 1
    strSql = "SELECT    T.KEY, T.NOME, Count(T.NOME) AS TOTALE " & _
             "FROM      ((SCHEDE_DIALISI " & _
             "          INNER JOIN STORICO_DIALISI_GIORNALIERA ON STORICO_DIALISI_GIORNALIERA.KEY=SCHEDE_DIALISI.CODICE_STORICO_DIALISI) " & _
             "          INNER JOIN " & strNomeTabella & " T ON T.NOME=STORICO_DIALISI_GIORNALIERA." & strNomeCampo & ")" & _
             " WHERE    ((YEAR([DATA])) = " & cboAnno.Text & " AND " & _
             "          ERRATA=FALSE) " & _
             "GROUP BY  T.NOME, T.KEY " & _
             "ORDER BY  T.NOME"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia(0)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("NOME")
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    With flxGriglia(0)
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = 0
        .TextMatrix(.Rows - 1, 1) = "TOTALE"
    End With
    
    For i = 1 To 12
        strSql = "SELECT    T.KEY, T.NOME, Count(T.NOME) AS TOTALE " & _
                 "FROM      ((SCHEDE_DIALISI " & _
                 "          INNER JOIN STORICO_DIALISI_GIORNALIERA ON STORICO_DIALISI_GIORNALIERA.KEY=SCHEDE_DIALISI.CODICE_STORICO_DIALISI) " & _
                 "          INNER JOIN " & strNomeTabella & " T ON T.NOME=STORICO_DIALISI_GIORNALIERA." & strNomeCampo & ")" & _
                 " WHERE    (((YEAR([DATA])) = " & cboAnno.Text & ") AND " & _
                 "          ((MONTH([DATA])) = " & i & ") AND " & _
                 "          ERRATA=FALSE) " & _
                 "GROUP BY  T.NOME, T.KEY"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataset.EOF
            With flxGriglia(0)
                riga = getRiga(rsDataset("KEY"))
                .TextMatrix(riga, i + 1) = rsDataset("TOTALE")
            End With
            rsDataset.MoveNext
            frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
        Loop
        rsDataset.Close
    Next i
    Set rsDataset = Nothing
    
    With flxGriglia(0)
        For i = 1 To .Rows - 1
            .TextMatrix(i, 14) = 0
            For j = 2 To 13
                If .TextMatrix(i, j) <> "" Then
                    somma = .TextMatrix(i, j)
                Else
                    somma = 0
                End If
                .TextMatrix(i, 14) = Int(.TextMatrix(i, 14)) + somma
            Next j
        Next i
    
        For i = 2 To 14
            .TextMatrix(.Rows - 1, i) = 0
            For j = 1 To .Rows - 2
                If .TextMatrix(j, i) <> "" Then
                    somma = .TextMatrix(j, i)
                Else
                    somma = 0
                End If
                .TextMatrix(.Rows - 1, i) = Int(.TextMatrix(.Rows - 1, i)) + somma
            Next j
        Next i
    End With
End Sub

'' Restituisce il numero di riga dove è presente il campo
Private Function getRiga(key As Integer) As Integer
    Dim i As Integer
    For i = 1 To flxGriglia(0).Rows - 1
        If flxGriglia(0).TextMatrix(i, 0) = key Then
            getRiga = i
            Exit Function
        End If
    Next i
End Function

Private Sub CalcolaPerProdotto()
    Dim strSql As String
    Dim rsDataset As New Recordset
    Dim intNumMax As Integer
    
    If cboProdotto.ListIndex = -1 Then Exit Sub
    
    Select Case cboProdotto.ListIndex
        Case 0
            strNomeTabella = "AGO"
            strNomeCampo = "AGO1"
        Case 1
            strNomeTabella = "AGO"
            strNomeCampo = "AGO2"
        Case 2
            strNomeTabella = "CARTUCCE"
            strNomeCampo = "CARTUCCIA"
        Case 3
            strNomeTabella = "FILTRI"
            strNomeCampo = "TIPO_FILTRO"
        Case 4
            strNomeTabella = "LINEE"
            strNomeCampo = "TIPO_LINEE"
        Case 5
            strNomeTabella = "SOL_DIALITICHE"
            strNomeCampo = "SOL_DIALITICA"
        Case 6
            strNomeTabella = "SOL_INFUSIONALI"
            strNomeCampo = "SOL_INFUSIONALE"
    End Select
    
    strSql = "SELECT    COUNT(T.KEY) AS TOTALE " & _
            "FROM       ANAMNESI_DIALITICHE A " & _
            "           INNER JOIN " & strNomeTabella & " T ON T.KEY=A." & strNomeCampo
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        intNumMax = rsDataset("TOTALE")
    End If
    rsDataset.Close
    Set rsDataset = Nothing

    flxGriglia(0).Rows = 1
    flxGriglia(1).Rows = 1
    
    If intNumMax > 0 Then
        Call StartProgressBar(intNumMax * 12, 0, Me)
        Call CalcolaConsumi
        Call CalcolaPrevisioni
        Call StopProgressBar(Me)
    End If
End Sub

Private Sub cboAnno_Click()
    Call CalcolaPerProdotto
End Sub

Private Sub cboProdotto_Click()
    Call CalcolaPerProdotto
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    
    Dim strSql As String
    Dim i As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If cboProdotto.Text = "" Then
        MsgBox "Selezionare il prodotto", vbInformation, "Attenzione"
        Exit Sub
    End If
 
    strSql = "SHAPE APPEND  " & _
                    "       NEW adVarChar (50) as NOME_CONSUMI, " & _
                    "       NEW adVarChar (10) as CONSUMI_GEN, " & _
                    "       NEW adVarChar (10) as CONSUMI_FEB, " & _
                    "       NEW adVarChar (10) as CONSUMI_MAR, " & _
                    "       NEW adVarChar (10) as CONSUMI_APR, " & _
                    "       NEW adVarChar (10) as CONSUMI_MAG, " & _
                    "       NEW adVarChar (10) as CONSUMI_GIU, " & _
                    "       NEW adVarChar (10) as CONSUMI_LUG, " & _
                    "       NEW adVarChar (10) as CONSUMI_AGO, " & _
                    "       NEW adVarChar (10) as CONSUMI_SET, " & _
                    "       NEW adVarChar (10) as CONSUMI_OTT, " & _
                    "       NEW adVarChar (10) as CONSUMI_NOV, " & _
                    "       NEW adVarChar (10) as CONSUMI_DIC, " & _
                    "       NEW adInteger  as TOTALE_CONSUMI "
                
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
    
    Set rptConsumiPrevisioni.DataSource = rsMain
    rptConsumiPrevisioni.Orientation = rptOrientLandscape
               
  If ChkCons.Value Then  ' stampa consumi
    With rsMain
        For i = 1 To flxGriglia(0).Rows - 1
        .AddNew
        .Fields("NOME_CONSUMI") = " " & flxGriglia(0).TextMatrix(i, 1)
        .Fields("CONSUMI_GEN") = " " & flxGriglia(0).TextMatrix(i, 2)
        .Fields("CONSUMI_FEB") = " " & flxGriglia(0).TextMatrix(i, 3)
        .Fields("CONSUMI_MAR") = " " & flxGriglia(0).TextMatrix(i, 4)
        .Fields("CONSUMI_APR") = " " & flxGriglia(0).TextMatrix(i, 5)
        .Fields("CONSUMI_MAG") = " " & flxGriglia(0).TextMatrix(i, 6)
        .Fields("CONSUMI_GIU") = " " & flxGriglia(0).TextMatrix(i, 7)
        .Fields("CONSUMI_LUG") = " " & flxGriglia(0).TextMatrix(i, 8)
        .Fields("CONSUMI_AGO") = " " & flxGriglia(0).TextMatrix(i, 9)
        .Fields("CONSUMI_SET") = " " & flxGriglia(0).TextMatrix(i, 10)
        .Fields("CONSUMI_OTT") = " " & flxGriglia(0).TextMatrix(i, 11)
        .Fields("CONSUMI_NOV") = " " & flxGriglia(0).TextMatrix(i, 12)
        .Fields("CONSUMI_DIC") = " " & flxGriglia(0).TextMatrix(i, 13)
        .Fields("TOTALE_CONSUMI") = " " & flxGriglia(0).TextMatrix(i, 14)
        Next i
      End With
      
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lbl").Caption = "CONSUMI"
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lblProdotto").Caption = cboProdotto.Text
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
      rptConsumiPrevisioni.PrintReport True, rptRangeAllPages

    End If
    
    rsMain.Close
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
    
    If ChkPrev.Value Then  ' stampa previsioni
      With rsMain
        For i = 1 To flxGriglia(0).Rows - 1
        .AddNew
        .Fields("NOME_CONSUMI") = " " & flxGriglia(1).TextMatrix(i, 1)
        .Fields("CONSUMI_GEN") = " " & flxGriglia(1).TextMatrix(i, 2)
        .Fields("CONSUMI_FEB") = " " & flxGriglia(1).TextMatrix(i, 3)
        .Fields("CONSUMI_MAR") = " " & flxGriglia(1).TextMatrix(i, 4)
        .Fields("CONSUMI_APR") = " " & flxGriglia(1).TextMatrix(i, 5)
        .Fields("CONSUMI_MAG") = " " & flxGriglia(1).TextMatrix(i, 6)
        .Fields("CONSUMI_GIU") = " " & flxGriglia(1).TextMatrix(i, 7)
        .Fields("CONSUMI_LUG") = " " & flxGriglia(1).TextMatrix(i, 8)
        .Fields("CONSUMI_AGO") = " " & flxGriglia(1).TextMatrix(i, 9)
        .Fields("CONSUMI_SET") = " " & flxGriglia(1).TextMatrix(i, 10)
        .Fields("CONSUMI_OTT") = " " & flxGriglia(1).TextMatrix(i, 11)
        .Fields("CONSUMI_NOV") = " " & flxGriglia(1).TextMatrix(i, 12)
        .Fields("CONSUMI_DIC") = " " & flxGriglia(1).TextMatrix(i, 13)
        .Fields("TOTALE_CONSUMI") = " " & flxGriglia(1).TextMatrix(i, 14)
        Next i
      End With
      
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lbl").Caption = "PREVISIONI"
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lblProdotto").Caption = cboProdotto.Text
      rptConsumiPrevisioni.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
      
      If ChkCons.Value Then  ' se stampa anche i consumi non visualizza il pannello di scelta stampante
          rptConsumiPrevisioni.PrintReport
      Else
          rptConsumiPrevisioni.PrintReport True, rptRangeAllPages
      End If
    End If


End Sub

Private Sub flxGriglia_Click(Index As Integer)
    If VerificaClickFlx(flxGriglia(Index)) = False Then
        ' annulla le row e col
        flxGriglia(Index).Row = 0
        flxGriglia(Index).Col = 0
    End If
End Sub

