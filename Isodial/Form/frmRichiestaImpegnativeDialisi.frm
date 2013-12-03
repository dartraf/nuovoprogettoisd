VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmRichiestaImpegnativeDialisi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Richiesta Impegnative Dialisi"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTempo 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3360
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   3
         ToolTipText     =   "Cerca data"
         Top             =   750
         Width           =   360
      End
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
         ItemData        =   "frmRichiestaImpegnativeDialisi.frx":0000
         Left            =   5040
         List            =   "frmRichiestaImpegnativeDialisi.frx":0002
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
         ItemData        =   "frmRichiestaImpegnativeDialisi.frx":0004
         Left            =   1200
         List            =   "frmRichiestaImpegnativeDialisi.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblAnno 
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
         Left            =   4320
         TabIndex        =   17
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblMese 
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
         Left            =   240
         TabIndex        =   16
         Top             =   250
         Width           =   585
      End
      Begin VB.Label lblDatadiStampa 
         AutoSize        =   -1  'True
         Caption         =   "Data di Stampa"
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
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   760
         Width           =   1785
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   795
         Width           =   1215
      End
   End
   Begin VB.Frame fraDati 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
      Begin VB.ComboBox cboStato 
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
         ItemData        =   "frmRichiestaImpegnativeDialisi.frx":009A
         Left            =   3840
         List            =   "frmRichiestaImpegnativeDialisi.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkTutti 
         Caption         =   "Stampa tutti i pazienti"
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
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   240
         Picture         =   "frmRichiestaImpegnativeDialisi.frx":009E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
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
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   4575
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
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   4575
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
         Index           =   2
         Left            =   840
         TabIndex        =   10
         Top             =   720
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
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   7095
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
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   1140
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
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
   End
   Begin MSComctlLib.ProgressBar prgBarra 
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Max             =   70
   End
   Begin VB.Frame fraPrestazioni 
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   7095
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   5280
         TabIndex        =   25
         Top             =   1320
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.ComboBox cboPrestazioni 
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
         Left            =   4800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.ComboBox cboCodici 
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
         Left            =   4800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3735
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   99
         FormatString    =   "| Paziente                                                       | N° Dialisi      | Cod. Prestaz.          "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmRichiestaImpegnativeDialisi.frx":04F7
      End
   End
End
Attribute VB_Name = "frmRichiestaImpegnativeDialisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim totaleDialisi As Integer            ' es: 3
Dim giorniDialisi As String             ' es: "1 - 7 - 21"
Dim v_giorni(1 To 7) As Integer
Dim vRow As Integer
Dim intPazientiKey As Integer

Private Sub Form_Load()
    Dim i As Integer
    
    lblData.BackColor = vbWhite
    
    Select Case tStampa
        Case tpIMPEGNATIVE
            Me.Caption = "Richiesta Impegnative Dialisi"
            fraPulsanti.ZOrder 1
            fraPulsanti.Top = 6720
            Me.Height = 8070
            Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboCodici)
            Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboPrestazioni)
            With flxGriglia
                .Rows = 1
                .ColWidth(0) = 0
                For i = 1 To 3
                    .Col = i
                    .CellFontBold = True
                Next i
            End With
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) + 1
    End Select
    
    cboAnno.ListIndex = 0
    lblData = date
    picData.Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
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
    If lblData = "" Then
        MsgBox "Inserisci la data", vbCritical, "Attenzione"
        Exit Function
    End If
    If chkTutti.Value = Unchecked And intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente ", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

Private Sub CaricaDialisi(evStr As String)
    Dim rsDialisi As Recordset
    Set rsDialisi = New Recordset
    Dim v_giorni() As Integer
    Dim i As Integer
    ' resetta le var
    giorniDialisi = ""
    totaleDialisi = 0
    ReDim v_giorni(0)
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI " & evStr & " AND ERRATA=FALSE AND Month([DATA])=" & cboMese.ListIndex + 1 & " AND Year([DATA])=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDialisi.EOF
        totaleDialisi = totaleDialisi + 1
        ReDim Preserve v_giorni(UBound(v_giorni) + 1)
        v_giorni(UBound(v_giorni)) = Day(rsDialisi("DATA"))
        rsDialisi.MoveNext
    Loop
    Call BubbleSort(v_giorni)
    For i = 1 To totaleDialisi
        giorniDialisi = giorniDialisi & v_giorni(i) & " - "
    Next i
    ' elimina il - finale
    If giorniDialisi <> "" Then
        giorniDialisi = Left(giorniDialisi, Len(giorniDialisi) - 3)
    End If
    Set rsDialisi = Nothing
End Sub

'' Calcola quanti giorni (lun, mar, merc...) ci sono in un mese
Private Sub CalcolaGiorni()
    Dim i As Integer
    Dim j As Integer
    Dim tipoGiorno As Integer
    
    For i = 1 To 7
        tipoGiorno = Weekday(DateValue(i & "/" & cboMese.ListIndex + 1 & "/" & cboAnno.Text), vbMonday)
        j = i
        v_giorni(tipoGiorno) = 0
        Do While Not j > Day(GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text))
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

Private Sub CaricaFlx()
    Dim rsPazienti As New Recordset
    Dim strSingoloPaziente As String
    
    If cboMese.ListIndex = -1 Then Exit Sub
    
    Call CalcolaGiorni
    flxGriglia.Rows = 1
    If chkTutti.Value = Unchecked Then
        strSingoloPaziente = " AND KEY=" & intPazientiKey
    End If
    
    rsPazienti.Open "SELECT NOME, COGNOME, KEY FROM PAZIENTI WHERE (STATO=0) " & strSingoloPaziente & " ORDER BY COGNOME, NOME", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsPazienti.EOF
 
         If rsPazienti.RecordCount = 1 And GetNumeroDialisiFuture(rsPazienti("KEY")) = 0 Then
            MsgBox "IMPOSSIBILE ELABORARE IL NUMERO DELLE IMPEGNATIVE" & vbCrLf & "Non risulta compilata l'anamnesi dialitica oppure al paziente non è stato attribuito alcun turno dialitico", vbCritical, "ATTENZIONE!!!"
            rsPazienti.Close
            Set rsPazienti = Nothing
            Exit Sub
         ElseIf GetNumeroDialisiFuture(rsPazienti("KEY")) = 0 Then
 '    elimina pazienti con turni di dialisi non definiti
         Else
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsPazienti("KEY")
                .TextMatrix(.Rows - 1, 1) = rsPazienti("COGNOME") & " " & rsPazienti("NOME")
                .TextMatrix(.Rows - 1, 2) = GetNumeroDialisiFuture(rsPazienti("KEY"))
                .TextMatrix(.Rows - 1, 3) = CaricaCodicePrestazione(rsPazienti("KEY"))
            End With
         End If
        rsPazienti.MoveNext
    Loop
    rsPazienti.Close
    Set rsPazienti = Nothing
End Sub

Private Function CaricaCodicePrestazione(key As Integer) As String
    Dim rsPazienti As New Recordset
    Dim CodicePrestazione As String

    rsPazienti.Open "SELECT * FROM ANAMNESI_DIALITICHE WHERE CODICE_PAZIENTE =" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsPazienti.EOF
            CodicePrestazione = rsPazienti("CODICE_PRESTAZIONE") & " "
            rsPazienti.MoveNext
        Loop
    rsPazienti.Close
    Set rsPazienti = Nothing
    
    If CodicePrestazione = " " Then
        cboCodici.ListIndex = 4 - 1                 ' meno 1 perchè sballa di una posizione selezionando il codice della prestazione successiva
        CaricaCodicePrestazione = cboCodici.Text
    Else
        cboCodici.ListIndex = CodicePrestazione - 1 ' meno 1 perchè sballa di una posizione selezionando il codice della prestazione successiva
        CaricaCodicePrestazione = cboCodici.Text
    End If
    
End Function

Private Sub cmdTrova_Click()
    tTrova.Tipo = tpPAZIENTE
    If tStampa = tpKTVANNUALE Or tStampa = tpTSATANNUALE Then
        tTrova.condStato = "(-1)"
        If cboStato.ListIndex = cboStato.ListCount - 1 Then
            tTrova.condizione = " NOT STATO=-1 "
        Else
            tTrova.condizione = "STATO= " & cboStato.ItemData(cboStato.ListIndex)
        End If
    Else
        tTrova.condStato = ""
        tTrova.condizione = ""
    End If
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        If intPazientiKey = tTrova.keyReturn Then
            Call CaricaFlx
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsPazienti As Recordset
    Dim rsAppo As Recordset
    
    Dim strSingoloPaziente As String
    Dim i As Integer

    If Completo Then
        If chkTutti.Value = Unchecked Then
            strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
        End If

        strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(100) AS PAZIENTE, " & _
                "       NEW adInteger AS TOTALE_DIALISI, " & _
                "       NEW adVarChar(100) AS TIPO_DIALISI, " & _
                "       NEW adVarChar(10) AS CODICE_NOM, " & _
                "       NEW adVarChar(10) AS CODICE_ESENZIONE "
        
        ' apre la connessione per lo shape
        Set cnConn = New ADODB.Connection
        cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
        Set rsMain = New ADODB.Recordset
        rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic

        Set rsPazienti = New Recordset
        Set rsAppo = New Recordset
        
        Call CalcolaGiorni
        
        i = 0
        strSql = "SELECT    PAZIENTI.NOME, PAZIENTI.COGNOME, PAZIENTI.KEY, CODICE " & _
                "FROM       ((PAZIENTI " & _
                "           INNER JOIN TURNI ON TURNI.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
                "           LEFT OUTER JOIN TIPOLOGIE_ESENZIONE ON PAZIENTI.CODICE_ESENZIONE=TIPOLOGIE_ESENZIONE.KEY) " & _
                "WHERE      (STATO=0) " & _
                strSingoloPaziente & " " & _
                "ORDER BY   PAZIENTI.COGNOME, PAZIENTI.NOME"
        rsPazienti.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        Do While Not rsPazienti.EOF
            i = i + 1
            With rsMain
                .AddNew
                .Fields("PAZIENTE") = rsPazienti("COGNOME") & " " & rsPazienti("NOME")
                .Fields("TOTALE_DIALISI") = flxGriglia.TextMatrix(i, 2)
                
                rsAppo.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE CODICE='" & flxGriglia.TextMatrix(i, 3) & "'", cnPrinc, adOpenForwardOnly, adLockOptimistic, adCmdText
                 .Fields("TIPO_DIALISI") = rsAppo("NOME")
                rsAppo.Close

                .Fields("CODICE_NOM") = flxGriglia.TextMatrix(i, 3)
                .Fields("CODICE_ESENZIONE") = rsPazienti("CODICE")
                .Update
            End With
            rsPazienti.MoveNext
        Loop
        Set rsAppo = Nothing
        Set rsPazienti = Nothing
        
        If rsMain.RecordCount = 0 Then
            MsgBox "Nessuna dialisi per il mese di " & cboMese, vbCritical, Me.Caption
            Exit Sub
        End If
        
        Set rptImpegnativeDialisi.DataSource = rsMain
        rptImpegnativeDialisi.Sections("corpo").Controls.Item("lblMeseAnno").Caption = cboMese & " " & cboAnno
        rptImpegnativeDialisi.Sections("corpo").Controls.Item("lblLì").Caption = structIntestazione.sCitta & " lì, " & lblData
        rptImpegnativeDialisi.PrintReport True, rptRangeAllPages
    End If
    
End Sub

Private Sub chkTutti_Click()
    If chkTutti.Value = Checked Then
        intPazientiKey = 0
        lblCognome = ""
        lblNome = ""
        Call CaricaFlx
    End If
End Sub

Private Sub cboMese_Click()
    If tStampa = tpIMPEGNATIVE Then
        Call CaricaFlx
    End If
End Sub

Private Sub cboPrestazioni_Click()
    flxGriglia.TextMatrix(vRow, 3) = cboPrestazioni.Text
    cboPrestazioni.Visible = False
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
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If flxGriglia.Rows = 1 Then Exit Sub
    If flxGriglia.Row = flxGriglia.Rows - 1 Then
        i = 1
    Else
        i = flxGriglia.Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia.TextMatrix(i, 1), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia.Row = i
            If i >= 10 Or flxGriglia.TopRow > 10 Then
                flxGriglia.TopRow = i
            End If
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Do
        End If
        If i = flxGriglia.Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxGriglia.Row
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
    If cboPrestazioni.Visible Then
        cboPrestazioni.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub flxGriglia_DblClick()
    ' fase di modifica
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        .SetFocus
        If .Col = 2 Then
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        ElseIf .Col = 3 Then
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            cboPrestazioni.Left = .colPos(.Col) + .Left + 45
            cboPrestazioni.Top = .rowPos(.Row) + .Top + 45
            cboPrestazioni.ListIndex = GetIndex(cboPrestazioni, .TextMatrix(.Row, .Col))
            cboPrestazioni.Visible = True
            cboPrestazioni.SetFocus
        End If
    End With
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
    chkTutti.Value = Unchecked
    Call CaricaFlx
End Sub

Private Sub lblData_Click()
    lblData = ""
    laData = ""
End Sub

Private Sub picData_Click()
    frmCalendario.Show 1
    If laData <> "" Then lblData = laData
End Sub

Private Sub picData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case vbKeyReturn
            flxGriglia.SetFocus
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtAppo_LostFocus()
    txtAppo.Visible = False
    If (flxGriglia.TextMatrix(vRow, 2)) <> (txtAppo) Then
        If txtAppo = "" Then
            MsgBox "Impossibile memorizzare dati vuoti", vbCritical, "Attenzione"
            flxGriglia.Row = vRow
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Sub
        End If
        flxGriglia.TextMatrix(vRow, 2) = (txtAppo.Text)
    End If
End Sub

Private Sub WheelCatcher1_WheelRotation(Rotation As Long, X As Long, Y As Long, CtrlHwnd As Long)
On Error GoTo gestione
' se NON è stata selezionata una riga esce e NON attiva lo scroll
'    If flxGriglia.Row = 0 Then
'       Exit Sub
'    End If

    Select Case CtrlHwnd

        Case flxGriglia.hWnd
            If flxGriglia.TopRow - Rotation > 0 Then
               flxGriglia.TopRow = flxGriglia.TopRow - Rotation
            End If
    
        End Select
' Evita crash in caso di griglia non completa
gestione:
End Sub



