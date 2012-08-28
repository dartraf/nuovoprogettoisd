VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStampaFogliViaggio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa "
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBarra 
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Max             =   70
   End
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
         ItemData        =   "frmStampaFatt.frx":0000
         Left            =   5040
         List            =   "frmStampaFatt.frx":0002
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
         ItemData        =   "frmStampaFatt.frx":0004
         Left            =   1200
         List            =   "frmStampaFatt.frx":002C
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
         ItemData        =   "frmStampaFatt.frx":009A
         Left            =   3840
         List            =   "frmStampaFatt.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   24
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
         Picture         =   "frmStampaFatt.frx":009E
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
   Begin VB.Frame fraPrestazioni 
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   7095
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
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
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
         TabIndex        =   22
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
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3135
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5530
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
         MouseIcon       =   "frmStampaFatt.frx":04F7
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Prestazione"
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmStampaFogliViaggio"
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
        Case tpFOGLIOVIAGGIO
            Me.Caption = Me.Caption & "fogli di viaggio"
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) - 1
            
        Case tpIMPEGNATIVE
            Me.Caption = "Richiesta Impegnative Dialisi"
            fraPulsanti.ZOrder 1
            fraPulsanti.Top = 6720
            Me.Height = 8070
            Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboCodici)
            Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboPrestazioni)
            cboCodici.ListIndex = GetIndex(cboCodici, "39.95.4")
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
            
        Case tpMODULOFIRMEPAZIENTE
            Me.Caption = "Modulo Firme Paziente"
            cboMese.ListIndex = Month(Now) - 1
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) + 1
            
        Case tpSCHEDADIALITICASETTIMANALE
            Me.Caption = "Scheda Dialitica Settimanale"
            cboMese.ListIndex = Month(Now) - 1
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) + 1
            fraTempo.Height = 735
            fraDati.Top = 600
            fraPulsanti.Top = 2160
            prgBarra.Top = 3000
            Me.Height = 3345
            lblDatadiStampa(5).Visible = False
            lblData.Visible = False
            picData.Visible = False
            
        Case tpKTVANNUALE, tpTSATANNUALE
            cboStato.Visible = True
            Call RicaricaComboBox("TIPO_STATO", "NOME", cboStato)
            cboStato.AddItem "Tutti"
            cboStato.ItemData(cboStato.NewIndex) = 0
            cboStato.ListIndex = 0
            fraTempo.Height = 735
            lblAnno.Left = lblMese.Left
            cboAnno.Left = cboMese.Left
            lblMese.Visible = False
            cboMese.Visible = False
            fraDati.Top = fraTempo.Top + fraTempo.Height - 100
            fraPulsanti.Top = fraDati.Top + fraDati.Height - 110
            Me.Height = fraPulsanti.Top + fraPulsanti.Height + 370
            Me.Caption = Me.Caption & IIf(tStampa = tpKTVANNUALE, "Kt/V annuale", "TSAT% annuale")
            For i = 0 To 5
                cboAnno.AddItem Year(Now) - i
            Next i
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

Private Sub StampaFogliViaggio()
    Dim strSqlStampa As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsPazienti As Recordset
    
    Dim strSingola As String
    Dim strSingoloPaziente As String

    If Completo Then
        If chkTutti.Value = Unchecked Then
            strSingola = "WHERE CODICE_PAZIENTE=" & intPazientiKey
            strSingoloPaziente = "WHERE PAZIENTI.KEY=" & intPazientiKey
        End If
        
        strSqlStampa = "SHAPE APPEND " & _
                "       NEW adVarChar(100) AS PAZIENTE, " & _
                "       NEW adVarChar(100) AS DOMICILIO, " & _
                "       NEW adVarChar(50) AS INDIRIZZO, " & _
                "       NEW adVarChar(15) AS MESE, " & _
                "       NEW adVarChar(4) AS ANNO, " & _
                "       NEW adDate AS DATA, " & _
                "       NEW adInteger AS TOTALE_DIALISI, " & _
                "       NEW adVarChar(100) AS GIORNI_DIALISI, " & _
                "       NEW adVarChar(25) AS ASL, " & _
                "       NEW adVarChar(6) AS DISTRETTO "
        
        ' apre la connessione per lo shape
        Set cnConn = New ADODB.Connection
        cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
        Set rsMain = New ADODB.Recordset
        rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
        Set rsPazienti = New Recordset
        strSql = "SELECT        PAZIENTI.COGNOME, PAZIENTI.NOME AS PAZIENTINOME, PAZIENTI.KEY, PROV_RESIDENZA, INDIRIZZO, COMUNI.NOME AS COMUNINOME, ASL.NOME AS ASLNOME, DISTRETTI.NOME AS DISTRETTINOME " & _
                    " FROM      (((PAZIENTI " & _
                    "           LEFT OUTER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) " & _
                    "           LEFT OUTER JOIN ASL ON ASL.KEY=PAZIENTI.CODICE_ASL) " & _
                    "           LEFT OUTER JOIN DISTRETTI ON DISTRETTI.KEY=PAZIENTI.CODICE_DISTRETTO) " & _
                    strSingoloPaziente & " " & _
                    "ORDER BY    COGNOME, PAZIENTI.NOME"
        
        rsPazienti.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rsPazienti.EOF And rsPazienti.BOF) Then
            Do While Not rsPazienti.EOF
                With rsMain
                    Call CaricaDialisi(IIf(strSingola <> "", strSingola, "WHERE CODICE_PAZIENTE=" & rsPazienti("KEY")))
                    If totaleDialisi <> 0 Then
                        .AddNew
                        .Fields("PAZIENTE") = rsPazienti("COGNOME") & " " & rsPazienti("PAZIENTINOME")
                        .Fields("DOMICILIO") = rsPazienti("COMUNINOME") & " (" & rsPazienti("PROV_RESIDENZA") & ") "
                        .Fields("INDIRIZZO") = rsPazienti("INDIRIZZO")
                        .Fields("MESE") = cboMese.Text
                        .Fields("ANNO") = cboAnno.Text
                        .Fields("DATA") = lblData
                        .Fields("ASL") = rsPazienti("ASLNOME")
                        .Fields("DISTRETTO") = rsPazienti("DISTRETTINOME")
                        .Fields("TOTALE_DIALISI") = totaleDialisi
                        .Fields("GIORNI_DIALISI") = giorniDialisi
                        .Update
                    End If
                End With
                rsPazienti.MoveNext
            Loop
            Set rsPazienti = Nothing
            
            If rsMain.RecordCount = 0 Then
                MsgBox "Nessuna dialisi per il mese di " & cboMese, vbCritical, Me.Caption
                Exit Sub
            End If
                        
            Set rptFogliViaggio.DataSource = rsMain
            rptFogliViaggio.RightMargin = rptFogliViaggio.RightMargin / 3
            
            rptFogliViaggio.Sections("pie").Controls.Item("lblLuogo").Caption = structIntestazione.sCitta & " lì, " & GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
            rptFogliViaggio.PrintReport True, rptRangeAllPages
        Else
            MsgBox "Nessuna dialisi per il mese di " & cboMese, vbCritical, Me.Caption
        End If
    End If
End Sub

Private Sub StampaImpegnativeDialisi()
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
    
 '    elimina paziente con zero dialisi
        If GetNumeroDialisiFuture(rsPazienti("KEY")) = 0 Then
            rsPazienti.MoveNext
        End If
    
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsPazienti("KEY")
            .TextMatrix(.Rows - 1, 1) = rsPazienti("COGNOME") & " " & rsPazienti("NOME")
            .TextMatrix(.Rows - 1, 2) = GetNumeroDialisiFuture(rsPazienti("KEY"))
            .TextMatrix(.Rows - 1, 3) = cboCodici.Text
        End With
        rsPazienti.MoveNext
    Loop
    rsPazienti.Close
    Set rsPazienti = Nothing
End Sub

Private Sub StampaModuloFirmePaziente()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsPazienti As Recordset
    
    Dim strSingoloPaziente As String

    If Completo Then
        If chkTutti.Value = Unchecked Then
            strSingoloPaziente = " AND KEY=" & intPazientiKey
        End If

        SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(110) AS ASSISTITO "
        
        ' apre la connessione per lo shape
        Set cnConn = New ADODB.Connection
        cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
        Set rsMain = New ADODB.Recordset
        rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic

        Set rsPazienti = New Recordset

        rsPazienti.Open "SELECT COGNOME,NOME, DATA_NASCITA, STATO FROM PAZIENTI WHERE (STATO=0) " & strSingoloPaziente & " ORDER BY COGNOME, NOME", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        Do While Not rsPazienti.EOF
            With rsMain
                .AddNew
                .Fields("ASSISTITO") = rsPazienti("COGNOME") & " " & rsPazienti("NOME") & " - " & rsPazienti("DATA_NASCITA")
                .Update
            End With
            rsPazienti.MoveNext
        Loop
        Set rsPazienti = Nothing
        
        If rsMain.RecordCount <> 0 Then
            Set rptModuloFirmePaziente.DataSource = rsMain
            rptModuloFirmePaziente.LeftMargin = 1100
            rptModuloFirmePaziente.RightMargin = 0
            rptModuloFirmePaziente.TopMargin = 0
            rptModuloFirmePaziente.Sections("intestazione").Controls.Item("lblMese").Caption = cboMese & " " & cboAnno
            rptModuloFirmePaziente.Sections("pie").Controls.Item("lblStampato").Caption = structIntestazione.sCitta & " lì, " & lblData
            rptModuloFirmePaziente.PrintReport True, rptRangeAllPages
        End If
    End If
End Sub

Private Sub StampaKtvAnnuale()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim cont As Integer
    Dim i As Integer

    If chkTutti.Value = Unchecked And intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente ", vbCritical, "Attenzione"
        Exit Sub
    End If
    If chkTutti.Value = Unchecked Then
        strSingoloPaziente = " AND KEY=" & intPazientiKey
    End If
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If

    SQLString = "SHAPE APPEND " & _
            "       NEW adInteger AS CODICE_PAZIENTE, " & _
            "       NEW adVarChar(35) AS COGNOME, " & _
            "       NEW adVarChar(35) AS NOME, " & _
            "       NEW adCurrency AS MESE1, " & _
            "       NEW adCurrency AS MESE2, " & _
            "       NEW adCurrency AS MESE3, " & _
            "       NEW adCurrency AS MESE4, " & _
            "       NEW adCurrency AS MESE5, " & _
            "       NEW adCurrency AS MESE6, " & _
            "       NEW adCurrency AS MESE7, " & _
            "       NEW adCurrency AS MESE8, " & _
            "       NEW adCurrency AS MESE9, " & _
            "       NEW adCurrency AS MESE10, " & _
            "       NEW adCurrency AS MESE11, " & _
            "       NEW adCurrency AS MESE12, " & _
            "       NEW adCurrency AS MEDIA "

    
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic

    Set rsDataset = New Recordset
    rsDataset.Open "SELECT KEY,COGNOME,NOME, STATO FROM PAZIENTI WHERE " & strStato & " " & strSingoloPaziente & " ORDER BY COGNOME, NOME, KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("CODICE_PAZIENTE") = rsDataset("KEY")
            .Fields("COGNOME") = rsDataset("COGNOME")
            .Fields("NOME") = rsDataset("NOME")
            .Update
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM KTV WHERE ANNO=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        rsMain.Filter = "CODICE_PAZIENTE=" & rsDataset("CODICE_PAZIENTE")
        If rsMain.RecordCount <> 0 Then
            rsMain("MEDIA") = 0
            If IsNull(rsDataset("UREA_POST")) Or IsNull(rsDataset("UREA_PRE")) Or IsNull(rsDataset("DURATA")) Or IsNull(rsDataset("VOLUME")) Or IsNull(rsDataset("PESO")) Then
                rsMain.Fields("MESE" & rsDataset("MESE")) = Null
            Else
                rsMain.Fields("MESE" & rsDataset("MESE")) = CalcolaKtv(rsDataset("UREA_POST"), rsDataset("UREA_PRE"), rsDataset("DURATA"), rsDataset("VOLUME"), rsDataset("PESO"))
            End If
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    rsMain.Filter = ""
    rsMain.MoveFirst
    
    Do While Not rsMain.EOF
        cont = 0
        For i = 1 To 12
            If Not IsNull(rsMain("MESE" & i)) Then
                rsMain("MEDIA") = rsMain("MEDIA") + rsMain("MESE" & i)
                cont = cont + 1
            End If
        Next i
        If cont <> 0 Then
            rsMain("MEDIA") = rsMain("MEDIA") / cont
        Else
            rsMain("MEDIA") = Null
        End If
        rsMain.MoveNext
    Loop
    
    
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatAnnuale.DataSource = rsMain
        rptKtvTsatAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "KT/V ANNO " & cboAnno.Text
        rptKtvTsatAnnuale.LeftMargin = 500
        rptKtvTsatAnnuale.RightMargin = 0
        rptKtvTsatAnnuale.TopMargin = 0
        rptKtvTsatAnnuale.PrintReport True, rptRangeAllPages
    End If
End Sub

Private Sub StampaTsatAnnuale()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim cont As Integer
    Dim i As Integer

    If chkTutti.Value = Unchecked And intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente ", vbCritical, "Attenzione"
        Exit Sub
    End If
    If chkTutti.Value = Unchecked Then
        strSingoloPaziente = " AND KEY=" & intPazientiKey
    End If
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If

    SQLString = "SHAPE APPEND " & _
            "       NEW adInteger AS CODICE_PAZIENTE, " & _
            "       NEW adVarChar(35) AS COGNOME, " & _
            "       NEW adVarChar(35) AS NOME, " & _
            "       NEW adCurrency AS MESE1, " & _
            "       NEW adCurrency AS MESE2, " & _
            "       NEW adCurrency AS MESE3, " & _
            "       NEW adCurrency AS MESE4, " & _
            "       NEW adCurrency AS MESE5, " & _
            "       NEW adCurrency AS MESE6, " & _
            "       NEW adCurrency AS MESE7, " & _
            "       NEW adCurrency AS MESE8, " & _
            "       NEW adCurrency AS MESE9, " & _
            "       NEW adCurrency AS MESE10, " & _
            "       NEW adCurrency AS MESE11, " & _
            "       NEW adCurrency AS MESE12, " & _
            "       NEW adCurrency AS MEDIA "

    
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic

    Set rsDataset = New Recordset
    rsDataset.Open "SELECT KEY,COGNOME,NOME,STATO FROM PAZIENTI WHERE " & strStato & " " & strSingoloPaziente & " ORDER BY COGNOME, NOME, KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("CODICE_PAZIENTE") = rsDataset("KEY")
            .Fields("COGNOME") = rsDataset("COGNOME")
            .Fields("NOME") = rsDataset("NOME")
            .Update
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM TSAT WHERE ANNO=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        rsMain.Filter = "CODICE_PAZIENTE=" & rsDataset("CODICE_PAZIENTE")
        If rsMain.RecordCount <> 0 Then
            rsMain("MEDIA") = 0
            If IsNull(rsDataset("SIDEREMIA")) Or IsNull(rsDataset("TRANSFERRINA")) Then
                rsMain.Fields("MESE" & rsDataset("MESE")) = Null
            Else
                rsMain.Fields("MESE" & rsDataset("MESE")) = CalcolaTsat(rsDataset("SIDEREMIA"), rsDataset("TRANSFERRINA"))
            End If
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    rsMain.Filter = ""
    rsMain.MoveFirst
    
    Do While Not rsMain.EOF
        cont = 0
        For i = 1 To 12
            If Not IsNull(rsMain("MESE" & i)) Then
                rsMain("MEDIA") = rsMain("MEDIA") + rsMain("MESE" & i)
                cont = cont + 1
            End If
        Next i
        If cont <> 0 Then
            rsMain("MEDIA") = rsMain("MEDIA") / cont
        Else
            rsMain("MEDIA") = Null
        End If
        rsMain.MoveNext
    Loop
    
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatAnnuale.DataSource = rsMain
        rptKtvTsatAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "TSAT% ANNO " & cboAnno.Text
        rptKtvTsatAnnuale.LeftMargin = 500
        rptKtvTsatAnnuale.RightMargin = 0
        rptKtvTsatAnnuale.TopMargin = 0
        rptKtvTsatAnnuale.PrintReport True, rptRangeAllPages
    End If
End Sub

Private Sub StampaSchedaDialiticaSettimanale()
    Dim strSqlStampa As String
    Dim strSql As String
    Dim i As Integer
    Dim intNumeroGiorni As Integer
    Dim strGiorni As String
    Dim intCodiceID As Integer
    Dim valore As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsAppo As Recordset
    Dim rsAppo2 As Recordset
    Dim rsDialitiche As Recordset
    Dim rsDomiciliari As Recordset
    Dim rsEsami As Recordset
    
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim incrementa As Integer

    If chkTutti.Value = Checked Then
        prgBarra.Value = 0
        prgBarra.Visible = True
        Me.Height = 3720
    End If
    
    If chkTutti.Value = Unchecked And intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente ", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If chkTutti.Value = Unchecked Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
        
    strSqlStampa = "    NEW adVarChar (50) as COGNOME, " & _
                "       NEW adVarChar (50) as NOME, " & _
                "       NEW adInteger AS ANNI, " & _
                "       NEW adInteger AS CODICE_PAZIENTE, " & _
                "       NEW adVarChar (20) as EMOGRUPPO, " & _
                "       NEW adVarChar (50) as ACCESSO_VASCOLARE, " & _
                "       NEW adVarChar (150) as EDTA, " & _
                "       NEW adVarChar (50) as DURATA_SEDUTA, " & _
                "       NEW adSingle as PESO_SECCO, " & _
                "       NEW adSingle as PESO_DIALITICO_PRECEDENTE, " & _
                "       NEW adVarChar (20) as PRESSIONE_PRECEDENTE, " & _
                "       NEW adVarChar (20) as FREQUENZA_PRECEDENTE, " & _
                "       NEW adVarChar (50) AS FILTRO, " & _
                "       NEW adVarChar (70) AS FILTRO_SEDUTA, " & _
                "       NEW adVarChar (50) as BAGNO_DIALISI, " & _
                "       NEW adVarChar (50) AS ANTICOAGULANTE, " & _
                "       NEW adLongVarChar as ESAME1, "
    strSqlStampa = strSqlStampa & _
                "       NEW adLongVarChar as ESAME2, " & _
                "       NEW adLongVarChar as ESAME3, " & _
                "       NEW adLongVarChar as ESAME4, " & _
                "       NEW adLongVarChar as ESAME5, " & _
                "       NEW adLongVarChar as ESAME6, " & _
                "       NEW adLongVarChar as ESAME7, " & _
                "       NEW adLongVarChar as ESAME8, " & _
                "       NEW adLongVarChar as ESAME9, " & _
                "       NEW adLongVarChar as ESAME10, " & _
                "       NEW adLongVarChar as ESAME11, " & _
                "       NEW adLongVarChar as ESAME12, " & _
                "       NEW adLongVarChar as ESAME13, " & _
                "       NEW adLongVarChar as ESAME14, " & _
                "       NEW adLongVarChar as ESAME15, " & _
                "       NEW adLongVarChar as ESAME16, " & _
                "       NEW adLongVarChar as FARMACO_DIALISI, " & _
                "       NEW adLongVarChar as POSOLOGIA_DIALISI, " & _
                "       NEW adLongVarChar as NOTE_DIALISI, " & _
                "       NEW adLongVarChar as GIORNI_DIALISI, " & _
                "       NEW adLongVarChar as FARMACO_DOMICILIARE, " & _
                "       NEW adLongVarChar as POSOLOGIA_DOMICILIARE, " & _
                "       NEW adLongVarChar as NOTE_DOMICILIARE, " & _
                "       NEW adLongVarChar as GIORNI_DOMICILIARE"
    
    ' stringa di shape
    strSqlStampa = "SHAPE APPEND " & strSqlStampa
     
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    ' carica il recordset padre
    Set rsDataset = New Recordset
    Set rsAppo = New Recordset

    With rsMain
        
        ' Dati di anamnesi dialitica e nefrologica
        strSql = "SELECT * " & _
            " FROM (((((((ANAMNESI_DIALITICHE AN_D " & _
            "       LEFT OUTER JOIN ANAMNESI_NEFROLOGICHE AN_N ON AN_N.CODICE_PAZIENTE=AN_D.CODICE_PAZIENTE) " & _
            "       LEFT OUTER JOIN EDTA EDTA ON EDTA.KEY=AN_N.CODICE_EDTA) " & _
            "       LEFT OUTER JOIN PAZIENTI ON PAZIENTI.KEY=AN_D.CODICE_PAZIENTE) " & _
            "       LEFT OUTER JOIN FILTRI ON FILTRI.KEY=AN_D.TIPO_FILTRO) " & _
            "       LEFT OUTER JOIN ANTICOAGULANTI ON ANTICOAGULANTI.KEY=AN_D.ANTICOAGULANTE1) " & _
            "       LEFT OUTER JOIN ACCESSI_VASCOLARI ON ACCESSI_VASCOLARI.KEY=AN_D.ACCESSO_VASCOLARE) " & _
            "       LEFT OUTER JOIN TIPI_DIALISI ON TIPI_DIALISI.KEY=AN_D.TIPO_DIALISI) " & _
            " WHERE (STATO=0) " & _
            strSingoloPaziente & " " & _
            "ORDER BY   PAZIENTI.COGNOME, PAZIENTI.NOME"
        
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataset.EOF
            .AddNew
            .Fields("CODICE_PAZIENTE") = rsDataset("CODICE_ID")
            .Fields("COGNOME") = IIf(Len(rsDataset("COGNOME")) >= 3, Left(rsDataset("COGNOME"), 3), rsDataset("COGNOME"))
            .Fields("NOME") = IIf(Len(rsDataset("PAZIENTI.NOME")) >= 3, Left(rsDataset("PAZIENTI.NOME"), 3), rsDataset("PAZIENTI.NOME"))
            
            Dim somma As Integer
            If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
                somma = -1
            ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
                somma = -1
            Else
                somma = 0
            End If
            .Fields("ANNI") = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
            
            If rsDataset("G_SANGUIGNO") <> -1 Then
                If rsDataset("RH") <> -1 Then
                    .Fields("EMOGRUPPO") = Choose(rsDataset("G_SANGUIGNO") + 1, "A", "B", "AB", "0") & " " & Choose(rsDataset("RH") + 1, "POSITIVO", "NEGATIVO")
                Else
                    .Fields("EMOGRUPPO") = Choose(rsDataset("G_SANGUIGNO") + 1, "A", "B", "AB", "0")
                End If
            Else
                .Fields("EMOGRUPPO") = "- -"
            End If
            
            .Fields("ACCESSO_VASCOLARE") = rsDataset("ACCESSI_VASCOLARI.NOME")
            .Fields("EDTA") = rsDataset("CODICE")
            .Fields("DURATA_SEDUTA") = rsDataset("ORE") & " ore e " & rsDataset("MINUTI") & " minuti"
            .Fields("FILTRO") = rsDataset("FILTRI.NOME")
            .Fields("FILTRO_SEDUTA") = rsDataset("FILTRI.NOME") & " : LOTTO N°"
            .Fields("BAGNO_DIALISI") = "Na+ " & rsDataset("SODIO") & " / " & "K+ " & rsDataset("POTASSIO") & " / " & "HC03- " & rsDataset("BICARBONATO") & " / " & "Ca " & rsDataset("CALCIO") & " / " & "Gluc " & rsDataset("GLUCOSIO")
            .Fields("ANTICOAGULANTE") = rsDataset("ANTICOAGULANTI.NOME")
            .Fields("PESO_SECCO") = rsDataset("PESO_SECCO")
            
            
            strSql = "SELECT        TOP 1  PESO_FINE AS PESO, PA_MIN5, PA_MAX5, FC5   " & _
                    "FROM           PAZIENTI " & _
                    "               LEFT JOIN SCHEDE_DIALISI ON SCHEDE_DIALISI.CODICE_PAZIENTE=PAZIENTI.KEY  " & _
                    "Where          CODICE_PAZIENTE =" & rsDataset("Pazienti.KEY") & " " & _
                    "ORDER BY       DATA DESC"
            rsAppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rsAppo.RecordCount > 0 Then
                .Fields("PESO_DIALITICO_PRECEDENTE") = rsAppo("PESO")
                .Fields("PRESSIONE_PRECEDENTE") = rsAppo("PA_MAX5") & " / " & rsAppo("PA_MIN5")
                .Fields("FREQUENZA_PRECEDENTE") = rsAppo("FC5")
            Else
                .Fields("PESO_DIALITICO_PRECEDENTE") = 0
                .Fields("PRESSIONE_PRECEDENTE") = "--"
                .Fields("FREQUENZA_PRECEDENTE") = "--"
            End If
            rsAppo.Close
            
            Set rsEsami = New Recordset
            Set rsAppo2 = New Recordset
            rsEsami.Open "SELECT KEY, NOME, UNITA FROM VOCI_ESAMI WHERE ESAMI_DA_STAMPARE=TRUE ORDER BY NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            incrementa = 1
            Do While Not rsEsami.EOF
                strSql = "SELECT        TOP 1 VALORE, DATA " & _
                         "FROM          (ANAMNESI_ESAMI " & _
                         "              INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                         "WHERE         CODICE_PAZIENTE=" & rsDataset("PAZIENTI.KEY") & " AND " & _
                         "              CODICE_ESAME=" & rsEsami("KEY") & " " & _
                         "ORDER BY      DATA DESC"
            
                rsAppo2.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsAppo2.EOF And rsAppo2.BOF) Then
                    Select Case rsAppo2("VALORE")
                        Case -2
                            valore = "NEGATIVO"         ' controllo per i Marker Virali
                        Case -1
                            valore = "POSITIVO"
                        Case Else
                            valore = rsAppo2("VALORE")
                    End Select
                    .Fields("ESAME" & incrementa) = rsEsami("NOME") & " (" & rsEsami("UNITA") & ") " & valore & vbCrLf & "del " & rsAppo2("DATA")
                Else
                    .Fields("ESAME" & incrementa) = rsEsami("NOME") & " (" & rsEsami("UNITA") & ")"
                End If
                rsAppo2.Close
                incrementa = incrementa + 1
                rsEsami.MoveNext
            Loop
            rsEsami.Close
            
            

            ' Dati delle terapie dialitica e domiciliare
            
            Set rsDialitiche = New Recordset
            strSql = "  SELECT      * " & _
                    "   FROM        (TERAPIE_DIALITICHE " & _
                    "               INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
                    "   WHERE       TERAPIE_DIALITICHE.CODICE_PAZIENTE=" & rsDataset("PAZIENTI.KEY") & " AND SOSPESA=FALSE " & _
                    "   ORDER BY    TERAPIE_DIALITICHE.DATA DESC"
            rsDialitiche.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDialitiche.EOF And rsDialitiche.BOF) Then
                Do While Not rsDialitiche.EOF
                    strGiorni = ""
                If CBool(rsDialitiche("TUTTI_GIORNI")) Then
                    strGiorni = "Tutti i giorni"
                Else
                    For i = 1 To 7
                        If CBool(rsDialitiche("GIORNO" & i)) Then
                            strGiorni = strGiorni & UCase(Mid(WeekdayName(i), 1, 1)) & Mid(WeekdayName(i), 2, Len(WeekdayName(i))) & ", "
                        End If
                    Next
                    If strGiorni <> "" Then strGiorni = Mid(strGiorni, 1, Len(strGiorni) - 2)
                End If
                
                .Fields("FARMACO_DIALISI") = .Fields("FARMACO_DIALISI") & vbCrLf & rsDialitiche("NOME")
                .Fields("POSOLOGIA_DIALISI") = .Fields("POSOLOGIA_DIALISI") & vbCrLf & rsDialitiche("POSOLOGIA")
                .Fields("GIORNI_DIALISI") = .Fields("GIORNI_DIALISI") & vbCrLf & strGiorni
                .Fields("NOTE_DIALISI") = .Fields("NOTE_DIALISI") & vbCrLf & rsDialitiche("NOTE")
                rsDialitiche.MoveNext
                Loop
            Else
                .Fields("FARMACO_DIALISI") = "- -"
            End If
            rsDialitiche.Close
        
        
            Set rsDomiciliari = New Recordset
            strSql = "  SELECT      * " & _
                    "   FROM        (TERAPIE_DOMICILIARI " & _
                    "               INNER JOIN MEDICINALI ON TERAPIE_DOMICILIARI.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
                    "   WHERE       TERAPIE_DOMICILIARI.CODICE_PAZIENTE=" & rsDataset("PAZIENTI.KEY") & " AND SOSPESA=FALSE " & _
                    "   ORDER BY    TERAPIE_DOMICILIARI.DATA DESC"
            rsDomiciliari.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDomiciliari.EOF And rsDomiciliari.BOF) Then
                Do While Not rsDomiciliari.EOF
                    strGiorni = ""
                If CBool(rsDomiciliari("TUTTI_GIORNI")) Then
                    strGiorni = "Tutti i giorni"
                Else
                    For i = 1 To 7
                        If CBool(rsDomiciliari("GIORNO" & i)) Then
                            strGiorni = strGiorni & UCase(Mid(WeekdayName(i), 1, 1)) & Mid(WeekdayName(i), 2, Len(WeekdayName(i))) & ", "
                        End If
                    Next
                    If strGiorni <> "" Then strGiorni = Mid(strGiorni, 1, Len(strGiorni) - 2)
                End If
                
                .Fields("FARMACO_DOMICILIARE") = .Fields("FARMACO_DOMICILIARE") & vbCrLf & rsDomiciliari("NOME")
                .Fields("POSOLOGIA_DOMICILIARE") = .Fields("POSOLOGIA_DOMICILIARE") & vbCrLf & rsDomiciliari("POSOLOGIA")
                .Fields("GIORNI_DOMICILIARE") = .Fields("GIORNI_DOMICILIARE") & vbCrLf & strGiorni
                .Fields("NOTE_DOMICILIARE") = .Fields("NOTE_DOMICILIARE") & vbCrLf & rsDomiciliari("SOMMINISTRAZIONE")
                rsDomiciliari.MoveNext
                Loop
            Else
                .Fields("FARMACO_DOMICILIARE") = "- -"
            End If
            rsDomiciliari.Close
            
            If chkTutti.Value = Checked Then
                prgBarra.Value = prgBarra.Value + 1
                Else
            End If
            
            rsDataset.MoveNext
            .Update
        Loop
        rsDataset.Close
    End With
           
    prgBarra.Value = prgBarra.max
    Me.Height = 3345
            
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    Set rsAppo2 = Nothing
    Set rsEsami = Nothing
    Set rsDialitiche = Nothing
    Set rsDomiciliari = Nothing
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Il paziente non ha sedute dialitiche"
    Else
    
        Set rptModuloBartoli.DataSource = rsMain
        rptModuloBartoli.LeftMargin = 300
        rptModuloBartoli.TopMargin = 0
        rptModuloBartoli.BottomMargin = 0
        rptModuloBartoli.PrintReport True, rptRangeAllPages
    End If
End Sub

Private Function CalcolaTsat(c1 As Single, c2 As Single) As Double
    On Error GoTo gestione
    CalcolaTsat = Format(c1 / c2 * CSng("70,9"), "##.##")
    Exit Function
gestione:
    CalcolaTsat = 0
End Function

Private Function CalcolaKtv(c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single) As Double
    On Error GoTo gestione
    CalcolaKtv = Format((4 - 3.5 * c1 / c2) * (c4 / c5) - Log(c1 / c2 - 0.008 * c3), "##.##")
    Exit Function
gestione:
    CalcolaKtv = 0
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
    Select Case tStampa
        Case tpFOGLIOVIAGGIO
            Call StampaFogliViaggio
        Case tpIMPEGNATIVE
            Call StampaImpegnativeDialisi
        Case tpMODULOFIRMEPAZIENTE
            Call StampaModuloFirmePaziente
        Case tpSCHEDADIALITICASETTIMANALE
            Call StampaSchedaDialiticaSettimanale
        Case tpKTVANNUALE
            Call StampaKtvAnnuale
        Case tpTSATANNUALE
            Call StampaTsatAnnuale
    End Select
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



