VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
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
         ToolTipText     =   "Seleziona il paziente"
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
         FormatString    =   "| Paziente                                                       | N� Dialisi      | Cod. Prestaz.          "
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
            lblDatadiStampa(5).Enabled = False
            lblData.Enabled = False
            picData.Visible = False
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) - 1
            
        Case tpMODULOFIRMEPAZIENTE
            Me.Caption = "Modulo Firme Paziente"
            cboMese.ListIndex = Month(Now) - 1
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
                "       NEW adVarChar(110) AS GIORNI_DIALISI, " & _
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
            
            ' Gli passo da qui i valori del campo m� Fax
            ' per evitare che carichi il valore precedente dopo una modifica
            rptFogliViaggio.Sections("intestazione").Controls.Item("lblAltro").Caption = structIntestazione.sIndirizzo & " - " & structIntestazione.sCap & " " & structIntestazione.sCitta & " " & structIntestazione.sProv & vbCrLf & _
            "TEL: " & structIntestazione.sTelefono & "   FAX: " & structIntestazione.sFax & vbCrLf & _
            "P. Iva: " & structIntestazione.sIva & "   Cod. Fiscale: " & structIntestazione.sCodiceFiscale & vbCrLf & _
            "E-mail: " & structIntestazione.sMail & vbCrLf & _
            "Sito web: " & structIntestazione.sSitoWeb
            
            rptFogliViaggio.Sections("pie").Controls.Item("lblLuogo").Caption = structIntestazione.sCitta & " l�, " & GetUltimoGiorno(cboMese.ListIndex + 1, cboAnno.Text)
            
            If structIntestazione.sCodiceSTS = CODICESTS_S_LUCA Then
                rptFogliViaggio.Sections("corpo").Controls.Item("Label15").Caption = ""
                rptFogliViaggio.Sections("corpo").Controls.Item("Line2").Visible = False
            Else
                rptFogliViaggio.Sections("corpo").Controls.Item("Label15").Caption = "L'ASSISTITO A CONFERMA"
            End If

            rptFogliViaggio.PrintReport True, rptRangeAllPages
        Else
            MsgBox "Nessuna dialisi per il mese di " & cboMese, vbCritical, Me.Caption
        End If
    End If
End Sub

Private Sub StampaModuloFirmePaziente()
    If structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Or structIntestazione.sCodiceSTS = CODICESTS_SODAV Then
    Else
        MsgBox "MODULO DI STAMPA OPZIONALE ATTIVABILE A RICHIESTA", vbInformation, "INFORMAZIONE"
        Exit Sub
    End If

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
          If structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Then
            Set rptModuloFirmePaziente.DataSource = rsMain
            rptModuloFirmePaziente.LeftMargin = 1100
            rptModuloFirmePaziente.RightMargin = 0
            rptModuloFirmePaziente.TopMargin = 0
            rptModuloFirmePaziente.Sections("intestazione").Controls.Item("lblMese").Caption = cboMese & " " & cboAnno
            rptModuloFirmePaziente.Sections("pie").Controls.Item("lblStampato").Caption = structIntestazione.sCitta & " l�, " & lblData
            rptModuloFirmePaziente.PrintReport True, rptRangeAllPages
          ElseIf structIntestazione.sCodiceSTS = CODICESTS_SODAV Then
            Set rptModuloFirmeSodav.DataSource = rsMain
            rptModuloFirmeSodav.LeftMargin = 0 '1100
            rptModuloFirmeSodav.RightMargin = 0
            rptModuloFirmeSodav.TopMargin = 0
            rptModuloFirmeSodav.Sections("intestazione").Controls.Item("lblMese").Caption = cboMese & " " & cboAnno
            rptModuloFirmeSodav.Sections("pie").Controls.Item("lblStampato").Caption = structIntestazione.sCitta & " l�, " & lblData
            rptModuloFirmeSodav.PrintReport True, rptRangeAllPages
          End If
        End If
    End If
End Sub

Private Sub cmdTrova_Click()
    tTrova.Tipo = tpPAZIENTE
    tTrova.condStato = ""
    tTrova.condizione = ""
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
    End If
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    Select Case tStampa
        Case tpFOGLIOVIAGGIO
            Call StampaFogliViaggio
        Case tpMODULOFIRMEPAZIENTE
            Call StampaModuloFirmePaziente
    End Select
End Sub

Private Sub chkTutti_Click()
    If chkTutti.Value = Checked Then
        intPazientiKey = 0
        lblCognome = ""
        lblNome = ""
    End If
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

