VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStampaFiltri 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTempo 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6855
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
         ItemData        =   "frmStampaFiltri.frx":0000
         Left            =   1200
         List            =   "frmStampaFiltri.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2535
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
         ItemData        =   "frmStampaFiltri.frx":0096
         Left            =   5040
         List            =   "frmStampaFiltri.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   855
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
         TabIndex        =   14
         Top             =   250
         Width           =   585
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
         TabIndex        =   13
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame fraDati 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   6855
      Begin VB.OptionButton optTutti 
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
         TabIndex        =   24
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   240
         Picture         =   "frmStampaFiltri.frx":009A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   450
      End
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
         ItemData        =   "frmStampaFiltri.frx":04F3
         Left            =   3840
         List            =   "frmStampaFiltri.frx":04F5
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2775
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
         TabIndex        =   19
         Top             =   1200
         Width           =   735
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
         TabIndex        =   18
         Top             =   720
         Width           =   1005
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
         TabIndex        =   17
         Top             =   1200
         Width           =   4575
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
         TabIndex        =   16
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.Frame fraPeriodo 
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   6855
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Mattina"
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
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Pomeriggio"
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
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Sera"
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
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Mattina"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Pomeriggio"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Sera"
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
         Index           =   5
         Left            =   3960
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame panProgress 
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   6855
      Begin MSComctlLib.ProgressBar prgBarra 
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   6855
      Begin VB.CommandButton cmdAvanti 
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
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   1260
      End
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
         Left            =   5400
         TabIndex        =   11
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmStampaFiltri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intPazientiKey As Integer

Private Sub Form_Load()
    Dim i As Integer
    Dim label As String
        
    Select Case tStampa
        Case tpSCHEDADIALITICASETTIMANALE
            Me.Caption = "Scheda Dialitica Settimanale"
            cboMese.ListIndex = Month(Now) - 1
            cboAnno.AddItem Year(Now)
            cboAnno.AddItem Year(Now) + 1
            cboStato.Visible = False
            
        Case tpKTVANNUALE, tpTSATANNUALE, tpPTHAnnuale, tpCAPAnnuale
            Call RicaricaComboBox("TIPO_STATO", "NOME", cboStato)
            cboStato.AddItem "Tutti"
            cboStato.ItemData(cboStato.NewIndex) = 0
            cboStato.ListIndex = 0
            lblAnno.Left = lblMese.Left
            cboAnno.Left = cboMese.Left
            lblMese.Visible = False
            cboMese.Visible = False
            If tStampa = tpKTVANNUALE Then
                label = "Kt/V Annuale"
            ElseIf tStampa = tpTSATANNUALE Then
                label = "TSAT% Annuale"
            ElseIf tStampa = tpPTHAnnuale Then
                label = "PTH Annuale"
            Else: tStampa = tpCAPAnnuale
                label = " Ca/P Annuale"
            End If
            
            Me.Caption = Me.Caption & label
            For i = 0 To 5
                cboAnno.AddItem Year(Now) - i
            Next i
    End Select
    cboAnno.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Function Completo() As Boolean
    Completo = False
    If cboMese.ListIndex = -1 And tStampa = tpSCHEDADIALITICASETTIMANALE Then
        MsgBox "Selezionare il mese", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboAnno.ListIndex = -1 Then
        MsgBox "Selezionare l'anno", vbCritical, "Attenzione"
        Exit Function
    End If
    If optTutti.Value = False And intPazientiKey = 0 And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        MsgBox "Selezionare il paziente ", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

Private Sub StampaKtvAnnuale()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim cont As Integer
    Dim media_ann As Currency
    Dim i As Integer
    Dim condizione  As String
    
    If Not Completo Then Exit Sub

    If optTutti.Value = False And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
    
    condizione = GetCondizione
        
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
    rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME, STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " AND (STATODATA IS NULL OR YEAR(STATODATA)=" & cboAnno.Text & ") ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
       
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessun paziente presente nel turno selezionato", vbInformation, "Informazione"
        Exit Sub
    ElseIf rsMain.RecordCount > 0 Then rsMain.MoveFirst
     
        media_ann = 0
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
            media_ann = media_ann + rsMain("MEDIA")
        Else
            rsMain("MEDIA") = Null
        End If
        rsMain.MoveNext
    Loop
    
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatCapAnnuale.DataSource = rsMain
        rptKtvTsatCapAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "KT/V ANNO " & cboAnno.Text & " - PAZIENTI CON STATUS ''" & cboStato & "''"
        rptKtvTsatCapAnnuale.LeftMargin = 500
        rptKtvTsatCapAnnuale.RightMargin = 0
        rptKtvTsatCapAnnuale.TopMargin = 0
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblPazienti").Caption = rsMain.RecordCount
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblMedia_Ann").Caption = Int(media_ann / rsMain.RecordCount * 100) / 100
        rptKtvTsatCapAnnuale.PrintReport True, rptRangeAllPages
    End If
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
    Dim media_ann As Currency
    Dim i As Integer
    Dim condizione As String
    
    If Not Completo Then Exit Sub

    If optTutti.Value = False And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
    
    condizione = GetCondizione

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
    rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME, STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " AND (STATODATA IS NULL OR YEAR(STATODATA)=" & cboAnno.Text & ") ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
'    rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME,STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessun paziente presente nel turno selezionato", vbInformation, "Informazione"
        Exit Sub
    ElseIf rsMain.RecordCount > 0 Then rsMain.MoveFirst
        media_ann = 0
    
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
            media_ann = media_ann + rsMain("MEDIA")
        Else
            rsMain("MEDIA") = Null
        End If
        rsMain.MoveNext
    Loop
    
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatCapAnnuale.DataSource = rsMain
        rptKtvTsatCapAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "TSAT% ANNO " & cboAnno.Text & " - PAZIENTI CON STATUS ''" & cboStato & "''"
        rptKtvTsatCapAnnuale.LeftMargin = 500
        rptKtvTsatCapAnnuale.RightMargin = 0
        rptKtvTsatCapAnnuale.TopMargin = 0
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblPazienti").Caption = rsMain.RecordCount
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblMedia_Ann").Caption = Int(media_ann / rsMain.RecordCount * 100) / 100
        rptKtvTsatCapAnnuale.PrintReport True, rptRangeAllPages
    End If
    End If
End Sub

Private Sub StampaPthAnnuale()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim rsDataselect As Recordset
    Dim SQLString2 As String
        
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim cont As Integer
    Dim media_ann As Currency
    Dim i As Integer
    Dim condizione As String
    Dim keyEsame As Integer
    Dim keyGruppo As Integer
    Dim keyAnamnesi As Integer
    Dim keyRecord As Integer
    Dim mese As String
    Dim primogg As Date
    Dim ultimogg As Date
    Dim SommaMedia As Variant
    
    If Not Completo Then Exit Sub

    If optTutti.Value = False And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
    
    condizione = GetCondizione

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
    
 ' verifica se esiste l'esame pth
    rsDataset.Open "SELECT * FROM VOCI_ESAMI WHERE NOME like '%PTH%' ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
       keyEsame = rsDataset("KEY")
    Else
       MsgBox "La voce PTH non è presente nell'elenco degli esami di laboratorio", vbCritical, "Attenzione"
       rsDataset.Close
       Exit Sub
    End If
    rsDataset.Close
   
  ' verifica se esiste l'associazione con qualche gruppo esami lab
     rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_ESAME=" & keyEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
     If Not (rsDataset.EOF And rsDataset.BOF) Then
        keyGruppo = rsDataset("CODICE_GRUPPO")
     Else
        MsgBox "L'esame PTH è presente ma NON è associato ad alcun gruppo", vbCritical, "Attenzione"
        rsDataset.Close
        Exit Sub
     End If
     rsDataset.Close
     
    'If optTutti.Value = True Or optSessione(0).Value = True Or optSessione(1).Value = True Or optSessione(2).Value = True Or optSessione(3).Value = True Or optSessione(4).Value = True Or optSessione(5).Value = True Then
    '    prgBarra.Value = 0
    '    panProgress.Visible = True
    '    fraPulsanti.Top = 4080
    '    Me.Height = 5400
    'End If
    'DoEvents
    
   Set rsDataselect = New Recordset
   media_ann = 0
   cont = 0
   rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME, STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " AND (STATODATA IS NULL OR YEAR(STATODATA)=" & cboAnno.Text & ") ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
'   rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME,STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.RecordCount > 0 Then       ' se ci sono record attiva la barra di progresso
        prgBarra.Value = 0
        panProgress.Visible = True
        fraPulsanti.Top = 4080
        Me.Height = 5400
        prgBarra.max = rsDataset.RecordCount
    End If
    DoEvents
    Do While Not rsDataset.EOF
        With rsMain
            .AddNew
            .Fields("CODICE_PAZIENTE") = rsDataset("KEY")
            .Fields("COGNOME") = rsDataset("COGNOME")
            .Fields("NOME") = rsDataset("NOME")
            
           For i = 1 To 12
             Select Case i
                Case 1 To 9
                   mese = "0" & i
                Case Else
                   mese = i
             End Select
                               
             primogg = DateValue(mese & "/01/" & cboAnno.Text)
             ultimogg = DateValue(mese & "/" & Day(GetUltimoGiorno(val(mese), cboAnno.Text)) & "/" & cboAnno.Text)
            
            SQLString2 = "SELECT ANAMNESI_ESAMI.KEY,ESAMI_LAB.VALORE " & _
            "FROM (ANAMNESI_ESAMI " & _
            "INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY = ESAMI_LAB.CODICE_ANAMNESI_ESAMI ) " & _
            "WHERE ESAMI_LAB.CODICE_ESAME=" & keyEsame & " AND ANAMNESI_ESAMI.CODICE_PAZIENTE=" & rsDataset("KEY") & " AND ANAMNESI_ESAMI.DATA BETWEEN #" & primogg & "# AND #" & ultimogg & "#"

            rsDataselect.Open SQLString2, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
 
 '     Debug.Print SQLString2
 '     Debug.Print rsDataselect!key
 '     Debug.Print rsDataselect!valore
 
            If rsDataselect.RecordCount <> 0 Then
                .Fields("MESE" & i) = rsDataselect("VALORE")
                SommaMedia = SommaMedia + rsDataselect("VALORE")
                cont = cont + 1
            Else
                .Fields("MESE" & i) = Null
            End If
            rsDataselect.Close
          Next i

          If cont <> 0 Then
            .Fields("MEDIA") = Int(SommaMedia / cont * 100) / 100
            media_ann = media_ann + rsMain("MEDIA")
          End If

          .Update
        End With
        SommaMedia = 0
        cont = 0
            
        If optTutti.Value = True Or optSessione(0).Value = True Or optSessione(1).Value = True Or optSessione(2).Value = True Or optSessione(3).Value = True Or optSessione(4).Value = True Or optSessione(5).Value = True Then
            prgBarra.Value = prgBarra.Value + 1
            Else
        End If
        
        rsDataset.MoveNext
    Loop
    rsDataset.Close
   
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessun paziente presente nel turno selezionato", vbInformation, "Informazione"
        Exit Sub
    ElseIf rsMain.RecordCount > 0 Then rsMain.MoveFirst
      
    prgBarra.Value = prgBarra.max
    panProgress.Visible = False
    fraPulsanti.Top = 3360
    Me.Height = 4665
  
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatCapAnnuale.DataSource = rsMain
        rptKtvTsatCapAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "PTH ANNO " & cboAnno.Text & " - PAZIENTI CON STATUS ''" & cboStato & "''"
        rptKtvTsatCapAnnuale.LeftMargin = 500
        rptKtvTsatCapAnnuale.RightMargin = 0
        rptKtvTsatCapAnnuale.TopMargin = 0
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblPazienti").Caption = rsMain.RecordCount
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblMedia_Ann").Caption = Int(media_ann / rsMain.RecordCount * 100) / 100
        rptKtvTsatCapAnnuale.PrintReport True, rptRangeAllPages
    End If
    End If
End Sub

Private Sub StampaCAPAnnuale()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim strSingoloPaziente As String
    Dim strStato As String
    Dim cont As Integer
    Dim media_ann As Currency
    Dim i As Integer
    Dim condizione As String
    
    If Not Completo Then Exit Sub

    If optTutti.Value = False And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
    
    condizione = GetCondizione

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
    rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME, STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " AND (STATODATA IS NULL OR YEAR(STATODATA)=" & cboAnno.Text & ") ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
 '   rsDataset.Open "SELECT PAZIENTI.KEY,COGNOME,NOME,STATO FROM PAZIENTI left outer join turni on turni.codice_paziente=pazienti.key WHERE " & strStato & " " & strSingoloPaziente & condizione & " ORDER BY COGNOME, NOME, PAZIENTI.KEY", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    
    rsDataset.Open "SELECT * FROM PRODOTTO_CALCIO_FOSFORO WHERE ANNO=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        rsMain.Filter = "CODICE_PAZIENTE=" & rsDataset("CODICE_PAZIENTE")
        If rsMain.RecordCount <> 0 Then
            rsMain("MEDIA") = 0
            If IsNull(rsDataset("CALCEMIA")) Or IsNull(rsDataset("FOSFOREMIA")) Then
                rsMain.Fields("MESE" & rsDataset("MESE")) = Null
            Else
                 rsMain.Fields("MESE" & rsDataset("MESE")) = CalcolaCap(rsDataset("CALCEMIA"), rsDataset("FOSFOREMIA"))
            End If
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    rsMain.Filter = ""
    
    If rsMain.RecordCount = 0 Then
        MsgBox "Nessun paziente presente nel turno selezionato", vbInformation, "Informazione"
        Exit Sub
    ElseIf rsMain.RecordCount > 0 Then rsMain.MoveFirst
        
        media_ann = 0
    Do While Not rsMain.EOF
        cont = 0
        For i = 1 To 12
            If Not IsNull(rsMain("MESE" & i)) Then
                rsMain("MEDIA") = rsMain("MEDIA") + rsMain("MESE" & i)
                cont = cont + 1
            End If
        Next i
        If cont <> 0 Then
            rsMain("MEDIA") = Int(rsMain("MEDIA") / cont * 100) / 100
            media_ann = media_ann + rsMain("MEDIA")
        Else
            rsMain("MEDIA") = Null
        End If
        rsMain.MoveNext
    Loop
    
    Set rsDataset = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptKtvTsatCapAnnuale.DataSource = rsMain
        rptKtvTsatCapAnnuale.Sections("intestazione").Controls("lblTitolo").Caption = "Ca/P ANNO " & cboAnno.Text & " - PAZIENTI CON STATUS ''" & cboStato & "''"
        rptKtvTsatCapAnnuale.LeftMargin = 500
        rptKtvTsatCapAnnuale.RightMargin = 0
        rptKtvTsatCapAnnuale.TopMargin = 0
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblPazienti").Caption = rsMain.RecordCount
        rptKtvTsatCapAnnuale.Sections("Section5").Controls.Item("lblMedia_Ann").Caption = Int(media_ann / rsMain.RecordCount * 100) / 100
        rptKtvTsatCapAnnuale.PrintReport True, rptRangeAllPages
    End If
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
    Dim condizione As String

    If Not Completo Then Exit Sub

    If optTutti.Value = False And optSessione(0).Value = False And optSessione(1).Value = False And optSessione(2).Value = False And optSessione(3).Value = False And optSessione(4).Value = False And optSessione(5).Value = False Then
        strSingoloPaziente = " AND PAZIENTI.KEY=" & intPazientiKey
    End If
    
    If cboStato.ListIndex = cboStato.ListCount - 1 Then
        strStato = "TRUE"
    Else
        strStato = "STATO = " & cboStato.ListIndex
    End If
    
    condizione = GetCondizione
        
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
                "       NEW adVarChar (35) as ALLERGIA, " & _
                "       NEW adSingle as  FLUSSO_QB, " & _
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
            "       left outer join turni on turni.codice_paziente=pazienti.key " & _
            " WHERE (STATO=0) " & _
            strSingoloPaziente & " " & _
            condizione & " " & _
            "ORDER BY   PAZIENTI.COGNOME, PAZIENTI.NOME"
        
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If rsDataset.RecordCount = 0 Then
            MsgBox "Nessun paziente presente nel turno selezionato", vbInformation, "Informazione"
            Exit Sub
        End If
        
        If optTutti.Value = True Or optSessione(0).Value = True Or optSessione(1).Value = True Or optSessione(2).Value = True Or optSessione(3).Value = True Or optSessione(4).Value = True Or optSessione(5).Value = True Then
            prgBarra.Value = 0
            panProgress.Visible = True
            fraPulsanti.Top = 4080
            Me.Height = 5400
        End If
        DoEvents
        
        prgBarra.max = rsDataset.RecordCount
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
            .Fields("ALLERGIA") = rsDataset("ALLERGIA")
            .Fields("FLUSSO_QB") = rsDataset("FLUSSO_SANGUE")
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
            
            If optTutti.Value = True Or optSessione(0).Value = True Or optSessione(1).Value = True Or optSessione(2).Value = True Or optSessione(3).Value = True Or optSessione(4).Value = True Or optSessione(5).Value = True Then
                prgBarra.Value = prgBarra.Value + 1
                Else
            End If
            
            rsDataset.MoveNext
            .Update
        Loop
        rsDataset.Close
    End With
           
    prgBarra.Value = prgBarra.max
    panProgress.Visible = False
    fraPulsanti.Top = 3360
    Me.Height = 4665
            
    Set rsDataset = Nothing
    Set rsAppo = Nothing
    Set rsAppo2 = Nothing
    Set rsEsami = Nothing
    Set rsDialitiche = Nothing
    Set rsDomiciliari = Nothing
    
    Set rptModuloBartoli.DataSource = rsMain
    rptModuloBartoli.LeftMargin = 300
    rptModuloBartoli.TopMargin = 0
    rptModuloBartoli.BottomMargin = 0
    rptModuloBartoli.PrintReport True, rptRangeAllPages
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

Private Function CalcolaCap(c1 As Single, c2 As Single) As Double
    On Error GoTo gestione
    CalcolaCap = Format(c1 * c2, "##.##")
    Exit Function
gestione:
    CalcolaCap = 0
End Function

Private Function GetCondizione()
    Dim condizione As String
    Dim strPeriodo As String
    Dim giorni() As Variant
    Dim i As Integer
    
    If optSessione(0).Value Or optSessione(1).Value Or optSessione(2).Value Or optSessione(3).Value Or optSessione(4).Value Or optSessione(5).Value Then
        If optSessione(0).Value Or optSessione(3).Value Then strPeriodo = "AM_INIZIO"
        If optSessione(1).Value Or optSessione(4).Value Then strPeriodo = "PM_INIZIO"
        If optSessione(2).Value Or optSessione(5).Value Then strPeriodo = "SR_INIZIO"
        
        If optSessione(0).Value Or optSessione(1).Value Or optSessione(2).Value Then giorni = Array(2, 4, 6)
        If optSessione(3).Value Or optSessione(4).Value Or optSessione(5).Value Then giorni = Array(1, 3, 5)
        
        For i = 0 To UBound(giorni)
            condizione = condizione & strPeriodo & giorni(i) & "<>"""" OR "
        Next i
        condizione = " and (" & Left(condizione, Len(condizione) - 4) & ")"
    Else
        condizione = ""
    End If
    GetCondizione = condizione
End Function

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then Exit Sub
    ' carica i dati del paziente
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognome = rsDataset("COGNOME")
    lblNome = rsDataset("NOME")
    Set rsDataset = Nothing
    optTutti.Value = False
End Sub

Private Sub cmdAvanti_Click()
    Select Case tStampa
        Case tpSCHEDADIALITICASETTIMANALE
            Call StampaSchedaDialiticaSettimanale
        Case tpKTVANNUALE
            Call StampaKtvAnnuale
        Case tpTSATANNUALE
            Call StampaTsatAnnuale
        Case tpPTHAnnuale
            Call StampaPthAnnuale
        Case tpCAPAnnuale
            Call StampaCAPAnnuale
    End Select
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdTrova_Click()
Dim i As Integer

    For i = 0 To 5
        optSessione(i).Value = False
    Next i

    optTutti.Value = False
    
    tTrova.Tipo = tpPAZIENTE
    If tStampa = tpKTVANNUALE Or tStampa = tpTSATANNUALE Or tStampa = tpPTHAnnuale Or tStampa = tpCAPAnnuale Then
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
        If intPazientiKey <> tTrova.keyReturn Then
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub optTutti_Click()
Dim i As Integer

    For i = 0 To 5
        optSessione(i).Value = False
    Next i
    
    If optTutti.Value = True Then
        intPazientiKey = 0
        lblCognome = ""
        lblNome = ""
    End If
End Sub

Private Sub optSessione_Click(Index As Integer)
    optTutti.Value = False
    If optSessione(Index).Value = True Then
        intPazientiKey = 0
        lblCognome = ""
        lblNome = ""
    End If
End Sub
