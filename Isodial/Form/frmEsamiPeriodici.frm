VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEsamiPeriodici 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esami periodici in ED"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmEsamiPeriodici.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   2280
         TabIndex        =   21
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
         Left            =   6840
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Left            =   10440
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame fraListaMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   12015
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Mensili                                                           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":0459
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   1
         Left            =   8040
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Trimestrali                                                       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":05B3
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Semestrali                                                      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":070D
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   3
         Left            =   4080
         TabIndex        =   3
         Top             =   2400
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Annuali                                                          "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":0867
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   4
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "|| Bimestrale                                                           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":09C1
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1980
         Index           =   5
         Left            =   8040
         TabIndex        =   5
         Top             =   2400
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3493
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Se problemi clinici                                       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsamiPeriodici.frx":0B1B
      End
   End
   Begin VB.Frame fraAzioni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   12015
      Begin VB.CommandButton cmdCopiaPerTuttiPazienti 
         Caption         =   "Duplica per &tutti i pazienti"
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
         Left            =   4920
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopiaPerPazienteSingolo 
         Caption         =   "Duplica per p&aziente"
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
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   1695
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
         Height          =   600
         Left            =   6720
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
         Left            =   9360
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Inserisci"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8040
         TabIndex        =   7
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
         Height          =   600
         Left            =   10680
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1720
         TabIndex        =   13
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aggiornato al:"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1470
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEsamiPeriodici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmEsamiPeriodici.frm
'
' <b>Descrizione</b>: Scheda Esami Periodici associata alla tab ESAMI_PERIODICI
'
' @remarks
'
' @author
'
' @date 05/02/2011 18.19
Option Explicit

'' rs della scheda
Dim rsEsami As Recordset
Dim vRow As Integer
'' Tiene traccia della flx selezionata
Dim flxFocus As Integer
'' data di nascita del paziente
Dim dataNascita As Date
Dim intPeriodo As tipoPeriodo
Dim intPazientiKey As Integer

Private Type vettoreRiga
    riga(1 To 6) As String
End Type

'' Apre frmTrova se non c'è nessun paziente gia caricato
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
    intPeriodo = tipoPeriodo.tpMENSILE
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    For i = 0 To 5
        With flxGriglia(i)
            .ColWidth(0) = 0    ' key del record
            .ColWidth(1) = 0    ' codice associazione (valori positivi esami stumentali, valori negativi esami di lab)
            .Row = 0
            .Col = 2
            .CellFontBold = True
            .ColAlignment(2) = vbLeftJustify
            .MousePointer = flexDefault
        End With
    Next i
    flxFocus = -1
    vRow = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    Dim i As Integer
'    For i = 0 To 5
'        If ControlHWnd = flxGriglia(i).hWnd Then
'            If flxGriglia(i).TopRow - Rotation > 0 Then
'                If flxGriglia(i).TopRow - Rotation < flxGriglia(i).Rows Then
'                    flxGriglia(i).TopRow = flxGriglia(i).TopRow - Rotation
'                    Exit For
'                End If
'            End If
'        End If
'    Next i
'End Sub
'-----------------------------------

Public Sub InserisciEsame_Click()
    Call cmdInserisci_Click
End Sub

Public Sub StampaPrescrizione_Click()
    Call cmdStampa_Click
End Sub

Public Sub StampaStandard_Click()
    Call StampaStandard
End Sub

Public Sub StampaStandard()
    Dim strSql As String
    Dim strSqlShape As String
    Dim strPazienti As String
    Dim strNomeEsame As String
    Dim dataNascita As Date
    Dim i As Integer
    Dim k As Integer
    Dim intRecordCorrente As Integer
    Dim matrice() As vettoreRiga
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    strSqlShape = "SHAPE APPEND " & _
                "   NEW adLongVarChar AS ESAME1, " & _
                "   NEW adLongVarChar AS ESAME2, " & _
                "   NEW adLongVarChar AS ESAME3, " & _
                "   NEW adLongVarChar AS ESAME4, " & _
                "   NEW adLongVarChar AS ESAME5, " & _
                "   NEW adLongVarChar AS ESAME6 "
      

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlShape, cnConn, adOpenStatic, adLockOptimistic

    ReDim matrice(0)
    Set rsDataset = New Recordset
    For i = 1 To 6
        strSql = " Select       ESAMI_PERIODICI.PERIODO, ESAMI_PERIODICI.CODICE_ESAME, VOCI_ESAMI.NOME AS VOCI_ESAMINOME, ESAMI.NOME AS ESAMINOME, PAZIENTI.COGNOME, PAZIENTI.NOME AS PAZIENTINOME, PAZIENTI.DATA_NASCITA " & _
                " From         (((ESAMI_PERIODICI " & _
                "              Left join VOCI_ESAMI on VOCI_ESAMI.KEY=-ESAMI_PERIODICI.CODICE_ESAME) " & _
                "              Left Join ESAMI ON ESAMI.KEY=ESAMI_PERIODICI.CODICE_ESAME) " & _
                "              Inner Join PAZIENTI ON PAZIENTI.KEY=ESAMI_PERIODICI.CODICE_PAZIENTE) " & _
                " Where        CODICE_PAZIENTE=" & intPazientiKey & " AND PERIODO=" & i & _
                " Order by     PERIODO, VOCI_ESAMI.NOME, ESAMI.NOME"
             
    
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            strPazienti = rsDataset("COGNOME") & " " & rsDataset("PAZIENTINOME")
            dataNascita = rsDataset("DATA_NASCITA")
            intRecordCorrente = 0
            Do While Not rsDataset.EOF
                intRecordCorrente = intRecordCorrente + 1
                If rsDataset("CODICE_ESAME") < 0 Then
                    strNomeEsame = rsDataset("VOCI_ESAMINOME")
                Else
                    strNomeEsame = rsDataset("ESAMINOME")
                End If
                
                If Not intRecordCorrente > UBound(matrice) Then
                    matrice(intRecordCorrente).riga(i) = strNomeEsame
                Else
                    ReDim Preserve matrice(UBound(matrice) + 1)
                    matrice(UBound(matrice)).riga(i) = strNomeEsame
                End If
                rsDataset.MoveNext
            Loop
        End If
        rsDataset.Close
    Next i
    Set rsDataset = Nothing

    If UBound(matrice) <> 0 Then
        For i = 1 To UBound(matrice)
            rsMain.AddNew
            For k = 1 To 6
                rsMain.Fields("ESAME" & k) = matrice(i).riga(k)
            Next k
            rsMain.Update
        Next i
        
        structIntestazione.sPaziente = strPazienti
        structIntestazione.sDataPaziente = dataNascita
        laData = lblData
    
        Set rptEsamiPeriodici.DataSource = rsMain
        rptEsamiPeriodici.Orientation = rptOrientLandscape
        rptEsamiPeriodici.PrintReport False, rptRangeAllPages
        
    Else
        MsgBox "Nessun esame periodico assegnato al paziente", vbInformation, "Stampa Esami Periodici"
    End If

End Sub

Public Sub StampaPrescrizioni(intCodicePaziente As Integer, strPeriodo As String, blnStampaDicituraImpostata As Boolean, MeseRichiestaStampa As String, AnnoRichiestaStampa As String, DataRichiestaStampa As String)
    Const numMaxEsami As Integer = 19
    Dim strSql As String
    Dim strSqlShape As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset

    Dim intNumeroColonne As Integer
    Dim intNumeroPagina As Integer
    Dim intNumeroEsame As Integer
    Dim strPaziente As String
    Dim strNomeEsame As String
    Dim dicitura As String
        
    strSqlShape = "SHAPE APPEND " & _
                "       NEW adInteger AS INDEX, " & _
                "       NEW adVarChar(50) AS ESAME1, " & _
                "       NEW adVarChar(50) AS ESAME2 "

           
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlShape, cnConn, adOpenStatic, adLockOptimistic
       
    Set rsDataset = New Recordset
    strSql = " Select       ESAMI_PERIODICI.PERIODO, ESAMI_PERIODICI.CODICE_ESAME, VOCI_ESAMI.NOME AS VOCI_ESAMINOME, ESAMI.NOME AS ESAMINOME, PAZIENTI.COGNOME, PAZIENTI.NOME AS PAZIENTINOME, PAZIENTI.DATA_NASCITA " & _
             " From         (((ESAMI_PERIODICI " & _
             "              Left join VOCI_ESAMI on VOCI_ESAMI.KEY=-ESAMI_PERIODICI.CODICE_ESAME) " & _
             "              Left Join ESAMI ON ESAMI.KEY=ESAMI_PERIODICI.CODICE_ESAME) " & _
             "              Inner Join PAZIENTI ON PAZIENTI.KEY=ESAMI_PERIODICI.CODICE_PAZIENTE) " & _
             " Where        CODICE_PAZIENTE=" & intCodicePaziente & strPeriodo & _
             " Order by     VOCI_ESAMI.NOME, ESAMI.NOME, PERIODO"
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        intNumeroColonne = 1
        intNumeroPagina = 1
        intNumeroEsame = 0
        strPaziente = rsDataset("COGNOME") & " " & rsDataset("PAZIENTINOME")
        Do While Not rsDataset.EOF
            With rsMain
                If intNumeroColonne = 1 Then
                    .AddNew
                    .Fields("INDEX") = intNumeroPagina
                Else
                    If intNumeroEsame <> 0 Then .MoveNext
                End If
                
                If rsDataset("CODICE_ESAME") < 0 Then
                    strNomeEsame = rsDataset("VOCI_ESAMINOME")
                Else
                    strNomeEsame = rsDataset("ESAMINOME")
                End If
                .Fields("ESAME" & intNumeroColonne) = "- " & strNomeEsame
                
                intNumeroEsame = intNumeroEsame + 1
                If intNumeroEsame = numMaxEsami Then
                    intNumeroEsame = 0
                    If intNumeroColonne = 1 Then
                        .Filter = ("INDEX=" & intNumeroPagina)
                        intNumeroColonne = 2
                    Else
                        .Filter = ""
                        intNumeroPagina = intNumeroPagina + 1
                        intNumeroColonne = 1
                    End If
                End If
        
            End With
            rsDataset.MoveNext
        Loop
        rsMain.UpdateBatch
        
        ' controllo per stampare la dicitura impostata
        
        If blnStampaDicituraImpostata = True Then
            Set rsDataset = New Recordset
                rsDataset.Open "INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                dicitura = rsDataset("DICITURA_ESAMI_PERIODICI") & ""
                End If
        End If
                
        Set rptRichiestaEsamiLaboratorio.DataSource = rsMain
        rptRichiestaEsamiLaboratorio.Sections("intestazione").Controls.Item("lblTitolo").Caption = "Si richiedono i seguenti esami relativi al mese di: " & UCase(MeseRichiestaStampa) & " " & AnnoRichiestaStampa
        rptRichiestaEsamiLaboratorio.Sections("intestazione").Controls.Item("lblPaziente").Caption = strPaziente
        rptRichiestaEsamiLaboratorio.Sections("pie").Controls.Item("lblDicitura").Caption = dicitura
        rptRichiestaEsamiLaboratorio.Sections("pie").Controls.Item("lblData").Caption = "Data, " & UCase(DataRichiestaStampa)
        rptRichiestaEsamiLaboratorio.PrintReport False, rptRangeAllPages
        
    Else
        MsgBox "Nessun esame periodico assegnato al paziente nei periodi selezionati", vbInformation, "Stampa Esami Periodici"
    End If
    
    Set rsDataset = Nothing
    
End Sub

'' Pulisce l'intera scheda
Private Function PulisciTutto()
    Dim i As Integer
    intPazientiKey = 0
    lblData = ""
    flxFocus = -1
    For i = 0 To 5
        flxGriglia(i).Rows = 1
    Next i
    lblCognome = ""
    lblNome = ""
    lblEta = ""
    intPeriodo = tipoPeriodo.tpMENSILE
    cmdTrova.SetFocus
End Function

'' Elimina gli esami dei pazienti prima di effettuare i vari duplica esami...
Private Sub EliminaEsamiPeriodici(strCondizione As String)
    Dim cmCommand As New Command
    
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    cmCommand.CommandText = "DELETE * FROM ESAMI_PERIODICI WHERE (NOT CODICE_PAZIENTE=" & intPazientiKey & ") AND CODICE_PAZIENTE IN (" & strCondizione & ")"
    cmCommand.Execute
    
    Set cmCommand = Nothing
End Sub

'' Carica la flx
'
' @param index indice della flx da caricare
Private Sub CaricaFlx(Index As Integer)
    Dim rsLab As New Recordset
    Dim rsStrumentali As New Recordset
    Set rsEsami = New Recordset
    
    rsLab.Open "VOCI_ESAMI ", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
    rsStrumentali.Open "ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsEsami.Open "SELECT * FROM ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND PERIODO=" & Index + 1, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    flxGriglia(Index).Rows = 1
    Do While Not rsEsami.EOF
        With flxGriglia(Index)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsEsami("KEY")
            .TextMatrix(.Rows - 1, 1) = rsEsami("CODICE_ESAME")
            If .TextMatrix(.Rows - 1, 1) < 0 Then
                ' esami lab
                rsLab.Filter = "KEY=" & Abs(rsEsami("CODICE_ESAME"))
                .TextMatrix(.Rows - 1, 2) = rsLab("NOME")
            Else
                ' esami strumentali
                rsStrumentali.Filter = "KEY=" & rsEsami("CODICE_ESAME")
                .TextMatrix(.Rows - 1, 2) = rsStrumentali("NOME")
            End If
            rsEsami.MoveNext
        End With
    Loop
    Set rsLab = Nothing
    Set rsStrumentali = Nothing
    Set rsEsami = Nothing
End Sub

'' Carica la scheda
Private Sub CaricaScheda()
    Dim i As Integer
    Set rsEsami = New Recordset
    Dim rsLab As New Recordset
    Dim rsStrumentali As New Recordset
    
    For i = 0 To 5
        flxGriglia(i).Rows = 1
    Next i
    
    rsLab.Open "VOCI_ESAMI", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
    rsStrumentali.Open "ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsEsami.Open "SELECT * FROM ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsEsami.EOF And rsEsami.BOF) Then
        Do While Not rsEsami.EOF
            With flxGriglia(rsEsami("PERIODO") - 1)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsEsami("KEY")
                .TextMatrix(.Rows - 1, 1) = rsEsami("CODICE_ESAME")
                If .TextMatrix(.Rows - 1, 1) < 0 Then
                    ' esami lab
                    rsLab.Filter = "KEY=" & Abs(rsEsami("CODICE_ESAME"))
                    .TextMatrix(.Rows - 1, 2) = rsLab("NOME")
                Else
                    ' esami strumentali
                    rsStrumentali.Filter = "KEY=" & rsEsami("CODICE_ESAME")
                    .TextMatrix(.Rows - 1, 2) = rsStrumentali("NOME")
                End If
                rsEsami.MoveNext
            End With
        Loop
        rsEsami.Close
        
        rsEsami.Open "SELECT * FROM AGG_ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsEsami.EOF And rsEsami.BOF) Then
            lblData = rsEsami("DATA")
        End If
        rsEsami.Close
    End If
    
    Set rsLab = Nothing
    Set rsStrumentali = Nothing
    Set rsEsami = Nothing
End Sub

'' Verifica se l'esame è presente nella flx prima di inserirlo
Private Function IsPresente(codice As Integer) As Boolean
    Dim i As Integer
    
    IsPresente = False
    For i = 1 To flxGriglia(tEsamiPeriodici.periodo - 1).Rows - 1
        With flxGriglia(tEsamiPeriodici.periodo - 1)
            If .TextMatrix(i, 1) = codice Then
                IsPresente = True
                Exit For
            End If
        End With
    Next i
End Function

'' Inserisce un esame singolo nel db
'
' @return key dell'esame inserito nel db
Private Function InserisciSingolo(codice As Integer) As Integer
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim num As Integer
    Dim rsDataset As Recordset
    
    num = GetNumero("ESAMI_PERIODICI")
    v_Nomi = Array("KEY", "CODICE_PAZIENTE", "CODICE_ESAME", "PERIODO")
    v_Val = Array(num, intPazientiKey, codice, tEsamiPeriodici.periodo)
    
    Set rsDataset = New Recordset
    rsDataset.Open "ESAMI_PERIODICI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_Nomi, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
    
    InserisciSingolo = num
End Function

'' Aggiorna la flx e la data dell'ultima modifica sugli esami periodici
'
' @param num key dell'ultimo record inserito nel db
Private Sub AggiornaFlx(num As Integer)
    Dim rsDataset As Recordset
    Dim i As Integer
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM AGG_ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        rsDataset.AddNew Array("KEY", "CODICE_PAZIENTE", "DATA"), Array(GetNumero("AGG_ESAMI_PERIODICI"), intPazientiKey, date)
        rsDataset.Update
    Else
        rsDataset.Update "DATA", date
    End If
    Set rsDataset = Nothing
    lblData = date
    
    ' aggiorna solo quella flx
    Call CaricaFlx(tEsamiPeriodici.periodo - 1)
    
    ' si posiziona sul record e lo seleziona
    vRow = Esiste(flxGriglia(tEsamiPeriodici.periodo - 1), 0, vRow, num)
    flxGriglia(tEsamiPeriodici.periodo - 1).Row = vRow
    If vRow > 7 Then
        flxGriglia(tEsamiPeriodici.periodo - 1).TopRow = vRow
    End If
    
    ' discolora le altre flx
    For i = 0 To 5
        If i <> tEsamiPeriodici.periodo - 1 Then
            Call ColoraFlx(flxGriglia(i), flxGriglia(i).Cols - 1, True)
        End If
    Next i
    Call ColoraFlx(flxGriglia(tEsamiPeriodici.periodo - 1), flxGriglia(tEsamiPeriodici.periodo - 1).Cols - 1)
    
'    MsgBox "Inserimento valore effettuato", vbInformation, "Inserimento"
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
    Else
        Dim lfrmEsamiPeriodiciStampa As New frmEsamiPeriodiciStampa
        lfrmEsamiPeriodiciStampa.Show 1

        If Not lfrmEsamiPeriodiciStampa.blnStampa Then Exit Sub

        Dim strPeriodo As String
    
        On Error GoTo gestione
        cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
        cdlStampa.CancelError = True
        cdlStampa.ShowPrinter
        Me.Refresh

          If Not lfrmEsamiPeriodiciStampa.intTipoStampa = 1 Then
            strPeriodo = ""
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpMENSILE Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpMENSILE & " OR "
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpBIMESTRALE Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpBIMESTRALE & " OR "
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpTRIMESTRALE Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpTRIMESTRALE & " OR "
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpSEMESTRALE Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpSEMESTRALE & " OR "
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpANNUALE Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpANNUALE & " OR "
            If lfrmEsamiPeriodiciStampa.intPeriodo = tpPROBLEMI Then strPeriodo = strPeriodo & " PERIODO=" & tipoPeriodo.tpPROBLEMI & " OR "
            strPeriodo = "AND (" & Mid(strPeriodo, 1, Len(strPeriodo) - 3) & ")"
            
            If lfrmEsamiPeriodiciStampa.intTipoStampa = 3 Then
                Dim rsPazienti As New Recordset
                rsPazienti.Open "SELECT KEY FROM PAZIENTI WHERE (STATO=0) ORDER BY COGNOME, NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                Call StartProgressBar(rsPazienti.RecordCount, 0, Me)
                Do While Not rsPazienti.EOF
                    frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
                    Call StampaPrescrizioni(rsPazienti!key, strPeriodo, lfrmEsamiPeriodiciStampa.blnStampaDicituraImpostata, lfrmEsamiPeriodiciStampa.MeseRichiestaStampa, lfrmEsamiPeriodiciStampa.AnnoRichiestaStampa, lfrmEsamiPeriodiciStampa.DataRichiestaStampa)
                    rsPazienti.MoveNext
                Loop
                rsPazienti.Close
                Set rsPazienti = Nothing
                Call StopProgressBar(Me)
            Else
                Call StampaPrescrizioni(intPazientiKey, strPeriodo, lfrmEsamiPeriodiciStampa.blnStampaDicituraImpostata, lfrmEsamiPeriodiciStampa.MeseRichiestaStampa, lfrmEsamiPeriodiciStampa.AnnoRichiestaStampa, lfrmEsamiPeriodiciStampa.DataRichiestaStampa)
            End If
        Else
            Call StampaStandard
        End If
        
        Set lfrmEsamiPeriodiciStampa = Nothing
    End If
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdInserisci_Click()
    Dim Ok As Boolean
    Dim num As Integer
    Dim rsDataset As Recordset
        
    Unload frmEsamiPeriodiciInput
    tEsamiPeriodici.interoGruppo = -1
    Ok = True
    Do
       
        frmEsamiPeriodiciInput.LetPeriodo (intPeriodo)
        frmEsamiPeriodiciInput.Show 1
        If tEsamiPeriodici.interoGruppo > 0 Then
            ' inserimento di un gruppo di esami
            If tEsamiPeriodici.codiceAssociazione > 0 Then
                Set rsDataset = New Recordset
                rsDataset.Open "SELECT * FROM ESAMI WHERE CODICE_ORGANO=" & tEsamiPeriodici.codiceAssociazione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                ' inserisce l'intero gruppo esame per volta se non è gia presente
                Do While Not rsDataset.EOF
                    If Not IsPresente(rsDataset("KEY")) Then
                        num = InserisciSingolo(rsDataset("KEY"))
                    End If
                    rsDataset.MoveNext
                Loop
                Set rsDataset = Nothing
                Call AggiornaFlx(num)
                flxFocus = tEsamiPeriodici.periodo - 1
            Else
                Set rsDataset = New Recordset
                rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & -tEsamiPeriodici.codiceAssociazione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                ' inserisce l'intero gruppo esame per volta se non è gia presente
                Do While Not rsDataset.EOF
                    If Not IsPresente(-rsDataset("CODICE_ESAME")) Then
                        num = InserisciSingolo(-rsDataset("CODICE_ESAME"))
                    End If
                    rsDataset.MoveNext
                Loop
                Set rsDataset = Nothing
                Call AggiornaFlx(num)
                flxFocus = tEsamiPeriodici.periodo - 1
            End If
        ElseIf tEsamiPeriodici.interoGruppo = 0 Then
            ' inserimento di un singolo esame
            If IsPresente(tEsamiPeriodici.codiceAssociazione) Then
                Ok = False
                MsgBox "L'esame selezionato è già presente", vbCritical, "Attenzione"
            Else
                num = InserisciSingolo(tEsamiPeriodici.codiceAssociazione)
                Call AggiornaFlx(num)
                flxFocus = tEsamiPeriodici.periodo - 1
            End If
        Else
            ' premuto annulla
            Exit Sub
        End If
    Loop Until Ok
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    ' disattivo il form per evitare l'errore del click sulla flx ancora non caricata
    Me.Enabled = False
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    Call CaricaPaziente
    Me.Enabled = True
End Sub

'' Elimina solo un esame
Private Sub cmdElimina_Click()
    Dim eliminato As Boolean
    eliminato = False
    
    If intPazientiKey = 0 Then Exit Sub
    If vRow = 0 Then
        MsgBox "Selezionare l'esame da eliminare", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If MsgBox("Sei sicuro di voler eliminare l'esame " & UCase(flxGriglia(flxFocus).TextMatrix(vRow, 2)) & "?", vbQuestion & vbYesNo, "Eliminazione") = vbYes Then
        Set rsEsami = New Recordset
        rsEsami.Open "SELECT * FROM ESAMI_PERIODICI WHERE KEY=" & flxGriglia(flxFocus).TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        If rsEsami.BOF And rsEsami.EOF Then
            MsgBox "Impossibile eliminare", vbCritical, "Errore"
        Else
            rsEsami.Delete
            ' elimina dalla flx
            If flxGriglia(flxFocus).Rows = 2 Then
                flxGriglia(flxFocus).Rows = 1
            Else
                flxGriglia(flxFocus).RemoveItem (vRow)
            End If
            vRow = 0
            flxGriglia(flxFocus).Row = 0
            eliminato = True
        End If
        Set rsEsami = Nothing
    End If
    
    If eliminato Then
        Set rsEsami = New Recordset
        rsEsami.Open "SELECT COUNT(*) AS CONTEGGIO FROM ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsEsami("CONTEGGIO") = 0 Then
            rsEsami.Close
            rsEsami.Open "SELECT * FROM AGG_ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            If Not (rsEsami.EOF And rsEsami.BOF) Then
                rsEsami.Delete
                lblData = ""
            End If
        End If
        Set rsEsami = Nothing
    End If
End Sub

'' Duplica gli esami del pazienti a tutti i pazienti in dialisi
Private Sub cmdCopiaPerTuttiPazienti_Click()
    On Error GoTo gestione
    Dim rsAppo As New Recordset
    Dim rsPazienti As New Recordset
    
    If MsgBox("Sicuro di voler duplicare gli esami a tutti i pazienti in dialisi?", vbQuestion + vbYesNo, "Duplica per pazienti in dialisi") = vbYes Then
        Set rsEsami = New Recordset
        If MsgBox("ATTENZIONE!!! LA DUPLICAZIONE SOSTITUIRA' TUTTI GLI ESAMI PRESCRITTI" & vbCrLf & "Sicuro di volerli duplicare per tutti i pazienti in dialisi?", vbQuestion + vbYesNo, "Duplica per pazienti in dialisi") = vbYes Then
            Call EliminaEsamiPeriodici("SELECT KEY FROM PAZIENTI WHERE (STATO=0)")
            
            rsEsami.Open "ESAMI_PERIODICI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAppo.Open "SELECT * FROM ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                rsPazienti.Open "SELECT KEY FROM PAZIENTI WHERE (STATO=0) AND KEY<>" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                Call StartProgressBar(rsPazienti.RecordCount, 0, Me)
                
                Do While Not rsPazienti.EOF
                    frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
                    rsAppo.MoveFirst
                    Do While Not rsAppo.EOF
                        rsEsami.AddNew
                        rsEsami("KEY") = GetNumero("ESAMI_PERIODICI")
                        rsEsami("CODICE_PAZIENTE") = rsPazienti("KEY")
                        rsEsami("PERIODO") = rsAppo("PERIODO")
                        rsEsami("CODICE_ESAME") = rsAppo("CODICE_ESAME")
                        rsEsami.Update
                        rsAppo.MoveNext
                    Loop
                    rsPazienti.MoveNext
                Loop
                rsPazienti.Close
                
                Call StopProgressBar(Me)
                
                MsgBox "Duplicazione avvenuta con successo", vbInformation, "Duplica esami"
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
            End If
            rsAppo.Close
            rsEsami.Close
        End If
        Set rsEsami = Nothing
        
    End If
    
    Exit Sub
gestione:
    Call StopProgressBar(Me)
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

'' Duplica gli esami del paziente al paziente selezionato
Private Sub cmdCopiaPerPazienteSingolo_Click()
    On Error GoTo gestione
    Dim rsAppo As New Recordset
    Dim key As Integer
    
    If MsgBox("Sicuro di voler duplicare gli esami ad un paziente in dialisi?", vbQuestion + vbYesNo, "Duplica per paziente in dialisi") = vbYes Then
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = "STATO=0 AND NOT KEY=" & intPazientiKey
        tTrova.condStato = "(-1)"
        frmTrova.Show 1
        key = tTrova.keyReturn
        If key = 0 Then Exit Sub
        
        Set rsEsami = New Recordset
        If MsgBox("ATTENZIONE!!! LA DUPLICAZIONE SOSTITUIRA' TUTTI GLI ESAMI PRECEDENTI" & vbCrLf & "Sicuro di volerli duplicare per un altro paziente in dialisi?", vbQuestion + vbYesNo, "Duplica per paziente in dialisi") = vbYes Then
            Call EliminaEsamiPeriodici(Str(key))
            
            rsAppo.Open "SELECT * FROM ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                rsEsami.Open "ESAMI_PERIODICI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                Do While Not rsAppo.EOF
                    rsEsami.AddNew
                    rsEsami("KEY") = GetNumero("ESAMI_PERIODICI")
                    rsEsami("CODICE_PAZIENTE") = key
                    rsEsami("PERIODO") = rsAppo("PERIODO")
                    rsEsami("CODICE_ESAME") = rsAppo("CODICE_ESAME")
                    rsEsami.Update
                    rsAppo.MoveNext
                Loop
                rsEsami.Close
                
                rsEsami.Open "SELECT * FROM AGG_ESAMI_PERIODICI WHERE CODICE_PAZIENTE=" & key, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                If rsEsami.EOF And rsEsami.BOF Then
                    rsEsami.AddNew Array("KEY", "CODICE_PAZIENTE", "DATA"), Array(GetNumero("AGG_ESAMI_PERIODICI"), key, date)
                    rsEsami.Update
                Else
                    rsEsami.Update "DATA", date
                End If
                rsEsami.Close

                MsgBox "Duplicazione avvenuta con successo", vbInformation, "Duplica esami"
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
            End If
            rsAppo.Close
        End If
        Set rsEsami = Nothing
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub flxGriglia_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i As Integer
    
    If flxGriglia(Index).Rows = 1 Then Exit Sub
    If flxGriglia(Index).Row = flxGriglia(Index).Rows - 1 Then
        i = 1
    Else
        i = flxGriglia(Index).Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia(Index).TextMatrix(i, 2), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia(Index).Row = i
            If i >= 7 Or flxGriglia(Index).TopRow > 7 Then
                flxGriglia(Index).TopRow = i
            End If
            Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
            Exit Do
        End If
        If i = flxGriglia(Index).Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxGriglia(Index).Row
End Sub

Private Sub flxGriglia_Click(Index As Integer)
    Dim i As Integer
    flxGriglia(Index).SetFocus
    If VerificaClickFlx(flxGriglia(Index)) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1, True)
        flxFocus = -1
        ' annulla le row e col
        vRow = 0
        flxGriglia(Index).Row = 0
        flxGriglia(Index).Col = 0
        intPeriodo = 1
    Else
        ' discolora le altre flx
        For i = 0 To 5
            If i <> Index Then
                Call ColoraFlx(flxGriglia(i), flxGriglia(i).Cols - 1, True)
            End If
        Next i
        ' colora la selezionata
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
        flxFocus = Index
        vRow = flxGriglia(Index).Row
        intPeriodo = Index + 1
    End If
End Sub

Private Sub flxGriglia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        intPeriodo = Index + 1
        frmMain.PopupMenu frmMain.mnuPopUpEsamiPeriodici
    End If
End Sub

Private Sub flxGriglia_DblClick(Index As Integer)
    Call cmdInserisci_Click
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
    dataNascita = rsDataset("DATA_NASCITA")
    Dim somma As Integer
    If Month(dataNascita) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    lblEta = Year(date) - Year(dataNascita) + somma
    Set rsDataset = Nothing
    ' carica gli esami
    Call CaricaScheda
End Sub

