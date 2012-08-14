VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmRichiesteEsamiLab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RICHIESTA ESAMI DI LABORATORIO"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   240
         Picture         =   "frmRichiesteEsamiLab.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   450
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
         Left            =   6600
         TabIndex        =   5
         Top             =   360
         Width           =   630
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
         Left            =   960
         TabIndex        =   4
         Top             =   360
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
         Left            =   7320
         TabIndex        =   3
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
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   12495
      Begin VB.CommandButton cmdTutti 
         Caption         =   "D&eseleziona in tutti i gruppi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   10800
         TabIndex        =   21
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutti 
         Caption         =   "&Deseleziona tutti nel gruppo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   9240
         TabIndex        =   18
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutti 
         Caption         =   "&Seleziona tutti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   7800
         TabIndex        =   17
         Top             =   180
         Width           =   1335
      End
      Begin VB.ComboBox cboEsami 
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
         ItemData        =   "frmRichiesteEsamiLab.frx":0459
         Left            =   2160
         List            =   "frmRichiesteEsamiLab.frx":045B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gruppo di Esami"
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
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   285
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   12495
      Begin VB.CheckBox chkDicitura 
         Caption         =   "Stampa dicitura ""Prescrizione Esami su Unica Ricetta...."""
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
         Left            =   120
         TabIndex        =   20
         Top             =   5860
         Width           =   6255
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   5535
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmRichiesteEsamiLab.frx":045D
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
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   10320
         TabIndex        =   23
         Top             =   5805
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Richiesta"
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
         Left            =   8640
         TabIndex        =   22
         Top             =   5865
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   7440
      Width           =   12495
      Begin VB.CommandButton cmdCopiaPerPazienteSingolo 
         Caption         =   "Duplica per p&aziente"
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
         Height          =   615
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopiaPerPaziente 
         Caption         =   "Duplica per &paziente del turno"
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
         Height          =   615
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCopiaPerTuttiPazienti 
         Caption         =   "Duplica per &tutti i pazienti"
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
         Height          =   615
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopiaPerTurni 
         Caption         =   "&Duplica per turno"
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
         Height          =   615
         Left            =   7560
         TabIndex        =   14
         Top             =   240
         Width           =   1695
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
         Left            =   11040
         TabIndex        =   13
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
         Height          =   615
         Left            =   9480
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRichiesteEsamiLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmRichiesteEsamiLab.frm
'
' <b>Descrizione</b>: Scheda Richeste Esami Lab associata alla tab RICHIESTE_ESAMI
'
' @remarks: usata principalmente per la stampa
'
' @author
'
' @date 07/08/2011 11.59
Option Explicit

'' rs della scheda
Dim rsEsami As Recordset
Dim stoPulendo As Boolean
Dim periodo As Integer
Dim intPazientiKey  As Integer
Dim intMedicoKey As Integer

Const ICS As String = "       X"

'' Ricarica le cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("GRUPPI_ESAMI", "NOME", cboEsami)
    
    If intPazientiKey = 0 Then
        frmPannelloPeriodo.LetSenzaData = True
        frmPannelloPeriodo.Show 1
        periodo = frmPannelloPeriodo.GetPeriodo
        laData = frmPannelloPeriodo.getData
        Unload frmPannelloPeriodo
        If periodo = -1 Then
            Unload Me
        Else
            cmdTrova_Click
            If tTrova.keyReturn = 0 Then
                Unload Me
            End If
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
    
    stoPulendo = False
    With flxGriglia
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .Rows = 1
        .Row = 0
        For i = 1 To 8
            .Col = i
            .CellFontBold = True
        Next i
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(7) = vbLeftJustify
    End With
    
    chkDicitura.Value = GetSetting(appName, "Default", Me.Name & "." & chkDicitura.Name, 0)
    oData(0).txtBox = date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(appName, "Default", Me.Name & "." & chkDicitura.Name, chkDicitura.Value)
    intPazientiKey = 0
End Sub

'' Permette il funzionamento della rotellina del mouse nella flx
'Public Sub MouseWheel(flx As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'    Dim NewValue As Long
'    Dim Lstep As Single

'    On Error Resume Next
'    With flx
'        Lstep = .Height / .RowHeight(0)
'        Lstep = Int(Lstep)
'        If Lstep < 10 Then
'            Lstep = 10
'        End If
'        If Rotation > 0 Then
'            NewValue = .TopRow - Int(Lstep / 3)
'            If NewValue < 1 Then
'                NewValue = 1
'            End If
'        Else
'            NewValue = .TopRow + Int(Lstep / 3)
'            If NewValue > .Rows - 1 Then
'                NewValue = .Rows - 1
'            End If
'        End If
'        .TopRow = NewValue
'    End With
'End Sub
'-----------------------------------------


'' Carica la scheda del gruppo di esame e delle data selezionata
Private Sub CaricaScheda()
    Dim rigaEsame As Integer
    Dim colonnaEsame As Integer
    Dim Col As Integer
    Dim riga As Integer
    
    With flxGriglia
        If cboEsami.ListIndex = -1 Then
            Exit Sub
        End If
        ' pulisce
        .Rows = 1
        
        flxGriglia.TextMatrix(0, 0) = cboEsami.ItemData(cboEsami.ListIndex)
        ' carica le voci
        Set rsEsami = New Recordset
        rsEsami.Open "SELECT V.NOME, A.KEY FROM (ASSOCIAZIONE_ESAMI_LAB A INNER JOIN VOCI_ESAMI V ON V.KEY=A.CODICE_ESAME) WHERE CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 0) & " ORDER BY ORDINE_VISUALIZZAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Col = 1
        If Not (rsEsami.EOF And rsEsami.BOF) Then
            With flxGriglia
                .Rows = Int(rsEsami.RecordCount / 3 + 1) + 1
                riga = 0
                Do While Not rsEsami.EOF
                    riga = riga + 1
                    .TextMatrix(riga, Col - 1) = rsEsami("KEY")
                    .TextMatrix(riga, Choose(Col, 3, 5, 7)) = rsEsami("NOME")
                    If riga = .Rows - 1 Then
                        riga = 0
                        Col = Col + 1
                    End If
                    rsEsami.MoveNext
                Loop
            End With
        End If
        rsEsami.Close
        
        rsEsami.Open "SELECT A.KEY FROM (RICHIESTE_ESAMI R INNER JOIN ASSOCIAZIONE_ESAMI_LAB A ON A.KEY=R.CODICE_ASSOCIAZIONE) WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 0), cnPrinc, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not (rsEsami.EOF And rsEsami.BOF) Then
            Do While Not rsEsami.EOF
                Call getRigaColEsame(rsEsami("KEY"), rigaEsame, colonnaEsame)
                flxGriglia.TextMatrix(rigaEsame, Choose(colonnaEsame + 1, 4, 6, 8)) = ICS
                rsEsami.MoveNext
            Loop
        End If
        rsEsami.Close
    End With
End Sub

'' Restituisce il numero di riga e colonna dove è presente il codiceAssociazione
Private Sub getRigaColEsame(codiceAssociazione As Integer, ByRef riga As Integer, ByRef colonna As Integer)
    Dim i As Integer
    Dim k As Integer
    For i = 1 To flxGriglia.Rows - 1
        For k = 0 To 2
            If flxGriglia.TextMatrix(i, k) <> "" Then
                If flxGriglia.TextMatrix(i, k) = codiceAssociazione Then
                    riga = i
                    colonna = k
                    Exit Sub
                End If
            End If
        Next k
    Next i
End Sub

'' Elimina gli esami dei pazienti prima di effettuare i vari duplica esami...
Private Sub EliminaRichiesteEsami(condizione As String)
    Dim cmCommand As New Command
    
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    cmCommand.CommandText = "DELETE * FROM RICHIESTE_ESAMI WHERE (NOT CODICE_PAZIENTE=" & intPazientiKey & ") AND CODICE_PAZIENTE IN (" & condizione & ")"
    cmCommand.Execute
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    stoPulendo = True
    intPazientiKey = 0
    flxGriglia.Rows = 1
    lblCognome = ""
    lblNome = ""
    cboEsami.ListIndex = -1
    stoPulendo = False
    cmdCopiaPerPaziente.Enabled = False
    cmdCopiaPerPazienteSingolo.Enabled = False
    cmdCopiaPerTuttiPazienti.Enabled = False
    cmdCopiaPerTurni.Enabled = False
    cmdTrova.SetFocus
End Sub

Private Sub EliminaTuttiGliEsami()
    Dim cmCommand As New Command
    
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    
    cmCommand.CommandText = "Delete FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey
    cmCommand.Execute
    
    Set cmCommand = Nothing
End Sub

Private Sub EliminaEsame(riga As Integer, colonna As Integer)
    Dim Col As Integer
    
    Select Case colonna
        Case 4: Col = 0
        Case 6: Col = 1
        Case 8: Col = 2
    End Select
    
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_ASSOCIAZIONE=" & flxGriglia.TextMatrix(riga, Col), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsEsami.Delete
    rsEsami.Update
    rsEsami.Close
    Set rsEsami = Nothing
End Sub

Private Sub SalvaModifiche(riga As Integer, colonna As Integer)
    Dim Col As Integer
    
    Select Case colonna
        Case 4: Col = 0
        Case 6: Col = 1
        Case 8: Col = 2
    End Select
    
    Set rsEsami = New Recordset
    rsEsami.Open "RICHIESTE_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsEsami.AddNew
    rsEsami("KEY") = GetNumero("RICHIESTE_ESAMI")
    rsEsami("CODICE_PAZIENTE") = intPazientiKey
    rsEsami("CODICE_ASSOCIAZIONE") = flxGriglia.TextMatrix(riga, Col)
    rsEsami.Update
    rsEsami.Close
    Set rsEsami = Nothing
End Sub

Private Function CreaCondizione() As String
    ' crea la condizione del form Trova
    ' fa caricare solo i pazienti che hanno il turno dialitico
    ' oggi (o pom o mat)
    Dim rsAppo As New Recordset
    Dim giorno As Integer       ' 1 lun 2 mart 3 merc ..
    Dim rsPazientiTurni As Recordset
    Dim strTurno As String
    
    Select Case periodo
        Case 1
            strTurno = "AM"
        Case 2
            strTurno = "PM"
        Case 3
            strTurno = "SR"
    End Select
        
    giorno = Weekday(laData, vbMonday)
    Set rsPazientiTurni = New Recordset
    rsPazientiTurni.Open "SELECT P.KEY, T." & strTurno & "_INIZIO" & giorno & " " & _
                         "FROM ((PAZIENTI P INNER JOIN TURNI T ON P.KEY = T.CODICE_PAZIENTE) INNER JOIN RENI R ON R.KEY=T.CODICE_RENE) " & _
                         " WHERE  NOT ( T." & strTurno & "_INIZIO" & giorno & "="""")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsAppo.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Do While Not rsPazientiTurni.EOF
        ' es: KEY IN (1,2,4)
        ' effettua il controllo sulla data fine nnn in query perche il campo nn è obbligatorio
        rsAppo.Filter = ("CODICE_PAZIENTE=" & rsPazientiTurni("KEY"))
        ' se nn esiste nn  puo effettuare la dialisi
        If Not (rsAppo.BOF And rsAppo.EOF) Then
            If CDate(rsAppo("DATA_INIZIO")) <= date Then
                If rsAppo("DATA_FINE") <> "" Then
                    If CDate(rsAppo("DATA_FINE")) >= date Then
                        CreaCondizione = CreaCondizione & rsPazientiTurni("KEY") & ","
                    End If
                Else
                    CreaCondizione = CreaCondizione & rsPazientiTurni("KEY") & ","
                End If
            End If
        End If
        rsPazientiTurni.MoveNext
    Loop
    If CreaCondizione <> "" Then
        ' elimina la , finale e aggiunge le parentesi
        CreaCondizione = Left(CreaCondizione, Len(CreaCondizione) - 1)
        CreaCondizione = " KEY IN (" & CreaCondizione & ")"
    Else
        ' non deve trovare nessun paziente (key=-1 piezzo)
        CreaCondizione = " KEY IN (-1)"
    End If
    ' solo quelli in dialisi
    CreaCondizione = CreaCondizione & " AND (STATO=0 OR STATO=4)"
    Set rsAppo = Nothing
    Set rsPazientiTurni = Nothing
End Function

Private Sub Stampa(codicePaziente As Integer)
    Const numMaxEsami As Integer = 19
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape

    Dim numCol As Integer
    Dim numPag As Integer
    Dim numEsami As Integer
    Dim nomePaziente As String
    Dim nomeMese As String

    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
        Exit Sub
    End If

    SQLString = "SHAPE APPEND " & _
                "       NEW adInteger AS INDEX, " & _
                "       NEW adVarChar(50) AS ESAME1, " & _
                "       NEW adVarChar(50) AS ESAME2 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT NOME, COGNOME FROM PAZIENTI WHERE KEY=" & codicePaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomePaziente = rsEsami("COGNOME") & " " & rsEsami("NOME")
    rsEsami.Close
    
    rsEsami.Open "SELECT  V.NOME AS ESAME FROM ((RICHIESTE_ESAMI R INNER JOIN ASSOCIAZIONE_ESAMI_LAB A ON A.KEY=R.CODICE_ASSOCIAZIONE) INNER JOIN VOCI_ESAMI V ON V.KEY=A.CODICE_ESAME) WHERE R.CODICE_PAZIENTE=" & codicePaziente & " ORDER BY CODICE_GRUPPO, ORDINE_VISUALIZZAZIONE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsEsami.EOF And rsEsami.BOF) Then
        numCol = 1
        numPag = 1
        numEsami = 0
        Do While Not rsEsami.EOF
            With rsMain
                If numCol = 1 Then
                    .AddNew
                    .Fields("INDEX") = numPag
                Else
                    If numEsami <> 0 Then .MoveNext
                End If
                .Fields("ESAME" & numCol) = "- " & rsEsami("ESAME")
                'Debug.Print "NUM PAG " & numPag & "   NUM COL " & numCol & "   ESAME " & rsEsami("ESAME")
                numEsami = numEsami + 1
                If numEsami = numMaxEsami Then
                    numEsami = 0
                    If numCol = 1 Then
                        .Filter = ("INDEX=" & numPag)
                        numCol = 2
                    Else
                        .Filter = ""
                        numPag = numPag + 1
                        numCol = 1
                    End If
                End If
        
            End With
            rsEsami.MoveNext
        Loop
        rsMain.UpdateBatch
        
        nomeMese = MonthName(Month(Now))

        Set rptRichiestaEsamiLaboratorio.DataSource = rsMain
        rptRichiestaEsamiLaboratorio.Sections("intestazione").Controls.Item("lblTitolo").Caption = "Si richiedono i seguenti esami relativi al mese di " & UCase(Left(nomeMese, 1)) & Right(nomeMese, Len(nomeMese) - 1) & " " & Year(Now)
        rptRichiestaEsamiLaboratorio.Sections("intestazione").Controls.Item("lblPaziente").Caption = nomePaziente
        rptRichiestaEsamiLaboratorio.Sections("pie").Controls.Item("lblDicitura").Visible = IIf(chkDicitura.Value = Unchecked, False, True)
        rptRichiestaEsamiLaboratorio.Sections("pie").Controls.Item("lblData").Caption = "Data, " & oData(0).txtBox
        rptRichiestaEsamiLaboratorio.PrintReport False, rptRangeAllPages
    Else
        MsgBox "Nessun esame assegnato per il paziente " & nomePaziente, vbInformation, "Stampa richiesta esami"
    End If
    
    Set rsEsami = Nothing
    
End Sub

Private Sub StampaModuloSodav(codicePaziente As Integer)
    Const numMaxEsami As Integer = 16
    Dim SQLString As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    Dim numCol As Integer
    Dim numPag As Integer
    Dim numEsami As Integer
    Dim nomePaziente As String
    Dim dataNascita As String
    Dim SintesiStatoClinico As String
    Dim MedicoDiBase As String
    Dim MedicoDiTurno As String
    Dim DataSintesiStatoClinico As String
    Dim nomeMese As String
    

    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
        Exit Sub
    End If

    SQLString = "SHAPE APPEND " & _
                "       NEW adInteger AS INDEX, " & _
                "       NEW adVarChar(50) AS ESAME1, " & _
                "       NEW adVarChar(50) AS ESAME2, " & _
                "       NEW adDate AS DATA, " & _
                "       NEW adLongVarChar AS SINTESI_STATO_CLINICO "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT NOME, COGNOME, DATA_NASCITA, CODICE_MEDICO FROM PAZIENTI WHERE KEY=" & codicePaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    nomePaziente = rsEsami("COGNOME") & " " & rsEsami("NOME")
    dataNascita = rsEsami("DATA_NASCITA")
    intMedicoKey = rsEsami("CODICE_MEDICO")
    rsEsami.Close
    
    If intMedicoKey = 0 Then
        MedicoDiBase = "- -"
    Else
    rsEsami.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMedicoKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    MedicoDiBase = rsEsami("COGNOME") & " " & rsEsami("NOME")
    rsEsami.Close
    End If
    
    If tAccesso.Tipo = tpAMEDICO Then
        rsEsami.Open "SELECT * FROM MEDICI_DIALISI WHERE ELIMINATO=0", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsEsami.EOF
                MedicoDiTurno = UCase(rsEsami("COGNOME"))
                If MedicoDiTurno = UCase(tAccesso.cognome) Then
                    MedicoDiTurno = rsEsami("COGNOME") & " " & rsEsami("NOME") & " " & vbCrLf & "N. Iscrizione Albo " & rsEsami("CODICE_ALBO")
                    Exit Do
                End If
                rsEsami.MoveNext
            Loop
        rsEsami.Close
        Else
            MedicoDiTurno = ""
    End If
    
    rsEsami.Open "SELECT  V.NOME AS ESAME FROM ((RICHIESTE_ESAMI R INNER JOIN ASSOCIAZIONE_ESAMI_LAB A ON A.KEY=R.CODICE_ASSOCIAZIONE) INNER JOIN VOCI_ESAMI V ON V.KEY=A.CODICE_ESAME) WHERE R.CODICE_PAZIENTE=" & codicePaziente & " ORDER BY CODICE_GRUPPO, ORDINE_VISUALIZZAZIONE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsEsami.EOF And rsEsami.BOF) Then
        numCol = 1
        numPag = 1
        numEsami = 0
        Do While Not rsEsami.EOF
            With rsMain
                If numCol = 1 Then
                    .AddNew
                    .Fields("INDEX") = numPag
                Else
                    If numEsami <> 0 Then .MoveNext
                End If
                .Fields("ESAME" & numCol) = "- " & rsEsami("ESAME")
                'Debug.Print "NUM PAG " & numPag & "   NUM COL " & numCol & "   ESAME " & rsEsami("ESAME")
                numEsami = numEsami + 1
                If numEsami = numMaxEsami Then
                    numEsami = 0
                    If numCol = 1 Then
                        .Filter = ("INDEX=" & numPag)
                        numCol = 2
                    Else
                        .Filter = ""
                        numPag = numPag + 1
                        numCol = 1
                    End If
                End If
    
        ' Ricerca Sintesi Stato Clinico
        Set rsDataset = New Recordset
            strSql = "Select    Top 1 * " & _
                "From      DIARI_CLINICI " & _
                "          INNER JOIN TITOLI_DIARIO ON TITOLI_DIARIO.KEY=DIARI_CLINICI.CODICE_TITOLO " & _
                "Where     CODICE_PAZIENTE=" & intPazientiKey & " AND TITOLI_DIARIO.NOME LIKE '%SINTESI%' " & _
                "Order By  DATA DESC"
            rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    DataSintesiStatoClinico = rsDataset("DATA")
                    SintesiStatoClinico = rsDataset("DATI")
                Else
                    DataSintesiStatoClinico = "- -"
                End If
        rsDataset.Close
               
        End With
        rsEsami.MoveNext
        Loop
        rsMain.UpdateBatch
        
        Set rsDataset = Nothing
        
        nomeMese = MonthName(Month(Now))

        Set rptRichiestaEsamiModuloSodav.DataSource = rsMain
        rptRichiestaEsamiModuloSodav.Sections("intestazione").Controls.Item("lblMese").Caption = UCase(Left(nomeMese, 1)) & Right(nomeMese, Len(nomeMese) - 1)
        rptRichiestaEsamiModuloSodav.Sections("intestazione").Controls.Item("lblData").Caption = oData(0).txtBox
        rptRichiestaEsamiModuloSodav.Sections("intestazione").Controls.Item("lblMedicoDiBase").Caption = MedicoDiBase
        rptRichiestaEsamiModuloSodav.Sections("intestazione").Controls.Item("lblPaziente").Caption = nomePaziente
        rptRichiestaEsamiModuloSodav.Sections("intestazione").Controls.Item("lblDataNascita").Caption = dataNascita
        rptRichiestaEsamiModuloSodav.Sections("pie").Controls.Item("lblDataSintesiStatoClinico").Caption = DataSintesiStatoClinico
        rptRichiestaEsamiModuloSodav.Sections("pie").Controls.Item("lblSintesiStatoClinico").Caption = SintesiStatoClinico
        rptRichiestaEsamiModuloSodav.Sections("pie").Controls.Item("lblDicitura").Visible = IIf(chkDicitura.Value = Unchecked, False, True)
        rptRichiestaEsamiModuloSodav.Sections("pie").Controls.Item("lblMedico").Caption = MedicoDiTurno
        
        rptRichiestaEsamiModuloSodav.PrintReport False, rptRangeAllPages
    Else
        MsgBox "Nessun esame assegnato per il paziente " & nomePaziente, vbInformation, "Stampa richiesta esami"
    End If
    
    Set rsEsami = Nothing
    
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
    Set rsDataset = Nothing
    ' cerca i riferimenti al paziente solo dopo aver selezionato l'esame
End Sub

'' Duplica gli esami del paziente al paziente selezionato
Private Sub cmdCopiaPerPazienteSingolo_Click()
    On Error GoTo gestione
    Dim rsAppo As New Recordset
    Dim key As Integer
    Dim nomePaziente As String
    
    If MsgBox("Sicuro di voler duplicare gli esami ad un paziente in dialisi?", vbQuestion + vbYesNo, "Duplica per paziente in dialisi") = vbYes Then
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = "STATO=0 AND NOT KEY=" & intPazientiKey
        tTrova.condStato = "(-1)"
        frmTrova.Show 1
        key = tTrova.keyReturn
        If key = 0 Then Exit Sub
        Set rsEsami = New Recordset
        If MsgBox("ATTENZIONE!!! LA DUPLICAZIONE SOSTITUIRA' TUTTI GLI ESAMI PRESCRITTI" & vbCrLf & "Sicuro di volerli duplicare per un altro paziente in dialisi?", vbQuestion + vbYesNo, "Duplica per paziente in dialisi") = vbYes Then
            Call EliminaRichiesteEsami(Str(key))
            rsEsami.Open "RICHIESTE_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAppo.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                Do While Not rsAppo.EOF
                    rsEsami.Filter = "CODICE_PAZIENTE=" & key & " AND CODICE_ASSOCIAZIONE=" & rsAppo("CODICE_ASSOCIAZIONE")
                    If (rsEsami.EOF And rsEsami.BOF) Then
                        rsEsami.AddNew
                        rsEsami("KEY") = GetNumero("RICHIESTE_ESAMI")
                        rsEsami("CODICE_PAZIENTE") = key
                        rsEsami("CODICE_ASSOCIAZIONE") = rsAppo("CODICE_ASSOCIAZIONE")
                        rsEsami.Update
                    End If
                    rsAppo.MoveNext
                Loop
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
                Set rsEsami = Nothing
                Exit Sub
            End If
            rsAppo.Close
            rsEsami.Close
        Else
            Set rsEsami = Nothing
            Exit Sub
        End If
        Set rsEsami = Nothing
        
        rsAppo.Open "SELECT NOME, COGNOME FROM PAZIENTI WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        nomePaziente = rsAppo("COGNOME") & " " & rsAppo("NOME")
        rsAppo.Close
        
        If MsgBox("DUPLICAZIONE AVVENUTA CON SUCCESSO" & vbCrLf & "Stampare gli esami del paziente " & nomePaziente & "?", vbQuestion + vbYesNo, "Stampa esami") = vbYes Then
            cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
            cdlStampa.CancelError = True
            cdlStampa.ShowPrinter
            
            Call Stampa(key)
        End If
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

'' Duplica gli esami del paziente al paziente selezionato in turno
Private Sub cmdCopiaPerPaziente_Click()
    On Error GoTo gestione
    Dim rsAppo As New Recordset
    Dim key As Integer
    Dim nomePaziente As String
    
    If MsgBox("Sicuro di voler duplicare gli esami ad un paziente del turno?", vbQuestion + vbYesNo, "Duplica per paziente in turno") = vbYes Then
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = CreaCondizione & " AND NOT KEY=" & intPazientiKey
        tTrova.condStato = "(-1)"
        frmTrova.Show 1
        key = tTrova.keyReturn
        If key = 0 Then Exit Sub
        Set rsEsami = New Recordset
        If MsgBox("ATTENZIONE!!! LA DUPLICAZIONE SOSTITUIRA' TUTTI GLI ESAMI PRESCRITTI" & vbCrLf & "Sicuro di volerli duplicare per un altro paziente del turno?", vbQuestion + vbYesNo, "Duplica per paziente in turno") = vbYes Then
            Call EliminaRichiesteEsami(Str(key))
            rsEsami.Open "RICHIESTE_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAppo.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                Do While Not rsAppo.EOF
                    rsEsami.Filter = "CODICE_PAZIENTE=" & key & " AND CODICE_ASSOCIAZIONE=" & rsAppo("CODICE_ASSOCIAZIONE")
                    If (rsEsami.EOF And rsEsami.BOF) Then
                        rsEsami.AddNew
                        rsEsami("KEY") = GetNumero("RICHIESTE_ESAMI")
                        rsEsami("CODICE_PAZIENTE") = key
                        rsEsami("CODICE_ASSOCIAZIONE") = rsAppo("CODICE_ASSOCIAZIONE")
                        rsEsami.Update
                    End If
                    rsAppo.MoveNext
                Loop
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
                Set rsEsami = Nothing
                Exit Sub
            End If
            rsAppo.Close
            rsEsami.Close
        Else
            Set rsEsami = Nothing
            Exit Sub
        End If
        Set rsEsami = Nothing
        
        rsAppo.Open "SELECT NOME, COGNOME FROM PAZIENTI WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        nomePaziente = rsAppo("COGNOME") & " " & rsAppo("NOME")
        rsAppo.Close
        
        If MsgBox("DUPLICAZIONE AVVENUTA CON SUCCESSO" & vbCrLf & "Stampare gli esami del paziente " & nomePaziente & "?", vbQuestion + vbYesNo, "Stampa esami") = vbYes Then
            cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
            cdlStampa.CancelError = True
            cdlStampa.ShowPrinter
            
            Call Stampa(key)
        End If
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

'' Duplica gli esami del paziente a tutti gli altri pazienti del turno
Private Sub cmdCopiaPerTurni_Click()
    On Error GoTo gestione
    Dim rsAppo As New Recordset
    Dim v_key() As String
    Dim i As Integer
    Dim testo As String
    
    If MsgBox("Sicuro di voler duplicare gli esami a tutti i pazienti del turno?", vbQuestion + vbYesNo, "Duplica per pazienti in turno") = vbYes Then
        testo = CreaCondizione
        testo = Mid(testo, 10, InStr(1, testo, ")") - 10)
        v_key = Split(testo, ",")
        Set rsEsami = New Recordset
        If MsgBox("ATTENZIONE!!! LA DUPLICAZIONE SOSTITUIRA' TUTTI GLI ESAMI PRESCRITTI" & vbCrLf & "Sicuro di volerli duplicare per tutti i pazienti del turno?", vbQuestion + vbYesNo, "Duplica per pazienti in turno") = vbYes Then
            Call EliminaRichiesteEsami(testo)
            rsEsami.Open "RICHIESTE_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAppo.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                Call StartProgressBar(UBound(v_key), 0, Me)
                
                For i = 0 To UBound(v_key)
                    If v_key(i) <> intPazientiKey Then
                        rsAppo.MoveFirst
                        frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
                        Do While Not rsAppo.EOF
                            rsEsami.Filter = "CODICE_PAZIENTE=" & v_key(i) & " AND CODICE_ASSOCIAZIONE=" & rsAppo("CODICE_ASSOCIAZIONE")
                            If (rsEsami.EOF And rsEsami.BOF) Then
                                rsEsami.AddNew
                                rsEsami("KEY") = GetNumero("RICHIESTE_ESAMI")
                                rsEsami("CODICE_PAZIENTE") = v_key(i)
                                rsEsami("CODICE_ASSOCIAZIONE") = rsAppo("CODICE_ASSOCIAZIONE")
                                rsEsami.Update
                            End If
                            rsAppo.MoveNext
                        Loop
                    End If
                Next i
                
                Call StopProgressBar(Me)
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
                Set rsEsami = Nothing
                Exit Sub
            End If
            rsAppo.Close
            rsEsami.Close
        Else
            Set rsEsami = Nothing
            Exit Sub
        End If
        Set rsEsami = Nothing
        
        If MsgBox("DUPLICAZIONE AVVENUTA CON SUCCESSO" & vbCrLf & "Stampare gli esami di tutti i pazienti del turno?", vbQuestion + vbYesNo, "Stampa esami") = vbYes Then
            cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
            cdlStampa.CancelError = True
            cdlStampa.ShowPrinter
            
            For i = 0 To UBound(v_key)
                Call Stampa(CInt(v_key(i)))
            Next i
        End If
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
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
            Call EliminaRichiesteEsami("SELECT KEY FROM PAZIENTI WHERE (STATO=0)")
            rsEsami.Open "RICHIESTE_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAppo.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                rsPazienti.Open "SELECT KEY FROM PAZIENTI WHERE (STATO=0) AND KEY<>" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                Call StartProgressBar(rsPazienti.RecordCount, 0, Me)
                
                Do While Not rsPazienti.EOF
                    frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
                    rsAppo.MoveFirst
                    Do While Not rsAppo.EOF
                        rsEsami.Filter = "CODICE_PAZIENTE=" & rsPazienti("KEY") & " AND CODICE_ASSOCIAZIONE=" & rsAppo("CODICE_ASSOCIAZIONE")
                        If (rsEsami.EOF And rsEsami.BOF) Then
                            rsEsami.AddNew
                            rsEsami("KEY") = GetNumero("RICHIESTE_ESAMI")
                            rsEsami("CODICE_PAZIENTE") = rsPazienti("KEY")
                            rsEsami("CODICE_ASSOCIAZIONE") = rsAppo("CODICE_ASSOCIAZIONE")
                            rsEsami.Update
                        End If
                        rsAppo.MoveNext
                    Loop
                    rsPazienti.MoveNext
                Loop
                rsPazienti.Close
                
                Call StopProgressBar(Me)
            Else
                MsgBox "Il paziente selezionato non ha esami da duplicare", vbCritical, "Attenzione"
                Set rsEsami = Nothing
                Exit Sub
            End If
            rsAppo.Close
            rsEsami.Close
        Else
            Set rsEsami = Nothing
            Exit Sub
        End If
        Set rsEsami = Nothing
    
        If MsgBox("DUPLICAZIONE AVVENUTA CON SUCCESSO" & vbCrLf & "Stampare gli esami di tutti i pazienti in dialisi?", vbQuestion + vbYesNo, "Stampa esami") = vbYes Then
            cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
            cdlStampa.CancelError = True
            cdlStampa.ShowPrinter
            
            rsAppo.Open "SELECT KEY FROM PAZIENTI WHERE STATO=0 ORDER BY COGNOME, NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsAppo.EOF
                Call Stampa(rsAppo("KEY"))
                rsAppo.MoveNext
            Loop
            rsAppo.Close
        End If
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

Private Sub cmdTutti_Click(Index As Integer)
    Dim i As Integer
    
    With flxGriglia
        If Index = 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 4) <> ICS Then
                    .TextMatrix(i, 4) = ICS
                    Call SalvaModifiche(i, 4)
                End If
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 6) <> ICS Then
                    .TextMatrix(i, 6) = ICS
                    Call SalvaModifiche(i, 6)
                End If
                If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 8) <> ICS Then
                    .TextMatrix(i, 8) = ICS
                    Call SalvaModifiche(i, 8)
                End If
            Next i
        ElseIf Index = 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) <> "" And .TextMatrix(i, 4) = ICS Then
                    .TextMatrix(i, 4) = ""
                    Call EliminaEsame(i, 4)
                End If
                If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 6) = ICS Then
                    .TextMatrix(i, 6) = ""
                    Call EliminaEsame(i, 6)
                End If
                If .TextMatrix(i, 2) <> "" And .TextMatrix(i, 8) = ICS Then
                    .TextMatrix(i, 8) = ""
                    Call EliminaEsame(i, 8)
                End If
            Next i
        Else
            For i = 1 To .Rows - 1
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 6) = ""
                .TextMatrix(i, 8) = ""
            Next i
            Call EliminaTuttiGliEsami
        End If
    End With
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = CreaCondizione
    tTrova.condStato = "(-1)"
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    Call CaricaPaziente
End Sub

Private Sub cmdStampa_Click()
    
    If oData(0).txtBox = "" Then
        MsgBox "Selezionare la Data", vbInformation, "Informazione"
        Exit Sub
    End If
    
    On Error GoTo gestione
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    If structIntestazione.sCodiceSTS = CODICESTS_SODAV Then
        Call StampaModuloSodav(intPazientiKey)
    Else
        Call Stampa(intPazientiKey)
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

Private Sub cboEsami_Click()
    If stoPulendo Then Exit Sub
    Call CaricaScheda
    cmdCopiaPerPaziente.Enabled = True
    cmdCopiaPerPazienteSingolo.Enabled = True
    cmdCopiaPerTuttiPazienti.Enabled = True
    cmdCopiaPerTurni.Enabled = True
End Sub

Private Sub flxGriglia_DblClick()
    With flxGriglia
        .SetFocus
        If .Col = 4 Or .Col = 6 Or .Col = 8 Then
            If .TextMatrix(.Row, .Col) = ICS Then
                .TextMatrix(.Row, .Col) = ""
                Call EliminaEsame(.Row, .Col)
            Else
                If .TextMatrix(.Row, (.Col - 4) / 2) <> "" Then
                    .TextMatrix(.Row, .Col) = ICS
                    Call SalvaModifiche(.Row, .Col)
                End If
            End If
        End If
    End With
End Sub

'Private Sub flxGriglia_GotFocus()
'    Call WheelHook(Me, flxGriglia)
'End Sub

'Private Sub flxGriglia_LostFocus()
'    Call WheelUnHook
'End Sub
'-----------------------

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub


