VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTerapiaDialitica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TERAPIA DIALITICA"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   50
      TabIndex        =   15
      Top             =   0
      Width           =   15100
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmTerapiaDialitica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   22
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
         Left            =   8160
         TabIndex        =   21
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
         Left            =   13680
         TabIndex        =   20
         Top             =   360
         Width           =   615
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
         Left            =   13080
         TabIndex        =   18
         Top             =   360
         Width           =   465
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
         Left            =   7320
         TabIndex        =   17
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
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   50
      TabIndex        =   6
      Top             =   720
      Width           =   15100
      Begin VB.CommandButton cmdSposta 
         Caption         =   "Sospendi terapia"
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
         Index           =   1
         Left            =   13220
         Picture         =   "frmTerapiaDialitica.frx":0459
         TabIndex        =   12
         Top             =   180
         Width           =   1695
      End
      Begin VB.TextBox txtAppo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         Left            =   5400
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.ComboBox cboMedicinali 
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
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   3615
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2535
         Left            =   60
         TabIndex        =   7
         Top             =   480
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   15
         FormatString    =   $"frmTerapiaDialitica.frx":05A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTerapiaDialitica.frx":06B0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Terapia Corrente:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   50
      TabIndex        =   10
      Top             =   3720
      Width           =   15100
      Begin VB.CommandButton cmdSposta 
         Caption         =   "Riprendi terapia"
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
         Index           =   0
         Left            =   13220
         Picture         =   "frmTerapiaDialitica.frx":080A
         TabIndex        =   13
         Top             =   180
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrigliaSospese 
         Height          =   2535
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   18
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   15
         FormatString    =   $"frmTerapiaDialitica.frx":0954
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTerapiaDialitica.frx":0A5A
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Terapia Sospesa:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   50
      TabIndex        =   5
      Top             =   6720
      Width           =   15100
      Begin VB.CheckBox chkTerapiaSospesa 
         Caption         =   "Stampa TERAPIA SOSPESA"
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
         TabIndex        =   24
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkTerapiaCorrente 
         Caption         =   "Stampa TERAPIA CORRENTE"
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
         TabIndex        =   23
         Top             =   240
         Width           =   3735
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
         Left            =   6840
         TabIndex        =   0
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
         Height          =   585
         Left            =   11520
         TabIndex        =   3
         Top             =   240
         Width           =   1815
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
         Left            =   9960
         TabIndex        =   2
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
         Height          =   615
         Left            =   8400
         TabIndex        =   1
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
         Height          =   615
         Left            =   13680
         TabIndex        =   4
         Top             =   240
         Width           =   1215
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
Attribute VB_Name = "frmTerapiaDialitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TotRecord As Integer
Dim rsTerapia As Recordset
Dim stoPulendo As Boolean
Dim vCol As Integer
Dim vRow As Integer
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim i As Integer
  
Const icsGIORNI As String = " x"
Const icsCAS As String = "  x"

Private cod_paz As Integer
Private attivaPass As Boolean

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

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("MEDICINALI", "NOME", cboMedicinali)
    
    If attivaPass Then
        intPazientiKey = cod_paz
        Call CaricaPaziente
    Else
        Select Case CaricaPazienteInAperturaForm(Me.Caption, False, intPazientiKey)
            Case tpTrovaPaziente
                Call TrovaPaziente
            Case tpCaricaPaziente
                Call CaricaPaziente
        End Select
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    ' Griglia Terapia Corrento
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 16
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
        .MousePointer = flexCustom
        .Rows = 1
    End With
    
    ' Griglia Terapia Sospese
    With flxGrigliaSospese
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 17
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
        .MousePointer = flexCustom
        .Rows = 1
    End With
    stoPulendo = False
    
    ' carica l'oggetto
    Set objAnnulla = New CAnnulla
    Call ApriRsDisconnesso
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oPazientiKey.OnClosingForm (Me.Caption)
    intPazientiKey = 0
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

Private Sub ApriRsDisconnesso()
    ' apre il recordset disconnesso per la tracciatura
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "TERAPIE_DIALITICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Set rsDisco = New ADODB.Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsDisco.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsDisco.CursorLocation = adUseClient
    rsDisco.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

Private Sub Confronta()
    ' confronta i campi per rilevare le eventuali modifiche
    ' e le salva nella relativa tabella delle modifiche
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim v_modifiche() As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nome_campi As String
    Dim valori As String
    Dim trovato As Boolean
    
    ReDim v_modifiche(0)
    ' filtra per la presenza di piu record
    rsDisco.Filter = "(KEY=" & rsTerapia("KEY") & ")"
    For i = 0 To rsDisco.Fields.count - 1
        trovato = False
        If IsNull(rsDisco(i)) Or IsNull(rsTerapia(i)) Then
            If Not (IsNull(rsDisco(i)) And IsNull(rsTerapia(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(i) <> rsTerapia(i) Then
                trovato = True
            End If
        End If
        If trovato Then
            ReDim Preserve v_modifiche(UBound(v_modifiche) + 1)
            v_modifiche(UBound(v_modifiche)) = i
        End If
    Next i
    If UBound(v_modifiche) <> 0 Then
        For i = 1 To UBound(v_modifiche)
            nome_campi = nome_campi & rsDisco.Fields((v_modifiche(i))).Name & "&-&"
            valori = valori & IIf(IsNull(rsDisco.Fields((v_modifiche(i)))) Or rsDisco.Fields((v_modifiche(i))) = "", "NULL", rsDisco.Fields((v_modifiche(i)))) & "&-&"
            ' aggiorna il rsDisco
            rsDisco(v_modifiche(i)) = rsTerapia(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "TIPO_TERAPIA", "NOME_CAMPI", "VECCHI_VALORI", "DATA_1", "DATA_2", "DATA_3")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rsTerapia("KEY"), 1, nome_campi, valori, rsTerapia("DATA_1"), rsTerapia("DATA_2"), rsTerapia("DATA_3"))
        Set rsDataset = New Recordset
        rsDataset.Open "M_TERAPIE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

Private Sub PulisciTutto()
    stoPulendo = True
    intPazientiKey = 0
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    flxGrigliaSospese.Rows = 1
    lblCognome = ""
    lblNome = ""
    lblEta = ""
    stoPulendo = False
    cmdTrova.SetFocus
End Sub

Private Sub SalvaModifiche()
    Dim valore As Variant
    Dim nome As Variant
    Dim i As Integer
    
    Select Case vCol
        Case 1
            nome = "DATA"
            valore = flxGriglia.TextMatrix(vRow, vCol)
        
        Case 2
            nome = "CODICE_MEDICINALE"
            valore = GetNumeroDaNome("MEDICINALI", "NOME", flxGriglia.TextMatrix(vRow, vCol))
        
        Case 3
            nome = "POSOLOGIA"
            valore = flxGriglia.TextMatrix(vRow, vCol)
        
        Case 4
            nome = "SOMMINISTRAZIONE"
            Select Case flxGriglia.TextMatrix(flxGriglia.Rows - 1, 4)
                Case ""
                    valore = 0
                Case "Intradialitica"
                    valore = 1
                Case "Postdialitica"
                    valore = 2
            End Select
        
        Case 5
            nome = "NOTE"
            valore = flxGriglia.TextMatrix(vRow, vCol)
        
        Case 6 To 12
            nome = "GIORNO" & vCol - 5
            valore = IIf(flxGriglia.TextMatrix(vRow, vCol) = icsGIORNI, True, False)
        
        Case 13
            nome = "TUTTI_GIORNI"
            valore = IIf(flxGriglia.TextMatrix(vRow, vCol) = icsGIORNI, True, False)
        
        Case 14, 15, 16
            nome = "DATA_" & vCol - 13
            valore = IIf(flxGriglia.TextMatrix(vRow, vCol) = "", Null, (flxGriglia.TextMatrix(vRow, vCol)))
    
    End Select
    
    Set rsTerapia = New Recordset
    rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsTerapia.Update nome, valore
    
    If TRACCIATO Then
        Call Confronta
    End If
    rsTerapia.Close
    
    If vCol = 13 And valore = True Then
        rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        For i = 1 To 7
            rsTerapia("GIORNO" & i) = False
        Next i
        rsTerapia.Update
    ElseIf vCol <> 13 Then
        rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsTerapia("TUTTI_GIORNI") = False
        rsTerapia.Update
    End If
    
    Set rsTerapia = Nothing
End Sub

Private Sub CaricaScheda()
    Dim i As Integer
    Dim strSql As String
    
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    flxGrigliaSospese.Rows = 1
    vRow = 0
    vCol = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    If intPazientiKey = 0 Then Exit Sub
    
    ' Carica le terapie correnti
    strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
            "FROM       (TERAPIE_DIALITICHE " & _
            "           INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
            "WHERE      CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
            "           SOSPESA=FALSE " & _
            "ORDER BY   DATA DESC"
    Set rsTerapia = New Recordset
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    TotRecord = rsTerapia.RecordCount 'serve x il controllo max num farmaci Helios in stampa
    
    If rsTerapia.EOF And rsTerapia.BOF Then
    Else
        ' pulisce il rsDisco
        Do While Not rsDisco.EOF
            rsDisco.Delete
            rsDisco.MoveNext
        Loop
        
        Do While Not rsTerapia.EOF
            With flxGriglia
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, 0) = rsTerapia("KEY")
                .TextMatrix(.Rows - 1, 1) = rsTerapia("DATA")
                .TextMatrix(.Rows - 1, 2) = rsTerapia("MEDICINALINOME")
                .TextMatrix(.Rows - 1, 3) = rsTerapia("POSOLOGIA")
                
                Select Case rsTerapia("SOMMINISTRAZIONE")
                    Case 0
                        .TextMatrix(.Rows - 1, 4) = ""
                    Case 1
                        .TextMatrix(.Rows - 1, 4) = "Intradialitica"
                    Case 2
                        .TextMatrix(.Rows - 1, 4) = "Postdialitica"
                End Select
                
                .TextMatrix(.Rows - 1, 5) = rsTerapia("NOTE")
                
                For i = 1 To 7
                    .TextMatrix(.Rows - 1, 5 + i) = IIf(CBool(rsTerapia("GIORNO" & i)) = True, icsGIORNI, "")
                Next i
                
                .TextMatrix(.Rows - 1, 13) = IIf(CBool(rsTerapia("TUTTI_GIORNI")) = True, icsGIORNI, "")
                
                .TextMatrix(.Rows - 1, 14) = rsTerapia("DATA_1") & ""
                .TextMatrix(.Rows - 1, 15) = rsTerapia("DATA_2") & ""
                .TextMatrix(.Rows - 1, 16) = rsTerapia("DATA_3") & ""
                
                ' aggiorna i dati nel rsDisco
                rsDisco.AddNew
                For i = 0 To rsDisco.Fields.count - 1
                    rsDisco.Fields(i) = rsTerapia.Fields(i)
                Next i
                rsDisco.Update
                
                rsTerapia.MoveNext
            End With
        Loop
    End If
    rsTerapia.Close
    flxGriglia.Row = 0
    
    ' Carica le terapie sospese
    strSql = "SELECT      TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
            "FROM       (TERAPIE_DIALITICHE " & _
            "           INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
            "WHERE      CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
            "           SOSPESA=TRUE " & _
            "ORDER BY   DATA_SOSPESA DESC"
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        Do While Not rsTerapia.EOF
            With flxGrigliaSospese
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, 0) = rsTerapia("KEY")
                .TextMatrix(.Rows - 1, 1) = rsTerapia("DATA_SOSPESA")
                .TextMatrix(.Rows - 1, 2) = rsTerapia("DATA")
                .TextMatrix(.Rows - 1, 3) = rsTerapia("MEDICINALINOME")
                .TextMatrix(.Rows - 1, 4) = rsTerapia("POSOLOGIA")
                
                Select Case rsTerapia("SOMMINISTRAZIONE")
                    Case 0
                        .TextMatrix(.Rows - 1, 5) = ""
                    Case 1
                        .TextMatrix(.Rows - 1, 5) = "Intradialitica"
                    Case 2
                        .TextMatrix(.Rows - 1, 5) = "Postdialitica"
                End Select
                
                .TextMatrix(.Rows - 1, 6) = rsTerapia("NOTE")
                
                For i = 1 To 7
                    .TextMatrix(.Rows - 1, 6 + i) = IIf(CBool(rsTerapia("GIORNO" & i)) = True, icsGIORNI, "")
                Next i
                
                .TextMatrix(.Rows - 1, 14) = IIf(CBool(rsTerapia("TUTTI_GIORNI")) = True, icsGIORNI, "")
                .TextMatrix(.Rows - 1, 15) = rsTerapia("DATA_1") & ""
                .TextMatrix(.Rows - 1, 16) = rsTerapia("DATA_2") & ""
                .TextMatrix(.Rows - 1, 17) = rsTerapia("DATA_3") & ""
                
                ' aggiorna i dati nel rsDisco
                rsDisco.AddNew
                For i = 0 To rsDisco.Fields.count - 1
                    rsDisco.Fields(i) = rsTerapia.Fields(i)
                Next i
                rsDisco.Update
                
                rsTerapia.MoveNext
            End With
        Loop
    End If
    Set rsTerapia = Nothing
End Sub

Private Sub SalvaEliminazione(flx As MSFlexGrid)
    ' salva l'eliminazione nel db di tracciature
    Dim v_nome As Variant
    Dim v_Val As Variant
    Dim rsDataset As New Recordset
    Dim sospesa As Boolean
    Dim data_sospesa As Variant
    Dim conferma As Boolean
    Dim somministrazione As Integer
    Dim i As Integer
    Dim v_giorni(1 To 8) As Boolean
    Dim colonna As Integer
    
    v_nome = Array("KEY", "CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "DATA_TERAPIA", "CODICE_MEDICINALE", "POSOLOGIA", "SOMMINISTRAZIONE", "NOTE", "GIORNO1", "GIORNO2", "GIORNO3", "GIORNO4", "GIORNO5", "GIORNO6", "GIORNO7", "TUTTI_GIORNI", "SOSPESA", "DATA_SOSPESA", "DATA_1", "DATA_2", "DATA_3")
    
    With flx
    
        If .Name = flxGrigliaSospese.Name Then
            colonna = 1
            sospesa = True
            data_sospesa = .TextMatrix(.Row, 1)
        Else
            colonna = 0
            sospesa = False
            data_sospesa = Null
        End If
        
        Select Case .TextMatrix(.Row, 4 + colonna)
            Case ""
                somministrazione = 0
            Case "Intradialitica"
                somministrazione = 1
            Case "Postdialitica"
                somministrazione = 2
        End Select
        
        For i = 1 To 8
            v_giorni(i) = IIf(.TextMatrix(.Row, 5 + colonna + i) = icsGIORNI, True, False)
        Next i

        'conferma = IIf(.TextMatrix(.Row, 12 + colonna) = icsCAS, True, False)
        v_Val = Array(GetNumeroTracciatura("E_TERAPIE_DIALITICHE"), tAccesso.key, date, Time, intPazientiKey, .TextMatrix(.Row, 1 + colonna), GetNumeroDaNome("MEDICINALI", "NOME", .TextMatrix(.Row, 2 + colonna)), .TextMatrix(.Row, 3 + colonna), somministrazione, .TextMatrix(.Row, 5 + colonna), v_giorni(1), v_giorni(2), v_giorni(3), v_giorni(4), v_giorni(5), v_giorni(6), v_giorni(7), v_giorni(8), sospesa, data_sospesa, IIf(.TextMatrix(.Row, 14 + colonna) = "", Null, .TextMatrix(.Row, 14 + colonna)), IIf(.TextMatrix(.Row, 15 + colonna) = "", Null, .TextMatrix(.Row, 15 + colonna)), IIf(.TextMatrix(.Row, 16 + colonna) = "", Null, .TextMatrix(.Row, 16 + colonna)))
    
    End With
    
    rsDataset.Open "E_TERAPIE_DIALITICHE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_nome, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

Private Function GetNumeroTracciatura(nomeTabella As String) As Integer
    Dim rsDataset As Recordset
    Dim trovato As Boolean
    Set rsDataset = New Recordset
    
    rsDataset.Open nomeTabella, cnTrac, adOpenForwardOnly, adLockReadOnly, adCmdTable
    GetNumeroTracciatura = 0
    Do
        GetNumeroTracciatura = GetNumeroTracciatura + 1
        rsDataset.Filter = "KEY=" & GetNumeroTracciatura
        If Not (rsDataset.BOF And rsDataset.EOF) Then
            trovato = True
        ElseIf rsDataset.BOF And rsDataset.EOF Then
            trovato = False
        End If
    Loop Until trovato = False

    Set rsDataset = Nothing
End Function

Private Sub cmdSposta_Click(Index As Integer)
    Dim i As Integer
    Dim num As Integer
    Dim v_bool(8) As Boolean
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
   
    Set rsTerapia = New Recordset
    If Index = 0 Then
    
    'controllo massimo numero farmaci x stampa terapia Helios
    If NumMaxFarmaci Then
        Exit Sub
    End If
        
        ' riprende il farmaco
        If flxGrigliaSospese.Row = 0 Then Exit Sub
        
        num = GetNumero("TERAPIE_DIALITICHE")
        With flxGrigliaSospese
            For i = 0 To 7
                v_bool(i) = IIf(.TextMatrix(.Row, 7 + i) = icsGIORNI, True, False)
            Next i
   '         v_bool(8) = IIf(.TextMatrix(.Row, 14) = icsCAS, True, False)
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "CODICE_MEDICINALE", "POSOLOGIA", "SOMMINISTRAZIONE", "NOTE", "GIORNO1", "GIORNO2", "GIORNO3", "GIORNO4", "GIORNO5", "GIORNO6", "GIORNO7", "TUTTI_GIORNI", "DATA_1", "DATA_2", "DATA_3")
            v_Val = Array(num, intPazientiKey, date, GetNumeroDaNome("MEDICINALI", "NOME", .TextMatrix(.Row, 3)), .TextMatrix(.Row, 4), IIf(.TextMatrix(.Row, 5) = "Intradialitica", 1, 2), .TextMatrix(.Row, 6), v_bool(0), v_bool(1), v_bool(2), v_bool(3), v_bool(4), v_bool(5), v_bool(6), v_bool(7), IIf(.TextMatrix(vRow, 15) = "", Null, .TextMatrix(vRow, 15)), IIf(.TextMatrix(vRow, 16) = "", Null, .TextMatrix(vRow, 16)), IIf(.TextMatrix(vRow, 17) = "", Null, .TextMatrix(vRow, 17)))
        End With
        
        rsTerapia.Open "TERAPIE_DIALITICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsTerapia.AddNew v_Nomi, v_Val
        rsTerapia.Update
        rsTerapia.Close
        
        Call CaricaScheda
        flxGrigliaSospese.Row = 0
    Else
        'sospende il farmaco
        If flxGriglia.Row = 0 Then Exit Sub
            
        rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE KEY=" & flxGriglia.TextMatrix(flxGriglia.Row, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        rsTerapia("SOSPESA") = True
        rsTerapia("DATA_SOSPESA") = date
        rsTerapia.Update
        If TRACCIATO Then
            Call Confronta
        End If
        rsTerapia.Close
        
        Call CaricaScheda
    End If
    
    Set rsTerapia = Nothing
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdElimina_Click()
    If intPazientiKey = 0 Then Exit Sub
    Dim eliminato As Boolean
    Dim flx As MSFlexGrid
    
    If flxGriglia.Row <> 0 Then
        Set flx = flxGriglia
    ElseIf flxGrigliaSospese.Row <> 0 Then
        Set flx = flxGrigliaSospese
    Else
        MsgBox "Selezionare il farmaco da eliminare", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If MsgBox("Sei sicuro di eliminare il farmaco dalla terapia di: " & UCase(lblCognome) & " " & UCase(lblNome) & "?", vbQuestion & vbYesNo, "Eliminazione") = vbYes Then
        Set rsTerapia = New Recordset
        rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND KEY=" & flx.TextMatrix(flx.Row, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        ' la scheda clinica non � mai stata memorizzata
        If rsTerapia.BOF And rsTerapia.EOF Then
            MsgBox "Impossibile eliminare", vbCritical, "Errore"
        Else
            eliminato = True
            rsTerapia.Delete
            TotRecord = TotRecord - 1 'serve x il controllo max num farmaci Helios in stampa
        End If
        Set rsTerapia = Nothing
        
        If eliminato And TRACCIATO Then
            Call SalvaEliminazione(flx)
        End If
        
        If eliminato Then
            ' elimina dalla flx
            If flx.Rows = 2 Then
                flx.Rows = 1
            Else
                flx.RemoveItem (flx.Row)
            End If
            vRow = 0
            flx.Row = 0
        End If
    End If
End Sub

Private Sub cmdAnnulla_Click()
    Dim Dato As String
    Dim Col As Integer
    Dim RowKey As Integer
    Dim i As Integer
    
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    RowKey = objAnnulla.Row
    
    ' cerca la riga con il key memorizzato in rowkey
    With flxGriglia
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = RowKey Then
                ' annulla
                .TextMatrix(i, Col) = Dato
                objAnnulla.Remove
                ' modifica anche il db
                vRow = i
                vCol = Col
                Call SalvaModifiche
                If objAnnulla.Vuoto = True Then
                    cmdAnnulla.Enabled = False
                End If
                Exit For
            End If
        Next i
    End With
    
End Sub

Private Sub cmdStampa_Click()
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
        Exit Sub
    End If
    
    If chkTerapiaCorrente.Value = Unchecked And chkTerapiaSospesa.Value = Unchecked Then
        MsgBox "Selezionare la TERAPIA DA STAMPARE", vbInformation, "INFORMAZIONE"
        Exit Sub
    End If
    
    Set rsTerapia = New Recordset
    rsTerapia.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsTerapia("COGNOME") & " " & rsTerapia("NOME")
    structIntestazione.sDataPaziente = rsTerapia("DATA_NASCITA")
    Set rsTerapia = Nothing
    
    ' Seleziono sia la terapia corrente che sospesa da stampare
    If chkTerapiaCorrente.Value = Checked And chkTerapiaSospesa.Value = Checked Then
        If flxGriglia.Rows = 1 And flxGrigliaSospese.Rows = 1 Then
            MsgBox "Non sono presenti terapie da stampare", vbInformation, "INFORMAZIONE"
            Exit Sub
        Else
            Call StampaSettimaParte(False, intPazientiKey)
        End If
        
    ' Seleziono solo da stampare la terapia corrente
    ElseIf chkTerapiaCorrente.Value = Checked Then
        If flxGriglia.Rows = 1 Then
            MsgBox "Non sono presenti terapie da stampare", vbInformation, "INFORMAZIONE"
            Exit Sub
        Else
            Call StampaTerapiaDialiticaCorrente(intPazientiKey)
        End If
        
    ' Seleziono solo la terapia sospesa da stampare
    ElseIf chkTerapiaSospesa.Value = Checked Then
        If flxGrigliaSospese.Rows = 1 Then
            MsgBox "Non sono presenti terapie da stampare", vbInformation, "INFORMAZIONE"
            Exit Sub
        Else
            Call StampaTerapiaDialiticaSospesa(intPazientiKey)
        End If
    End If
End Sub

Private Sub cmdInserisci_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim num As Integer
    Dim v_bool() As String
    Dim DataFarmaco1 As Variant
    Dim DataFarmaco2 As Variant
    Dim DataFarmaco3 As Variant
    
    'controllo massimo numero farmaci x stampa terapia Helios
    If NumMaxFarmaci Then
        Exit Sub
    End If
    
    If intPazientiKey = 0 Then Exit Sub
        Unload frmInput
        tInput.Tipo = tpITERAPIADIALITICA
        frmInput.Show 1
    
    If Not (tInput.v_valori(1) = "") Then
    
        v_bool = Split(tInput.v_valori(5), "-")
        
        If tInput.v_valori(7) = "" Then
            DataFarmaco1 = Null
        Else
            DataFarmaco1 = tInput.v_valori(7)
        End If
        
        If tInput.v_valori(8) = "" Then
            DataFarmaco2 = Null
        Else
            DataFarmaco2 = tInput.v_valori(8)
        End If
                
        If tInput.v_valori(9) = "" Then
            DataFarmaco3 = Null
        Else
            DataFarmaco3 = tInput.v_valori(9)
        End If
        
        num = GetNumero("TERAPIE_DIALITICHE")
        
        
        v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "CODICE_MEDICINALE", "POSOLOGIA", "SOMMINISTRAZIONE", "NOTE", "GIORNO1", "GIORNO2", "GIORNO3", "GIORNO4", "GIORNO5", "GIORNO6", "GIORNO7", "TUTTI_GIORNI", "DATA_1", "DATA_2", "DATA_3")
        v_Val = Array(num, intPazientiKey, tInput.v_valori(1), GetNumeroDaNome("MEDICINALI", "NOME", cboMedicinali.List(tInput.v_valori(2))), tInput.v_valori(3), tInput.v_valori(4), tInput.v_valori(6), _
                v_bool(0), v_bool(1), v_bool(2), v_bool(3), v_bool(4), v_bool(5), v_bool(6), v_bool(7), DataFarmaco1, DataFarmaco2, DataFarmaco3)
        
        Set rsTerapia = New Recordset
        rsTerapia.Open "TERAPIE_DIALITICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsTerapia.AddNew v_Nomi, v_Val
        rsTerapia.Update
        
        ' aggiorna i dati nel rsDisco
        rsDisco.AddNew v_Nomi, v_Val
        rsDisco.Update
        
        Set rsTerapia = Nothing
        ' aggiorna la flx
        Call CaricaScheda
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
'        MsgBox "Inserimento effettuato", vbInformation, "Inserimento"
    End If
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
        vCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        ' discolora l'altra griglia
        Call ColoraFlx(flxGrigliaSospese, flxGrigliaSospese.Cols - 1, True)
        ' annulla le row e col
        flxGrigliaSospese.Row = 0
        flxGrigliaSospese.Col = 0
    End If
End Sub

Private Sub flxGriglia_DblClick()
    Dim i As Integer
    
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        .SetFocus
        Select Case .Col
            
            Case 1      ' data
                frmCalendario.Show 1
                Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
                cmdAnnulla.Enabled = True
                .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                Call SalvaModifiche
                ' cambia colonna per evitave di ricaricare il calendario
                .Col = 0
            
            Case 2      ' medicinale
                cboMedicinali.Left = .colPos(.Col) + .Left + 45
                cboMedicinali.Top = .rowPos(.Row) + .Top + 45
                cboMedicinali.ListIndex = GetIndex(cboMedicinali, .TextMatrix(.Row, .Col))
                cboMedicinali.Visible = True
                cboMedicinali.SetFocus
            
            Case 3, 5     ' posologia, note
                If .Col = 3 Then
                    txtAppo.MaxLength = 6
                Else
                    txtAppo.MaxLength = 0
                End If
                txtAppo.Left = .colPos(.Col) + .Left + 45
                txtAppo.Top = .rowPos(.Row) + .Top + 45
                txtAppo.Width = .ColWidth(.Col)
                txtAppo.Text = .TextMatrix(.Row, .Col)
                txtAppo.Visible = True
                txtAppo.SetFocus
            
            Case 4
                Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
                cmdAnnulla.Enabled = True
                Call GestisciPosologia
                Call SalvaModifiche
            
            Case 6 To 13       ' giorni
                Call objAnnulla.Add(.TextMatrix(.Row, .Col), .Col, .TextMatrix(.Row, 0))
                cmdAnnulla.Enabled = True
                If vCol = 13 And .TextMatrix(.Row, vCol) = "" Then
                    For i = 6 To 12
                        .TextMatrix(.Row, i) = ""
                    Next i
                ElseIf vCol <> 13 And .TextMatrix(.Row, 13) = icsGIORNI Then
                    .TextMatrix(.Row, 13) = ""
                End If
                If .TextMatrix(.Row, vCol) = "" Then
                    .TextMatrix(.Row, vCol) = icsGIORNI
                Else
                    .TextMatrix(.Row, vCol) = ""
                End If
                'controlla se ci sono date nei campi DATE_1,DATE_2,DATE_3
                If .TextMatrix(.Row, 14) <> "" Or .TextMatrix(.Row, 15) <> "" Or .TextMatrix(.Row, 16) <> "" Then
                    MsgBox "INDICAZIONE DOPPIA - Eliminare dapprima le date", vbCritical, "ATTENZIONE!!!!!!"
                    .TextMatrix(.Row, vCol) = ""
                  Exit Sub
                End If
                
                Call SalvaModifiche
            
            Case 14, 15, 16     ' data del medicinale
                Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
                cmdAnnulla.Enabled = True
                
                If .TextMatrix(.Row, .Col) = "" Then
                    frmCalendario.Show 1
                    .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
                'controlla se sono spuntati giorni
                For i = 6 To 13
                    If .TextMatrix(.Row, i) <> "" Then
                        MsgBox "INDICAZIONE DOPPIA - Deselezionare dapprima i giorni", vbCritical, "ATTENZIONE!!!!!!"
                        .TextMatrix(.Row, .Col) = ""
                        Exit Sub
                    End If
                Next
                
                Call SalvaModifiche
                ' cambia colonna per evitave di ricaricare il calendario
                .Col = 0
    
        End Select
    End With
End Sub

Private Sub flxGrigliaSospese_DblClick()
    If VerificaClickFlx(flxGrigliaSospese) = False Then Exit Sub
    With flxGrigliaSospese
        .SetFocus
        If .Col = 1 Then
            frmCalendario.Show 1
            Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
            Set rsTerapia = New Recordset
            rsTerapia.Open "SELECT * FROM TERAPIE_DIALITICHE WHERE KEY=" & .TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsTerapia.Update "DATA_SOSPESA", CDate(.TextMatrix(.Row, .Col))
    
            If TRACCIATO Then
                Call Confronta
            End If
            rsTerapia.Close
            ' cambia colonna per evitave di ricaricare il calendario
            .Col = 0
        End If
    End With
End Sub

Private Sub GestisciPosologia()
    Dim testo As String
    testo = flxGriglia.TextMatrix(flxGriglia.Row, 4)
    Select Case testo
        Case Is = "Intradialitica"
            flxGriglia.TextMatrix(flxGriglia.Row, 4) = "Postdialitica"
        Case Is = "Postdialitica"
            flxGriglia.TextMatrix(flxGriglia.Row, 4) = "Intradialitica"
    End Select
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
    If cboMedicinali.Visible Then
        cboMedicinali.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub flxGrigliaSospese_Click()
    vCol = flxGrigliaSospese.Col
    flxGrigliaSospese.SetFocus
    If VerificaClickFlx(flxGrigliaSospese) = False Then
        ' discolora
        Call ColoraFlx(flxGrigliaSospese, flxGrigliaSospese.Cols - 1, True)
        ' annulla le row e col
        flxGrigliaSospese.Row = 0
        flxGrigliaSospese.Col = 0
    Else
        Call ColoraFlx(flxGrigliaSospese, flxGrigliaSospese.Cols - 1)
        flxGrigliaSospese.Col = vCol
        vRow = flxGrigliaSospese.Row
        ' discolora l'altra griglia
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    End If
End Sub

Private Sub flxGrigliaSospese_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    If Row1 = 0 Then
        Cmp = 1
    Else
        With flxGrigliaSospese
            If CDate(.TextMatrix(Row1, 1)) < CDate(.TextMatrix(Row2, 1)) Then
                Cmp = 1
            Else
                Cmp = -1
            End If
        End With
    End If
End Sub

Private Sub flxGriglia_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    If Row1 = 0 Then
        Cmp = 1
    Else
        With flxGriglia
            If CDate(.TextMatrix(Row1, 1)) < CDate(.TextMatrix(Row2, 1)) Then
                Cmp = 1
            Else
                Cmp = -1
            End If
        End With
    End If
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        ' pulisce la griglia
        ' pulisce la flx azzerando le righe
        flxGriglia.Rows = 1
    Else
        ' carica i dati del paziente
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
        
        Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
        
        Call CaricaScheda
    End If
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn = 0 Then
        Unload Me
    Else
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
    End If
End Sub

Private Sub txtAppo_LostFocus()
    If UCase(flxGriglia.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(vRow, vCol) = txtAppo
        Call SalvaModifiche
    End If
    txtAppo.Visible = False
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub cboMedicinali_Click()
    If stoPulendo Then Exit Sub
    cboMedicinali.Visible = False
End Sub

Private Sub cboMedicinali_LostFocus()
    If flxGriglia.TextMatrix(vRow, vCol) <> cboMedicinali.Text Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(flxGriglia.Row, flxGriglia.Col) = cboMedicinali.Text
        Call SalvaModifiche
    End If
    cboMedicinali.Visible = False
End Sub

Private Function NumMaxFarmaci() As Boolean
    'controllo massimo numero farmaci
    If structIntestazione.sCodiceSTS = CODICESTS_HELIOS And TotRecord >= 6 Then
       MsgBox "IMPOSSIBILE INSERIRE ALTRI FARMACI!!! - Limite massimo raggiunto ", vbCritical, "Attenzione"
       NumMaxFarmaci = True
    Else
       NumMaxFarmaci = False
    End If
End Function



