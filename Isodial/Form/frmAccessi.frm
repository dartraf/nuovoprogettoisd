VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmAccessi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ACCESSI VASCOLARI"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmAccessi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Seleziona il paziente"
         Top             =   240
         Width           =   450
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   360
         Width           =   465
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
         Left            =   6840
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
         Left            =   11160
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   12015
      Begin VB.OptionButton optAnestesia 
         Caption         =   "Locale"
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
         Left            =   2280
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optAnestesia 
         Caption         =   "Locoregionale"
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
         Left            =   3600
         TabIndex        =   6
         Top             =   2640
         Width           =   1935
      End
      Begin VB.OptionButton optAnestesia 
         Caption         =   "Totale"
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
         Left            =   5760
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtDati 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3000
         Width           =   9495
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   1680
         Picture         =   "frmAccessi.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Seleziona l' operatore"
         Top             =   1560
         Width           =   450
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   2
         Left            =   1680
         Picture         =   "frmAccessi.frx":08B2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Seleziona l' operatore"
         Top             =   2040
         Width           =   450
      End
      Begin VB.TextBox txtIntervento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   9495
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   175
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anestesia"
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
         Index           =   11
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dati Rilevanti"
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
         Index           =   12
         Left            =   240
         TabIndex        =   39
         Top             =   3000
         Width           =   1410
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
         Index           =   7
         Left            =   2280
         TabIndex        =   38
         Top             =   1680
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
         Index           =   6
         Left            =   7680
         TabIndex        =   37
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operatore 1"
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
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1245
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
         Index           =   8
         Left            =   2280
         TabIndex        =   35
         Top             =   2160
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
         Index           =   9
         Left            =   7680
         TabIndex        =   34
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operatore 2"
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
         Index           =   10
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label lblCognomeMed 
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
         Index           =   0
         Left            =   3600
         TabIndex        =   32
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label lblNomeMed 
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
         Index           =   0
         Left            =   8640
         TabIndex        =   31
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label lblCognomeMed 
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
         Index           =   1
         Left            =   3600
         TabIndex        =   30
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label lblNomeMed 
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
         Index           =   1
         Left            =   8640
         TabIndex        =   29
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Intervento"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrizione Intervento"
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
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1485
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   12015
      Begin VB.TextBox txtCausaChiusuraAccesso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   9495
      End
      Begin DataTimeBox.uDataTimeBox oDataChiusuraAccesso 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   175
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Causa Chiusura Accesso"
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
         Index           =   14
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   1845
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Data Chiusura Accesso"
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
         Index           =   13
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   12015
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
         Height          =   495
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   1335
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
         Left            =   5520
         TabIndex        =   11
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
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAccessi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAccessi.frm
'
' <b>Descrizione</b>: Scheda Accessi Vascolari associata alla tab ACCESSI_VASCOLARI_TAB
'
' @remarks
'
' @author
'
' @date 31/01/2011 15.40
Option Explicit

'' indica se si è in fase di modifica
Dim modifica As Boolean
'' rs della scheda
Dim rsAccessi As Recordset
'' key del record aperto
Dim keyId As Integer
'' rs per la tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim intMedicoKey(1) As Integer
Dim blnModificato As Boolean

'' Apre frmTrova se non c'è nessun paziente gia caricato
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Select Case CaricaPazienteInAperturaForm(Me.Caption, blnModificato, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
    End Select
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    modifica = False
    intPazientiKey = 0
    
    Call ApriRsDisconnesso
    oData.ConnectionString = strConnectionStringCentro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        intMedicoKey(0) = 0
        intMedicoKey(1) = 0
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click (0)
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Apre il rs disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ACCESSI_VASCOLARI_TAB", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Set rsDisco = New ADODB.Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsDisco.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsDisco.CursorLocation = adUseClient
    rsDisco.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

'' Confronta i campi per rilevare le eventuali modifiche
' e le salva nella relativa tabella delle modifiche
'
' @param rs rs che contiene lo stato del record che si è memorizzato
Private Sub Confronta(rs As Recordset)
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim v_modifiche() As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nome_campi As String
    Dim valori As String
    Dim trovato As Boolean
    
    ReDim v_modifiche(0)
    For i = 0 To rsDisco.Fields.count - 1
        trovato = False
        If IsNull(rsDisco(i)) Or IsNull(rs(i)) Then
            If Not (IsNull(rsDisco(i)) And IsNull(rs(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(i) <> rs(i) Then
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
            rsDisco(v_modifiche(i)) = rs(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "DATA_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), oData.data, nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_ACCESSI_VASCOLARI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Carica i dati della scheda dalla tabella nel form e nel rs per la tracciatura
Private Sub CaricaScheda()
    Dim i As Integer
    Dim data As Date
    Dim strSql As String
    
    ' la data americana
    data = Month(oData.data) & "/" & Day(oData.data) & "/" & Year(oData.data)
    
    'data = oData.DataAmericana in questo form da errore incrementando il giorno di 2
    
    strSql = "SELECT    * " & _
             "FROM      ACCESSI_VASCOLARI_TAB " & _
             "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
             "          DATA=#" & data & "#"
    
    Call Pulisci
    Set rsAccessi = New Recordset
    rsAccessi.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsAccessi.EOF And rsAccessi.BOF) Then
        txtIntervento = rsAccessi("INTERVENTO") & ""
        txtDati = rsAccessi("DATI") & ""
        oDataChiusuraAccesso.data = rsAccessi("DATA_CHIUSURA_ACCESSO") & ""
        txtCausaChiusuraAccesso = rsAccessi("CAUSA_CHIUSURA_ACCESSO") & ""
        intMedicoKey(0) = rsAccessi("CODICE_MEDICO1")
        Call CaricaMedico(0)
        intMedicoKey(1) = rsAccessi("CODICE_MEDICO2")
        Call CaricaMedico(1)
        If rsAccessi("ANESTESIA") <> 0 Then
            optAnestesia(rsAccessi("ANESTESIA") - 1).Value = True
        End If
        keyId = rsAccessi("KEY")
        Call Upd_rsDisco
        
        modifica = True
    Else
        modifica = False
    End If
    Set rsAccessi = Nothing
    
    blnModificato = False
End Sub

'' Restituire il valore numerico associato all'optAnestesia
' 1 - Locale
' 2 - Locoregionale
' 3 - Totale
Private Function GestisciOpt() As Integer
    If optAnestesia(0).Value = False And optAnestesia(1).Value = False And optAnestesia(2).Value = False Then
        GestisciOpt = 0
    Else
        If optAnestesia(0).Value = True Then
            GestisciOpt = 1
        ElseIf optAnestesia(1).Value = True Then
            GestisciOpt = 2
        ElseIf optAnestesia(2).Value = True Then
            GestisciOpt = 3
        End If
    End If
End Function

'' Determina se la scheda è completa prima del salvataggio
Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    Else
        If oData.data = "" Then
            MsgBox "Inserire una data", vbCritical, "Attenzione"
            Exit Function
        End If
    End If
    Completo = True
End Function

'' Pulisce il form tranne i dati sul paziente
Private Sub Pulisci()
    Dim i As Integer
    For i = 0 To 2
        optAnestesia(i).Value = False
    Next i
    txtDati = ""
    txtIntervento = ""
    oDataChiusuraAccesso.Pulisci
    txtCausaChiusuraAccesso = ""
    For i = 0 To 1
        lblCognomeMed(i) = ""
        lblNomeMed(i) = ""
    Next i
    intMedicoKey(0) = 0
    intMedicoKey(1) = 0
    blnModificato = False
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    modifica = False
    intPazientiKey = 0
    intMedicoKey(0) = 0
    intMedicoKey(1) = 0
    oData.Pulisci
    oDataChiusuraAccesso.Pulisci
    Call PulisciForm(Me)
    cmdTrova(0).SetFocus
    For i = 0 To 2
        optAnestesia(i).Value = False
    Next i
    blnModificato = False
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "La scheda deve essere prima memorizzata", vbCritical, "Attenzione"
        Exit Sub
    End If
      
    Set rsAccessi = New Recordset
    rsAccessi.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsAccessi("COGNOME") & " " & rsAccessi("NOME")
    structIntestazione.sDataPaziente = rsAccessi("DATA_NASCITA")
    Set rsAccessi = Nothing
    
    Call StampaNonaParte(False, intPazientiKey)
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    
    If Completo Then
        If Not modifica Then
            keyId = GetNumero("ACCESSI_VASCOLARI_TAB")
        End If
        v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "INTERVENTO", "DATA_CHIUSURA_ACCESSO", "CAUSA_CHIUSURA_ACCESSO", "CODICE_MEDICO1", _
                        "CODICE_MEDICO2", "ANESTESIA", "DATI")
        v_Val = Array(keyId, intPazientiKey, oData.data, txtIntervento, IIf(oDataChiusuraAccesso.data = "", Null, oDataChiusuraAccesso.data), txtCausaChiusuraAccesso, intMedicoKey(0), intMedicoKey(1), GestisciOpt, txtDati)
        Set rsAccessi = New Recordset
        If modifica Then
            rsAccessi.Open "SELECT * FROM ACCESSI_VASCOLARI_TAB WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsAccessi.Update v_Nomi, v_Val
            If TRACCIATO Then
                Call Confronta(rsAccessi)
            End If
        Else
            rsAccessi.Open "ACCESSI_VASCOLARI_TAB", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAccessi.AddNew v_Nomi, v_Val
            rsAccessi.Update
            Call Upd_rsDisco
        End If
        Set rsAccessi = Nothing
        
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        blnModificato = False
        modifica = True
    End If
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    If Index = 0 Then
        If ControlloChiusuraForm(blnModificato, Me.Caption) Then
            ' pulisce per evitare problemi
            Call PulisciTutto
            tTrova.Tipo = tpPAZIENTE
        Else
            Exit Sub
        End If
    Else
        tTrova.Tipo = tpMEDICOREFER
    End If
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If Index = 0 Then
        If tTrova.keyReturn = 0 Then
            Unload Me
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    Else
        intMedicoKey(Index - 1) = tTrova.keyReturn
        Call CaricaMedico(Index - 1)
    End If
End Sub

Private Sub cmdElimina_Click()
    Dim data As Date
            
    If intPazientiKey <> 0 Then
        If modifica Then
            If MsgBox("Sei sicuro di voler eliminare la scheda di: " & UCase(lblCognome) & " " & UCase(lblNome) & "?", vbQuestion & vbYesNo, "Eliminazione") = vbYes Then
                ' la data americana
                data = Month(oData.data) & "/" & Day(oData.data) & "/" & Year(oData.data)
                Set rsAccessi = New Recordset
                rsAccessi.Open "SELECT * FROM ACCESSI_VASCOLARI_TAB WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "#", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                If Not (rsAccessi.BOF And rsAccessi.EOF) Then
                    rsAccessi.Delete
                End If
                rsAccessi.Close
                                    
                Call PulisciTutto
                MsgBox "Eliminazione effettuata con successo", vbInformation, "Eliminazione"
                Call TrovaPaziente
            End If
        End If
    End If
End Sub

'' Carica i dati del medico per il form
' @param inIndice definisce quale dei due medici caricare
Private Sub CaricaMedico(inIndice As Integer)
    Dim rsDataset As Recordset
    If intMedicoKey(inIndice) = 0 Then Exit Sub
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME FROM MEDICI_REFERTANTI WHERE KEY=" & intMedicoKey(inIndice), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognomeMed(inIndice) = rsDataset("COGNOME") & ""
    lblNomeMed(inIndice) = rsDataset("NOME") & ""
    Set rsDataset = Nothing
    blnModificato = True
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then Exit Sub
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA,CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    blnModificato = False
End Sub

Private Sub oData_OnCalendarClick(blnProsegui As Boolean)
     blnProsegui = ControlloChiusuraForm(blnModificato, Me.Caption)
End Sub

'' Pulisce e carica le nuova scheda (se data non è null)
Private Sub oData_OnDataChange()
    If IsDate(oData.data) Then
        Call CaricaScheda
    ElseIf oData.data = "" Then
        Call Pulisci
    End If
End Sub

Private Sub oData_OnElencaClick()
        If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        tElenca.Tipo = tpACCESSO
        tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey
        frmElencaDate.Show 1
        If laData <> "" Then oData.data = laData
    End If
End Sub

Private Sub optAnestesia_KeyPress(Index As Integer, KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCausaChiusuraAccesso_GotFocus()
    txtCausaChiusuraAccesso.BackColor = colArancione
End Sub

Private Sub txtCausaChiusuraAccesso_LostFocus()
    txtCausaChiusuraAccesso.BackColor = vbWhite
End Sub

Private Sub txtDati_GotFocus()
    txtDati.BackColor = colArancione
End Sub

Private Sub txtDati_LostFocus()
    txtDati.BackColor = vbWhite
End Sub

Private Sub txtIntervento_GotFocus()
    txtIntervento.BackColor = colArancione
End Sub

Private Sub txtIntervento_LostFocus()
    txtIntervento.BackColor = vbWhite
End Sub

'******** Gestione Modificato

Private Sub txtIntervento_Change()
    blnModificato = True
End Sub

Private Sub txtDati_Change()
    blnModificato = True
End Sub

Private Sub txtCausaChiusuraAccesso_Change()
    blnModificato = True
End Sub

Private Sub optAnestesia_Click(Index As Integer)
    blnModificato = True
End Sub

Private Sub Upd_rsDisco()
' aggiorna i dati nel rsDisco
    Dim i As Integer
    Do While Not rsDisco.EOF
       rsDisco.Delete
       rsDisco.MoveNext
    Loop
        rsDisco.AddNew
        For i = 0 To rsAccessi.Fields.count - 1
            rsDisco.Fields(i) = rsAccessi.Fields(i)
        Next i
        rsDisco.Update
End Sub


