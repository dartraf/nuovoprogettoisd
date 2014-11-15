VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmEsamiStrumentali 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ESAMI STRUMENTALI"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmEsamiStrumentali.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   12015
      Begin VB.CheckBox chkFiltra 
         Height          =   270
         Index           =   1
         Left            =   2400
         Picture         =   "frmEsamiStrumentali.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Filtra esami effettuati"
         Top             =   680
         Width           =   375
      End
      Begin VB.CheckBox chkFiltra 
         Height          =   270
         Index           =   0
         Left            =   2400
         Picture         =   "frmEsamiStrumentali.frx":05A3
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Filtra esami effettuati"
         Top             =   240
         Width           =   375
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
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   680
         Width           =   5655
      End
      Begin VB.ComboBox cboOrgano 
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
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Left            =   9480
         TabIndex        =   8
         Top             =   600
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
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
         Left            =   8880
         TabIndex        =   30
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di Esame"
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
         Left            =   360
         TabIndex        =   29
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Organo/Apparato"
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
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   1680
         Picture         =   "frmEsamiStrumentali.frx":06ED
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Seleziona il medico"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtReferto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1320
         Width           =   11655
      End
      Begin VB.CheckBox chkStampa 
         Caption         =   "Stampa su Cartella Clinica"
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
         Left            =   7800
         TabIndex        =   12
         Top             =   840
         Value           =   1  'Checked
         Width           =   3975
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
         Left            =   3600
         TabIndex        =   10
         Top             =   360
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
         Left            =   8640
         TabIndex        =   11
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrizione Referto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   2175
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
         Left            =   2400
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
         Index           =   6
         Left            =   7800
         TabIndex        =   24
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Medico Refer."
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   5160
      Width           =   12015
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
         Index           =   41
         Left            =   7560
         TabIndex        =   38
         Top             =   380
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
         Index           =   40
         Left            =   2400
         TabIndex        =   37
         Top             =   380
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultima modifica effettuata da"
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
         Index           =   39
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCognomeUtente 
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
         Left            =   3600
         TabIndex        =   14
         Top             =   350
         Width           =   3375
      End
      Begin VB.Label lblNomeUtente 
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
         Left            =   8400
         TabIndex        =   15
         Top             =   350
         Width           =   3375
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   5880
      Width           =   12015
      Begin VB.OptionButton OptStEsame 
         Caption         =   "Stampa &VISTA CORRENTE"
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
         Left            =   4220
         TabIndex        =   41
         Top             =   210
         Value           =   -1  'True
         Width           =   3160
      End
      Begin VB.OptionButton OptStEsame 
         Caption         =   "Stampa &TUTTO"
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
         Left            =   4220
         TabIndex        =   40
         Top             =   530
         Width           =   3160
      End
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "S&tampa"
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
         Left            =   7440
         TabIndex        =   17
         Top             =   240
         Width           =   960
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
         Height          =   495
         Left            =   8490
         TabIndex        =   18
         Top             =   240
         Width           =   975
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
         Left            =   10920
         TabIndex        =   21
         Top             =   240
         Width           =   975
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
         Left            =   9520
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblTesto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2355
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog cdlApri 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmEsamiStrumentali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmEsamiStrumentali.frm
'
' <b>Descrizione</b>: Scheda Esami Strumentali associata alla tab ESAMI_STRUMENTALI
'
' @remarks
'
' @author
'
' @date 05/02/2011 18.23
Option Explicit

'' rs della scheda
Dim rsEsami As Recordset
'' indica se si è in fase di modifica
Dim modifica As Boolean
'' evita l'evento click nella cbo quando si sta pulendo la scheda
Dim stoPulendo As Boolean
'' key del record caricato
Dim keyId As Integer
'' rs per la tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim intMedicoKey As Integer
Dim blnModificato As Boolean

'' Ricarica le cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("ORGANI", "NOME", cboOrgano)
    
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
    
    stoPulendo = False
    oData.ConnectionString = strConnectionStringCentro
    Call ApriRsDisconnesso
    Call EliminaScansioniSospese("SCAN_ESAMI_STRUMENTALI")
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click (0)
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        intMedicoKey = 0
        Call EliminaScansioniSospese("SCAN_ESAMI_STRUMENTALI")
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ESAMI_STRUMENTALI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
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
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "DATA_RECORD", "CODICE_ORGANO", "CODICE_ESAME", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), oData.data, cboOrgano.ItemData(cboOrgano.ListIndex), cboEsami.ItemData(cboEsami.ListIndex), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_ESAMI_STRUMENTALI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Carica i dati dell'ultimo utente modificatore
Private Sub CaricaUtenteModificatore(key As Integer)
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT COGNOME, NOME FROM LOGIN WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        lblCognomeUtente = rsDataset("COGNOME")
        lblNomeUtente = rsDataset("NOME")
    Else
        lblCognomeUtente = ""
        lblNomeUtente = ""
    End If
    rsDataset.Close
End Sub

'' Verifica prima di memorizzare che tutti i dati siano inseriti
Private Function Completo() As Boolean
    Completo = False
    If cboOrgano.ListIndex = -1 Then
        MsgBox "Selezionare l'organo", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboEsami.ListIndex = -1 Then
        MsgBox "Selezionare l'esame", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData.data = "" Then
        MsgBox "Inserire la data dell'esame", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

'' Pulisce la scheda quando si selezionano o organo o esame
'
' @param pureEsame se true pulisce anche la cboEsami e la data
Private Sub Pulisci(pureEsame As Boolean)
    stoPulendo = True
    If pureEsame Then
        cboEsami.Clear
        oData.Pulisci
    End If
    intMedicoKey = 0
    keyId = 0
    lblCognomeMed = ""
    lblNomeMed = ""
    txtReferto = ""
    lblTesto(0) = ""
    stoPulendo = False
    Call EliminaScansioniSospese("SCAN_ESAMI_STRUMENTALI")
    blnModificato = False
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    stoPulendo = True
    modifica = False
    intPazientiKey = 0
    intMedicoKey = 0
    oData.Pulisci
    lblTesto(0) = ""
    keyId = 0
    chkStampa.Value = Checked
    chkFiltra(0).Value = Unchecked
    chkFiltra(1).Value = Unchecked
    Call PulisciForm(Me)
    stoPulendo = False
    cmdTrova(0).SetFocus
    lblCognomeUtente.Caption = ""
    lblNomeUtente.Caption = ""
    Call RicaricaComboBox("ORGANI", "NOME", cboOrgano)
    blnModificato = False
End Sub

'' Aggiorna le info su eventuali referti
Private Sub AggiornaReferto()
    Dim rsDataset As Recordset
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM SCAN_ESAMI_STRUMENTALI WHERE CODICE_SCHEDA=" & keyId, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        lblTesto(0) = "E' presente un referto di " & rsDataset.RecordCount & " pag."
    Else
        lblTesto(0) = "NON sono presenti referti"
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

'' Carica i dati della scheda dalla tabella nel form e nel rs per la tracciatura
Private Sub CaricaScheda()
    Dim data As Date
    Dim i As Integer
    If oData.data = "" Then Exit Sub
    If cboEsami.ListIndex = -1 Then Exit Sub
    
    ' la data americana
    data = Month(oData.data) & "/" & Day(oData.data) & "/" & Year(oData.data)
    
    'data = oData.DataAmericana in questa sub da errore incrementando il giorno di 2
    
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT * FROM ESAMI_STRUMENTALI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & " AND CODICE_ESAME=" & cboEsami.ItemData(cboEsami.ListIndex) & " AND DATA=#" & data & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsEsami.EOF And rsEsami.BOF Then
        modifica = False
    Else
        keyId = rsEsami("KEY")
        modifica = True
        intMedicoKey = rsEsami("CODICE_MEDICO")
        Call CaricaMedico
        chkStampa.Value = IIf(CBool(rsEsami("STAMPA")) = True, Checked, Unchecked)
        txtReferto = rsEsami("REFERTO")
        Call CaricaUtenteModificatore(rsEsami("UTENTE_MODIFICATORE"))
        Call AggiornaReferto
        Call Upd_rsDisco
        
    End If
    Set rsEsami = Nothing
    
    blnModificato = False
End Sub

'' Salva l'eliminazione nel db di tracciature
Private Sub SalvaEliminazione()
    Dim v_nome As Variant
    Dim v_Val As Variant
    Dim rsDataset As New Recordset
    
    v_nome = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "DATA_ESAME", "CODICE_ORGANO", "CODICE_ESAME")
    v_Val = Array(tAccesso.key, date, Time, intPazientiKey, oData.data, cboOrgano.ItemData(cboOrgano.ListIndex), cboEsami.ItemData(cboEsami.ListIndex))
    rsDataset.Open "E_ESAMI_STRUMENTALI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_nome, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

'' Salva l'eliminazione del referto nel db di tracciature
Public Sub SalvaEliminazioneReferto(nomeFile)
    Dim v_nome As Variant
    Dim v_Val As Variant
    Dim rsDataset As New Recordset
    
    v_nome = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "DATA_ESAME", "CODICE_ORGANO", "CODICE_ESAME", "NOME_FILE")
    v_Val = Array(tAccesso.key, date, Time, intPazientiKey, oData.data, cboOrgano.ItemData(cboOrgano.ListIndex), cboEsami.ItemData(cboEsami.ListIndex), nomeFile)
    rsDataset.Open "E_SCAN_ESAMI_STRUMENTALI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_nome, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

Private Sub chkFiltra_Click(Index As Integer)
    Dim strSql As String
    If intPazientiKey <> 0 Then
        Call Pulisci(True)
        If Index = 0 Then
            If chkFiltra(0).Value = Checked Then
                strSql = "SELECT     DISTINCT ORGANI.NOME, ORGANI.KEY " & _
                        "FROM       (ESAMI_STRUMENTALI " & _
                        "           INNER JOIN ORGANI ON ORGANI.KEY=ESAMI_STRUMENTALI.CODICE_ORGANO) " & _
                        "WHERE      CODICE_PAZIENTE=" & intPazientiKey
                Call RicaricaComboBox(strSql, "NOME", cboOrgano)
            Else
                Call RicaricaComboBox("ORGANI", "NOME", cboOrgano)
            End If
        Else
            If cboOrgano.ListIndex <> -1 Then
                If chkFiltra(1).Value = Checked Then
                    strSql = "SELECT    DISTINCT ESAMI.NOME, ESAMI.KEY " & _
                            "FROM       (ESAMI_STRUMENTALI " & _
                            "           INNER JOIN ESAMI ON ESAMI.KEY=ESAMI_STRUMENTALI.CODICE_ESAME) " & _
                            "WHERE      (CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                            "           ESAMI_STRUMENTALI.CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & ")"
                    Call RicaricaComboBox(strSql, "NOME", cboEsami)
                Else
                    Call RicaricaComboBox("SELECT DISTINCT NOME, KEY FROM ESAMI WHERE (CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & ")", "NOME", cboEsami)
                End If
            End If
        End If
    End If
End Sub

Private Sub cboEsami_Click()
    If stoPulendo Then Exit Sub
    ' puo elencare solo se il esame è stato selezionato
    If cboEsami.ListIndex = -1 Then
        oData.EnableElenca (False)
    Else
        oData.EnableElenca (True)
        Call Pulisci(False)
        Call CaricaScheda
    End If
End Sub

Private Sub cboOrgano_Click()
    Dim strSql As String
    
    If stoPulendo Then Exit Sub
    Call Pulisci(True)
    ' carica la lista degli esami per quell'organo
    If chkFiltra(1).Value = Checked Then
        strSql = "SELECT    DISTINCT ESAMI.NOME, ESAMI.KEY " & _
                "FROM       (ESAMI_STRUMENTALI " & _
                "           INNER JOIN ESAMI ON ESAMI.KEY=ESAMI_STRUMENTALI.CODICE_ESAME) " & _
                "WHERE      (CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                "           ESAMI_STRUMENTALI.CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & ")"
        Call RicaricaComboBox(strSql, "NOME", cboEsami)
    Else
        strSql = "SELECT    DISTINCT NOME, KEY " & _
                "FROM       ESAMI " & _
                "WHERE      (CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & ")"
        Call RicaricaComboBox(strSql, "NOME", cboEsami)
    End If
End Sub

Private Sub cmdGestioneReferti_Click()
    Unload frmGestioneDocumentiEsterni
    Load frmGestioneDocumentiEsterni
    frmGestioneDocumentiEsterni.LetCodicePaziente = intPazientiKey
    If modifica Then
        frmGestioneDocumentiEsterni.letcodiceRecord = keyId
        frmGestioneDocumentiEsterni.LetNomeFile = E_ST & keyId & " " & Replace(date, "/", "-")
    Else
        frmGestioneDocumentiEsterni.letcodiceRecord = 0
        frmGestioneDocumentiEsterni.LetNomeFile = E_ST & 0 & " " & Replace(date, "/", "-")
    End If
    tDocumentiEsterni = tpSCANESAMISTRUMENTALI
    frmGestioneDocumentiEsterni.Show 1
    Call AggiornaReferto
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "Impossibile stampare", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If OptStEsame(1).Value Then  ' stampa tutto
        Set rsEsami = New Recordset
        rsEsami.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        structIntestazione.sPaziente = rsEsami("COGNOME") & " " & rsEsami("NOME")
        structIntestazione.sDataPaziente = rsEsami("DATA_NASCITA")
        Set rsEsami = Nothing
        Call StampaQuintaParte(False, intPazientiKey)
        Exit Sub
    End If
    
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsEsami("COGNOME") & " " & rsEsami("NOME")
    structIntestazione.sDataPaziente = rsEsami("DATA_NASCITA")
    Set rsEsami = Nothing

    Dim strShape As String
    Dim strSql As String
    Dim codicePaziente As Integer
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    codicePaziente = intPazientiKey
    
    strShape = "SHAPE APPEND " & _
                "   NEW adVarChar(50) AS NOME_ORGANO, " & _
                "   NEW adVarChar(50) AS NOME_ESAME, " & _
                "   NEW adDate AS DATA, " & _
                "   NEW adLongVarChar AS REFERTO "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
        
    ' carica il recordset padre
     strSql = "SELECT    ORGANI.NOME AS ORGANINOME, ESAMI.NOME AS ESAMINOME, DATA, REFERTO, UTENTE_MODIFICATORE " & _
             "FROM      ((ESAMI_STRUMENTALI INNER JOIN ORGANI ON ORGANI.KEY=ESAMI_STRUMENTALI.CODICE_ORGANO) " & _
             "          INNER JOIN ESAMI ON ESAMI.KEY=ESAMI_STRUMENTALI.CODICE_ESAME) " & _
             "WHERE     ESAMI_STRUMENTALI.KEY=" & keyId & " AND STAMPA=TRUE "
 
 '   strSql = "SELECT    ORGANI.NOME AS ORGANINOME, ESAMI.NOME AS ESAMINOME, DATA, REFERTO, UTENTE_MODIFICATORE " & _
             "FROM      ((ESAMI_STRUMENTALI INNER JOIN ORGANI ON ORGANI.KEY=ESAMI_STRUMENTALI.CODICE_ORGANO) " & _
             "          INNER JOIN ESAMI ON ESAMI.KEY=ESAMI_STRUMENTALI.CODICE_ESAME) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          STAMPA=TRUE AND DATA=#" & oData.DataAmericana & "# "
    
    Set rsEsami = New Recordset
    rsEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsEsami.BOF Or rsEsami.EOF) Then
        With rsMain
            .AddNew
            .Fields("NOME_ORGANO") = rsEsami("ORGANINOME")
            .Fields("NOME_ESAME") = rsEsami("ESAMINOME")
            .Fields("DATA") = rsEsami("DATA")
            .Fields("REFERTO") = rsEsami("REFERTO") & vbCrLf & vbCrLf & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsEsami("UTENTE_MODIFICATORE"))
        End With

        Set rptCartellaClinica_5 = Nothing
        Set rptCartellaClinica_5.DataSource = rsMain
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        rptCartellaClinica_5.PrintReport
    End If
    rsEsami.Close
    rsMain.Close
End Sub

Private Sub cmdElimina_Click()
    Dim data As Date
    Dim eliminato As Boolean
    Dim scansione As Boolean
    Dim nomeFile As String
    
    If intPazientiKey <> 0 Then
        If modifica Then
            If MsgBox("Sei sicuro di voler eliminare l'esame strumentale di: " & UCase(lblCognome) & " " & UCase(lblNome) & "?", vbQuestion & vbYesNo, "Eliminazione") = vbYes Then
                ' la data americana
                data = oData.DataAmericana
                Set rsEsami = New Recordset
                rsEsami.Open "SELECT * FROM ESAMI_STRUMENTALI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & " AND CODICE_ESAME=" & cboEsami.ItemData(cboEsami.ListIndex) & " AND DATA=#" & data & "#", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                If Not (rsEsami.BOF And rsEsami.EOF) Then
                    eliminato = True
                    rsEsami.Delete
                End If
                rsEsami.Close
                
                scansione = False
                rsEsami.Open "SELECT * FROM SCAN_ESAMI_STRUMENTALI WHERE CODICE_SCHEDA=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                Do While Not rsEsami.EOF
                    scansione = True
                    nomeFile = rsEsami("NOME_FILE")
                    rsEsami.Delete
                    If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                        Kill structApri.pathDB & "\" & nomeFile & ".jpg"
                    ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                        Kill structApri.pathDB & "\" & nomeFile & ".pdf"
                    End If
                    rsEsami.MoveNext
                Loop
                rsEsami.Close
            End If
        Else
            MsgBox "E' necessario memorizzare prima la scheda", vbInformation, "Informazione"
        End If
    End If
    If eliminato And TRACCIATO Then
        Call SalvaEliminazione
    End If
    If scansione And TRACCIATO Then
        Call SalvaEliminazioneReferto(nomeFile)
    End If
    Call PulisciTutto
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nomeFile As String
    
    If Completo Then
        If Not (modifica Or keyId <> 0) Then
            keyId = GetNumero("ESAMI_STRUMENTALI")
        End If
        v_Nomi = Array("KEY", "CODICE_PAZIENTE", "CODICE_ORGANO", "CODICE_ESAME", "CODICE_MEDICO", _
                     "STAMPA", "REFERTO", "DATA", "UTENTE_MODIFICATORE")
        v_Val = Array(keyId, intPazientiKey, cboOrgano.ItemData(cboOrgano.ListIndex), cboEsami.ItemData(cboEsami.ListIndex), _
                    intMedicoKey, IIf(chkStampa.Value = Checked, True, False), txtReferto, oData.data, tAccesso.key)
        
        Set rsEsami = New Recordset
        If modifica Then
            rsEsami.Open "SELECT * FROM ESAMI_STRUMENTALI WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsEsami.Update v_Nomi, v_Val
            If TRACCIATO Then
                Call Confronta(rsEsami)
            End If
        Else
            rsEsami.Open "ESAMI_STRUMENTALI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsEsami.AddNew v_Nomi, v_Val
            rsEsami.Update
            Call Upd_rsDisco
            rsEsami.Close
            ' controlla eventuali scansioni memorizzate in sospeso
            rsEsami.Open "SELECT * FROM SCAN_ESAMI_STRUMENTALI WHERE CODICE_SCHEDA=0", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            Do While Not rsEsami.EOF
                rsEsami("CODICE_SCHEDA") = keyId
                nomeFile = rsEsami("NOME_FILE")
                rsEsami("NOME_FILE") = E_ST & keyId & " " & Replace(date, "/", "-") & Right(nomeFile, 2)
                rsEsami.Update
                rsEsami.MoveNext
                If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".jpg" As structApri.pathDB & "\" & E_ST & keyId & " " & Replace(date, "/", "-") & Right(nomeFile, 2) & ".jpg"
                ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".pdf" As structApri.pathDB & "\" & E_ST & keyId & " " & Replace(date, "/", "-") & Right(nomeFile, 2) & ".pdf"
                End If
            Loop
            rsEsami.Close
        End If
        Set rsEsami = Nothing
        
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
        intMedicoKey = tTrova.keyReturn
        Call CaricaMedico
    End If
End Sub

'' Carica i dati del medico
Private Sub CaricaMedico()
    Dim rsDataset As Recordset
    If intMedicoKey = 0 Then Exit Sub
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME FROM MEDICI_REFERTANTI WHERE KEY=" & intMedicoKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognomeMed = rsDataset("COGNOME")
    lblNomeMed = "" & rsDataset("NOME")
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
    ' cerca i riferimenti al paziente solo dopo aver scelto anche l'esame e la data
    
    Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
    blnModificato = False
End Sub

Private Sub oData_OnCalendarClick(blnProsegui As Boolean)
    If cboOrgano.ListIndex = -1 Then
        MsgBox "Selezionare l' Organo/Apparato", vbInformation, "Informazione"
        blnProsegui = False
        Exit Sub
    ElseIf cboEsami.ListIndex = -1 Then
        MsgBox "Selezionare il Tipo di Esame", vbInformation, "Informazione"
        blnProsegui = False
        Exit Sub
    End If
    
    blnProsegui = ControlloChiusuraForm(blnModificato, Me.Caption)
End Sub

Private Sub oData_OnDataChange()
    If stoPulendo Then Exit Sub
    Call Pulisci(False)
    If oData.data <> "" Then
        Call CaricaScheda
        Frame3.Enabled = True
    End If
End Sub

Private Sub oData_OnDataClick()
    If cboOrgano.ListIndex = -1 Then
        MsgBox "Selezionare l' Organo/Apparato", vbInformation, "Informazione"
        Exit Sub
    ElseIf cboEsami.ListIndex = -1 Then
        MsgBox "Selezionare il Tipo di Esame", vbInformation, "Informazione"
        oData.Pulisci
        Exit Sub
    ElseIf ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oData.Pulisci
    End If
End Sub

Private Sub oData_OnElencaClick()
    If cboOrgano.ListIndex = -1 Then
        MsgBox "Selezionare l' Organo/Apparato", vbInformation, "Informazione"
        Exit Sub
    ElseIf cboEsami.ListIndex = -1 Then
        MsgBox "Selezionare il Tipo di Esame", vbInformation, "Informazione"
        Exit Sub
    ElseIf ControlloChiusuraForm(blnModificato, Me.Caption) Then
        ' setta le variabili che saranno viste dal frmElencaDate
        tElenca.Tipo = tpESAMISTRUMENTALI
        tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_ORGANO=" & cboOrgano.ItemData(cboOrgano.ListIndex) & " AND CODICE_ESAME=" & cboEsami.ItemData(cboEsami.ListIndex)
        frmElencaDate.Show 1
        If laData <> "" Then oData.data = laData
    Else
        Frame3.Enabled = True
    End If
End Sub

Private Sub txtReferto_GotFocus()
 ' se la data non è presente NON abilita a scrivere nel frame3
   If oData.data = "" Then
       Frame3.Enabled = False
       Exit Sub
   Else
       txtReferto.BackColor = colArancione
   End If
End Sub

Private Sub txtReferto_LostFocus()
    txtReferto.BackColor = vbWhite
End Sub

'******** Gestione Modificato

Private Sub txtReferto_Change()
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
      For i = 0 To rsEsami.Fields.count - 1
          rsDisco.Fields(i) = rsEsami.Fields(i)
      Next i
      rsDisco.Update
End Sub

