VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmAnamnesiNefro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ANAMNESI NEFROLOGICA"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmAnemnesiNefro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         Left            =   11400
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         Left            =   10680
         TabIndex        =   40
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   12375
      Begin VB.CheckBox chkTrattamento 
         Caption         =   "Trattamento Conservativo"
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
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chkIstologica 
         Caption         =   "Conferma Istologica"
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
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cboEDTA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   8775
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   7800
         TabIndex        =   48
         Top             =   660
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Diagnosi"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Malattia Renale Primitiva"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2616
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   1920
      Width           =   12375
      Begin VB.ComboBox cboCentro 
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
         Index           =   0
         Left            =   7920
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin DataTimeBox.uDataTimeBox oDataInizioInSede 
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   560
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataFine 
         Height          =   375
         Left            =   3600
         TabIndex        =   47
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   49
         Top             =   170
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Fine Emodialisi in Sede"
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
         TabIndex        =   35
         Top             =   1005
         Width           =   3030
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inizio Emodialisi"
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
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sede"
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
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inizio Emodialisi in Sede"
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
         Left            =   120
         TabIndex        =   32
         Top             =   630
         Width           =   3120
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   12375
      Begin VB.CheckBox chkSospensione 
         Caption         =   "Sosp.Temp."
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
         Left            =   10560
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtNote 
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
         Left            =   7920
         MaxLength       =   25
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtNote 
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
         Left            =   7920
         MaxLength       =   25
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkAttesaTrapianto 
         Caption         =   "In Attesa Trapianto"
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
         TabIndex        =   5
         Top             =   255
         Width           =   2415
      End
      Begin VB.ComboBox cboCentro 
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
         Index           =   1
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cboCentro 
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
         Index           =   2
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkSospensione 
         Caption         =   "Sosp.Temp."
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
         Left            =   10560
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   15
         Left            =   7200
         TabIndex        =   37
         Top             =   735
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   16
         Left            =   7200
         TabIndex        =   36
         Top             =   255
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prima Sede"
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
         Left            =   2640
         TabIndex        =   30
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seconda Sede"
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
         Left            =   2640
         TabIndex        =   29
         Top             =   765
         Width           =   1560
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   12375
      Begin VB.CheckBox chkPrecTrapianto 
         Caption         =   "Precedente Trapianto"
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
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboCentro 
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
         Index           =   3
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   2775
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   2
         Left            =   7800
         TabIndex        =   50
         Top             =   550
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
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
         Index           =   11
         Left            =   7200
         TabIndex        =   27
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sede"
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
         Index           =   13
         Left            =   2640
         TabIndex        =   26
         Top             =   645
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   12375
      Begin VB.CheckBox chkPrecEspianto 
         Caption         =   "Precedente Espianto"
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
         TabIndex        =   14
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox cboCentro 
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
         Index           =   4
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   2775
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   3
         Left            =   7800
         TabIndex        =   51
         Top             =   550
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
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
         Index           =   12
         Left            =   7200
         TabIndex        =   24
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sede"
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
         Index           =   14
         Left            =   2640
         TabIndex        =   23
         Top             =   645
         Width           =   570
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   12375
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
         Left            =   7440
         TabIndex        =   16
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
         Left            =   10920
         TabIndex        =   18
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
         Left            =   9120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAnamnesiNefro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAnemnesiNefro.frm
'
' <b>Descrizione</b>: Scheda Anamnesi Nefrologiche associata alla tab ANAMNESI_NEFROLOGICHE
'
' @remarks
'
' @author
'
' @date 01/02/2011 21.17
Option Explicit
'' rs della scheda
Dim rsAnamnesiNefro As Recordset
'' indica se si è in fase di modifica
Dim modifica As Boolean
'' key del record aperto
Dim keyId As Integer
'' rs per la tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim blnModificato As Boolean
    
'' Ricarica le cbo e apre il form Trova se non c'è nessun paziente caricato
Private Sub Form_Activate()
    Dim i As Integer
    Dim blnModificatoAppo As Boolean
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    blnModificatoAppo = blnModificato
    Call RicaricaComboBox("EDTA", "NOME", cboEDTA)
    For i = 0 To 4
        Call RicaricaComboBox("CENTRI_PROVENIENZA", "NOME", cboCentro(i))
    Next i
    blnModificato = blnModificatoAppo
    
    Select Case CaricaPazienteInAperturaForm(Me.Caption, blnModificato, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    modifica = False
    For i = 0 To 3
        oData(i).ConnectionString = strConnectionStringCentro
    Next i
    oDataFine.ConnectionString = strConnectionStringCentro
    oDataInizioInSede.ConnectionString = strConnectionStringCentro
    Call ApriRsDisconnesso
    blnModificato = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
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
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_NEFROLOGICHE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Determina se la scheda è completa prima del salvataggio
Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If oDataInizioInSede.data = "" Then
        MsgBox "Inserire la DATA di INIZIO EMODIALISI in SEDE", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    modifica = False
    For i = 0 To 3
        oData(i).Pulisci
    Next i
    oDataFine.Pulisci
    oDataInizioInSede.Pulisci
    intPazientiKey = 0
    chkAttesaTrapianto.Value = Unchecked
    chkIstologica.Value = Unchecked
    chkPrecEspianto.Value = Unchecked
    chkPrecTrapianto.Value = Unchecked
    chkTrattamento.Value = Unchecked
    chkSospensione(0).Value = Unchecked
    chkSospensione(1).Value = Unchecked
    Call PulisciForm(Me)
    cmdTrova.SetFocus
    blnModificato = False
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "La scheda deve essere prima memorizzata", vbCritical, "Attenzione"
        Exit Sub
    End If
      
    Set rsAnamnesiNefro = New Recordset
    rsAnamnesiNefro.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsAnamnesiNefro("COGNOME") & " " & rsAnamnesiNefro("NOME")
    structIntestazione.sDataPaziente = rsAnamnesiNefro("DATA_NASCITA")
    Set rsAnamnesiNefro = Nothing

    Call StampaTerzaParte(False, intPazientiKey)
End Sub

Private Sub cmdMemorizza_Click()
    Dim i As Integer
    Dim v_Nomi(1 To 23) As Variant
    Dim v_Val(1 To 23) As Variant
    
    If Completo Then
        ' gestisce il centro solo durante la memorizzazione
        For i = 0 To 4
            If cboCentro(i).Text <> "" Then
                Call GestisciNuovo("CENTRI_PROVENIENZA", cboCentro(i))
            End If
        Next i
        
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "CODICE_PAZIENTE"
        v_Nomi(3) = "CODICE_EDTA"
        v_Nomi(4) = "ISTOLOGICA"
        v_Nomi(5) = "TRATTAMENTO_CONS"
        v_Nomi(6) = "ATTESA_TRAPIANTO"
        v_Nomi(7) = "PREC_TRAPIANTO"
        v_Nomi(8) = "PREC_ESPIANTO"
        For i = 1 To 5
            v_Nomi(8 + i) = "SEDE" & i - 1
        Next i
        v_Nomi(14) = "NOTE1"
        v_Nomi(15) = "NOTE2"
        v_Nomi(16) = "SOSPENSIONE1"
        v_Nomi(17) = "SOSPENSIONE2"
        For i = 0 To 3
            v_Nomi(18 + i) = "DATA" & i
        Next i
        v_Nomi(22) = "DATA_INIZIO"
        v_Nomi(23) = "DATA_FINE"
        
        If Not modifica Then
            keyId = GetNumero("ANAMNESI_NEFROLOGICHE")
        End If
        v_Val(1) = keyId
        v_Val(2) = intPazientiKey
        If cboEDTA.ListIndex = -1 Then
            v_Val(3) = -1
        Else
            v_Val(3) = cboEDTA.ItemData(cboEDTA.ListIndex)
        End If
        v_Val(4) = IIf(chkIstologica.Value = Checked, True, False)
        v_Val(5) = IIf(chkTrattamento.Value = Checked, True, False)
        v_Val(6) = IIf(chkAttesaTrapianto.Value = Checked, True, False)
        v_Val(7) = IIf(chkPrecTrapianto.Value = Checked, True, False)
        v_Val(8) = IIf(chkPrecEspianto.Value = Checked, True, False)
        For i = 1 To 5
            If cboCentro(i - 1).ListIndex = -1 Then
                v_Val(8 + i) = -1
            Else
                v_Val(8 + i) = cboCentro(i - 1).ItemData(cboCentro(i - 1).ListIndex)
            End If
        Next i
        v_Val(14) = txtNote(0) & ""
        v_Val(15) = txtNote(1) & ""
        v_Val(16) = IIf(chkSospensione(0).Value = Checked, True, False)
        v_Val(17) = IIf(chkSospensione(1).Value = Checked, True, False)
        For i = 0 To 3
            v_Val(18 + i) = IIf(oData(i).data = "", Null, oData(i).data)
        Next i
        v_Val(22) = IIf(oDataInizioInSede.data = "", Null, oDataInizioInSede.data)
        v_Val(23) = IIf(oDataFine.data = "", Null, oDataFine.data)
        
        Set rsAnamnesiNefro = New Recordset
        If modifica Then
            rsAnamnesiNefro.Open "SELECT * FROM ANAMNESI_NEFROLOGICHE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsAnamnesiNefro.Update v_Nomi, v_Val
            If TRACCIATO Then
                Call Confronta(rsAnamnesiNefro)
            End If
        Else
            rsAnamnesiNefro.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAnamnesiNefro.AddNew v_Nomi, v_Val
            rsAnamnesiNefro.Update
            Call Upd_rsDisco
        End If
        Set rsAnamnesiNefro = Nothing
        
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        blnModificato = False
        modifica = True
    End If
End Sub

Private Sub cmdTrova_Click()
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
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
    End If
End Sub

'' Evita di far inserire testo > 30 caratteri
Private Sub cboCentro_KeyPress(Index As Integer, KeyAscii As Integer)
    If Len(cboCentro(Index).Text) >= 30 Then
        Beep
        KeyAscii = 0
    End If
End Sub

'' Carica i dati del paziente e carica i dati della scheda dalla tabella nel form e nel rs per la tracciatura
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    Dim i As Integer
    
    If intPazientiKey = 0 Then Exit Sub
        
    ' carica i dati del paziente
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
    
    ' cerca i riferimenti al paziente
    Set rsAnamnesiNefro = New Recordset
    rsAnamnesiNefro.Open "SELECT * FROM ANAMNESI_NEFROLOGICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsAnamnesiNefro.BOF And rsAnamnesiNefro.EOF Then
        ' il paziente non ha una scheda clinica
        modifica = False
    Else
        keyId = rsAnamnesiNefro("KEY")
        modifica = True
        cboEDTA.ListIndex = GetCboListIndex(rsAnamnesiNefro("CODICE_EDTA"), cboEDTA)
        txtNote(0) = rsAnamnesiNefro("NOTE1") & ""
        txtNote(1) = rsAnamnesiNefro("NOTE2") & ""
        chkAttesaTrapianto.Value = IIf(CBool(rsAnamnesiNefro("ATTESA_TRAPIANTO")) = True, Checked, Unchecked)
        chkIstologica.Value = IIf(CBool(rsAnamnesiNefro("ISTOLOGICA")) = True, Checked, Unchecked)
        chkPrecEspianto.Value = IIf(CBool(rsAnamnesiNefro("PREC_ESPIANTO")) = True, Checked, Unchecked)
        chkPrecTrapianto.Value = IIf(CBool(rsAnamnesiNefro("PREC_TRAPIANTO")) = True, Checked, Unchecked)
        chkTrattamento.Value = IIf(CBool(rsAnamnesiNefro("TRATTAMENTO_CONS")) = True, Checked, Unchecked)
        chkSospensione(0).Value = IIf(CBool(rsAnamnesiNefro("SOSPENSIONE1")), Checked, Unchecked)
        chkSospensione(1).Value = IIf(CBool(rsAnamnesiNefro("SOSPENSIONE2")), Checked, Unchecked)
        For i = 0 To 3
            If rsAnamnesiNefro("DATA" & i) <> "" Then
                oData(i).data = rsAnamnesiNefro("DATA" & i)
            End If
        Next i
        For i = 0 To 4
            cboCentro(i).ListIndex = GetCboListIndex(rsAnamnesiNefro("SEDE" & i), cboCentro(i))
        Next i
        If rsAnamnesiNefro("DATA_INIZIO") <> "" Then
            oDataInizioInSede.data = rsAnamnesiNefro("DATA_INIZIO")
        End If
        If rsAnamnesiNefro("DATA_FINE") <> "" Then
            oDataFine.data = rsAnamnesiNefro("DATA_FINE")
        End If
        
        Call Upd_rsDisco
        
    End If
    Set rsAnamnesiNefro = Nothing
    blnModificato = False
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    blnModificato = True
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

Private Sub oDataFine_OnDataClick()
    oDataFine.Pulisci
End Sub

Private Sub oDataFine_OnDataChange()
    blnModificato = True
End Sub

Private Sub oDataInizioInSede_OnDataChange()
    blnModificato = True
End Sub

Private Sub oDataInizioInSede_OnDataClick()
    oDataInizioInSede.Pulisci
End Sub

Private Sub txtNote_GotFocus(Index As Integer)
    txtNote(Index).BackColor = colArancione
End Sub

Private Sub txtNote_LostFocus(Index As Integer)
    txtNote(Index).BackColor = vbWhite
End Sub




'******** Gestione Modificato

Private Sub txtNote_Change(Index As Integer)
    blnModificato = True
End Sub
    
Private Sub chkAttesaTrapianto_Click()
    blnModificato = True
End Sub

Private Sub chkIstologica_Click()
    blnModificato = True
End Sub

Private Sub chkPrecEspianto_Click()
    blnModificato = True
End Sub

Private Sub chkPrecTrapianto_Click()
    blnModificato = True
End Sub

Private Sub chkSospensione_Click(Index As Integer)
    blnModificato = True
End Sub

Private Sub chkTrattamento_Click()
    blnModificato = True
End Sub

Private Sub cboCentro_Click(Index As Integer)
    blnModificato = True
End Sub

Private Sub cboEDTA_Click()
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
       For i = 0 To rsAnamnesiNefro.Fields.count - 1
           rsDisco.Fields(i) = rsAnamnesiNefro.Fields(i)
       Next i
       rsDisco.Update
End Sub
