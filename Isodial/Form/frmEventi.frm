VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEventi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Eventi"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   12615
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmEventi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Seleziona il paziente"
         Top             =   240
         Width           =   450
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
         Left            =   11640
         TabIndex        =   35
         Top             =   360
         Width           =   615
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
         Left            =   7080
         TabIndex        =   34
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
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   3255
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
         Left            =   10920
         TabIndex        =   32
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
         Left            =   6240
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   360
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab tabSchede 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ricoveri"
      TabPicture(0)   =   "frmEventi.frx":0459
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAzioni(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Episodi Edema Polmonare "
      TabPicture(1)   =   "frmEventi.frx":0475
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAzioni(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdStampa(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Trasfusioni"
      TabPicture(2)   =   "frmEventi.frx":0491
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAzioni(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdStampa(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Sieroconversioni"
      TabPicture(3)   =   "frmEventi.frx":04AD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraAzioni(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdStampa(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
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
         Height          =   735
         Index           =   3
         Left            =   -69720
         TabIndex        =   39
         Top             =   2040
         Width           =   1455
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
         Height          =   735
         Index           =   2
         Left            =   -68400
         TabIndex        =   38
         Top             =   2040
         Width           =   1455
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
         Height          =   735
         Index           =   1
         Left            =   -68400
         TabIndex        =   37
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Frame Frame5 
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   5055
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   2295
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            MousePointer    =   99
            FormatString    =   "| Data                      | Sieroc. HBV   | Sieroc. HCV   "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmEventi.frx":04C9
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
         Height          =   2895
         Index           =   3
         Left            =   -69840
         TabIndex        =   20
         Top             =   480
         Width           =   3255
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
            Height          =   735
            Index           =   3
            Left            =   1680
            TabIndex        =   23
            Top             =   1560
            Width           =   1455
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
            Height          =   735
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1455
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
            Height          =   735
            Index           =   3
            Left            =   1680
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   6375
         Begin VB.ComboBox cboTrasfusioni 
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
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   2295
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            MousePointer    =   99
            FormatString    =   "| Data                      | Note                                                                             "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmEventi.frx":0623
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   2895
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   6375
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
            Index           =   1
            Left            =   1920
            MaxLength       =   35
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   3960
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   2295
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            MousePointer    =   99
            FormatString    =   "| Data                      | Note                                                                             "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmEventi.frx":077D
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9135
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
            Index           =   0
            Left            =   360
            MaxLength       =   35
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   3840
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   2295
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            MousePointer    =   15
            FormatString    =   $"frmEventi.frx":08D7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmEventi.frx":0966
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
         Height          =   2895
         Index           =   0
         Left            =   9120
         TabIndex        =   7
         Top             =   480
         Width           =   3375
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
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   1560
            Width           =   1455
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
            Height          =   735
            Index           =   0
            Left            =   1800
            TabIndex        =   10
            Top             =   1560
            Width           =   1455
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
            Height          =   735
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   1455
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
            Height          =   735
            Index           =   0
            Left            =   1800
            TabIndex        =   8
            Top             =   600
            Width           =   1455
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
         Height          =   2895
         Index           =   1
         Left            =   -68760
         TabIndex        =   11
         Top             =   480
         Width           =   3495
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
            Height          =   735
            Index           =   1
            Left            =   1920
            TabIndex        =   14
            Top             =   600
            Width           =   1455
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
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   600
            Width           =   1455
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
            Height          =   735
            Index           =   1
            Left            =   1920
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
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
         Height          =   2895
         Index           =   2
         Left            =   -68760
         TabIndex        =   15
         Top             =   480
         Width           =   3495
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
            Height          =   735
            Index           =   2
            Left            =   1920
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
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
            Height          =   735
            Index           =   2
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   1455
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
            Height          =   735
            Index           =   2
            Left            =   1920
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmEventi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmEventi.frm
'
' <b>Descrizione</b>: Scheda Eventi associata alle tab RICOVERI, EPISODI, TRASFUSIONI, SIEROCONVERSIONI
'
' @remarks
'
' @author
'
' @date 05/02/2011 18.41
Option Explicit

'' rs della scheda
Dim rsEventi As Recordset
Dim vCol As Integer
Dim vRow As Integer
'' oggetto che gestisce l'annullamento dei dati nelle flx
Dim objAnnulla(3) As CAnnulla
'' rs per la traccitura
Dim rsDisco(3) As Recordset
Dim rsDialisi As Recordset
Dim intPazientiKey As Integer

Const ICS As String = "         X"

'' Ricarica la cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("TIPO_TRASFUSIONI", "NOME", cboTrasfusioni)
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
    Set rsEventi = Nothing
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim k As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    tabSchede.Tab = 0
    For i = 0 To 3
        With flxGriglia(i)
            .MousePointer = flexCustom
            .Row = 0
            .ColWidth(0) = 0
            .ColAlignment(1) = vbLeftJustify
            For k = 1 To flxGriglia(i).Cols - 1
                .Col = k
                .CellFontBold = True
                .ColAlignment(k) = vbLeftJustify
            Next k
        End With
        Set objAnnulla(i) = New CAnnulla
    Next i
    ' totale allineato a destra
    flxGriglia(0).ColAlignment(3) = vbRightJustify
    tabSchede.Tab = 0
    Call ApriRsDisconnesso
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

'' Apre i recordset disconnessi per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim k As Integer
    Dim v_nomeTabelle() As Variant
    Dim rsDataset As New Recordset
    v_nomeTabelle = Array("RICOVERI", "EPISODI", "TRASFUSIONI", "SIEROCONVERSIONI")
    For k = 0 To 3
        rsDataset.Open v_nomeTabelle(k), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
        Set rsDisco(k) = New ADODB.Recordset
        For i = 0 To rsDataset.Fields.count - 1
            rsDisco(k).Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
        Next i
        rsDisco(k).CursorLocation = adUseClient
        rsDisco(k).Open , , adOpenDynamic, adLockOptimistic
        rsDataset.Close
    Next k
    Set rsDataset = Nothing
End Sub

'' Confronta i campi per rilevare le eventuali modifiche
' e le salva nella relativa tabella delle modifiche
'
' @param index indice della sottoscheda che si sta modificando
' @param rs rs che contiene lo stato del record che si è memorizzato
Private Sub Confronta(Index As Integer, rs As Recordset)
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim v_modifiche() As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nome_campi As String
    Dim valori As String
    Dim trovato As Boolean
    Dim nomeTabella As String
    
    nomeTabella = Choose(Index + 1, "M_RICOVERI", "M_EPISODI", "M_TRASFUSIONI", "M_SIEROCONVERSIONI")
    ReDim v_modifiche(0)
    ' filtra per la presenza di piu record
    rsDisco(Index).Filter = "(KEY=" & rs("KEY") & ")"
    For i = 0 To rsDisco(Index).Fields.count - 1
        trovato = False
        If IsNull(rsDisco(Index).Fields(i)) Or IsNull(rs(i)) Then
            If Not (IsNull(rsDisco(Index).Fields(i)) And IsNull(rs(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(Index).Fields(i) <> rs(i) Then
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
            nome_campi = nome_campi & rsDisco(Index).Fields((v_modifiche(i))).Name & "&-&"
            valori = valori & IIf(IsNull(rsDisco(Index).Fields((v_modifiche(i)))), "NULL", rsDisco(Index).Fields((v_modifiche(i)))) & "&-&"
            ' aggiorna il rsDisco(index)
            rsDisco(Index)(v_modifiche(i)) = rs(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_RECORD", "CODICE_PAZIENTE", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, rs("KEY"), intPazientiKey, nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open nomeTabella, cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Salva i dati modificati
'
' @param index indice della sottoscheda che si sta modificando
Private Sub SalvaModifiche(Index As Integer)
    Dim val As Variant
    Dim nomi(3, 2) As Variant
    Dim nomeTabella As String
    nomi(0, 2) = "NOTE"
    nomi(0, 0) = "DAL"
    nomi(0, 1) = "AL"
    nomi(1, 0) = "DATA"
    nomi(1, 1) = "NOTE"
    nomi(2, 0) = "DATA"
    nomi(2, 1) = "TIPO_TRASFUSIONE"
    nomi(3, 0) = "DATA"
    nomi(3, 1) = "HBV"
    nomi(3, 2) = "HCV"
    Select Case Index
        Case 0:
            nomeTabella = "RICOVERI"
        Case 1:
            nomeTabella = "EPISODI"
        Case 2:
            nomeTabella = "TRASFUSIONI"
        Case 3:
            nomeTabella = "SIEROCONVERSIONI"
    End Select
    Select Case vCol
        Case 1, 4
            val = flxGriglia(Index).TextMatrix(vRow, vCol)
        Case 2
            If cboTrasfusioni.Text <> "" Then
                Call GestisciNuovo("TIPO_TRASFUSIONI", cboTrasfusioni)
            End If
            If Index = 2 Then
                val = GetNumeroDaNome("TIPO_TRASFUSIONI", "NOME", flxGriglia(Index).TextMatrix(vRow, vCol))
            ElseIf Index = 3 Then
                val = IIf(flxGriglia(Index).TextMatrix(vRow, vCol) = ICS, True, False)
            Else
                val = flxGriglia(Index).TextMatrix(vRow, vCol)
            End If
        Case 3
            If Index = 3 Then
                val = IIf(flxGriglia(Index).TextMatrix(vRow, vCol) = ICS, True, False)
            Else
                val = flxGriglia(Index).TextMatrix(vRow, vCol)
            End If
    End Select
    Set rsEventi = New Recordset
    rsEventi.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & flxGriglia(Index).TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsEventi.Update nomi(Index, vCol - IIf(vCol = 4, 2, 1)), val
    If TRACCIATO Then
        Call Confronta(Index, rsEventi)
    End If
    Set rsEventi = Nothing
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    For i = 0 To 3
        flxGriglia(i).Rows = 1
    Next i
    lblCognome = ""
    lblNome = ""
    lblEta = ""
    intPazientiKey = 0
    ' ripulisce anche gli oggetti annulla
    For i = 0 To 3
        objAnnulla(i).Refresh
    Next i
End Sub

'' Carica la flx
'
' @param index indice della flx da caricare 0 ricoveri, 1 episodi, 2 trasfusioni, 3 sieroconversioni
Private Sub CaricaFlx(Index As Integer)
    Dim ris As Integer
    Dim i As Integer
    Dim v_Nomi() As Variant
    Dim strFrom As String
    Dim strData As String
        
    Select Case Index
        Case 0:
            strFrom = "RICOVERI"
            v_Nomi = Array("DAL", "AL", "NOTE")
            strData = " ORDER BY DAL DESC"
        Case 1:
            strFrom = "EPISODI"
            v_Nomi = Array("DATA", "NOTE")
            strData = " ORDER BY DATA DESC"
        Case 2:
            v_Nomi = Array("DATA", "TIPO_TRASFUSIONE")
            strFrom = "(TRASFUSIONI N INNER JOIN TIPO_TRASFUSIONI T ON T.KEY=N.TIPO_TRASFUSIONE)"
            strData = " ORDER BY DATA DESC"
        Case 3:
            strFrom = "SIEROCONVERSIONI"
            v_Nomi = Array("DATA", "HBV", "HCV")
            strData = " ORDER BY DATA DESC"
            
    End Select
    flxGriglia(Index).Rows = 1
    vCol = 0
    vRow = 0
    Set rsEventi = New Recordset
    rsEventi.Open "SELECT * FROM " & strFrom & " WHERE CODICE_PAZIENTE=" & intPazientiKey & " " & strData, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsEventi.EOF And rsEventi.BOF) Then
        With flxGriglia(Index)
            For i = 0 To 3
                Do While Not rsDisco(i).EOF
                    rsDisco(i).Delete
                    rsDisco(i).MoveNext
                Loop
            Next i
            Do While Not rsEventi.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsEventi(IIf(Index = 2, "N.", "") & "KEY")
                .TextMatrix(.Rows - 1, 1) = rsEventi(v_Nomi(0)) & ""
                If Index = 2 Then
                    .TextMatrix(.Rows - 1, 2) = rsEventi("NOME")
                ElseIf Index = 3 Then
                    .TextMatrix(.Rows - 1, 2) = IIf(CBool(rsEventi(v_Nomi(1))), ICS, "")
                Else
                    .TextMatrix(.Rows - 1, 2) = rsEventi(v_Nomi(1)) & ""
                End If
                If Index = 0 Then
                    .TextMatrix(.Rows - 1, 4) = rsEventi(v_Nomi(2)) & ""
                    If .TextMatrix(.Rows - 1, 2) <> "" Then
                        ris = CalcolaTotale(.Rows - 1)
                        .TextMatrix(.Rows - 1, 3) = Space(21 - Len(CStr(ris)) + 1) & ris
                    Else
                        .TextMatrix(.Rows - 1, 3) = ""
                    End If
                ElseIf Index = 3 Then
                    .TextMatrix(.Rows - 1, 3) = IIf(CBool(rsEventi(v_Nomi(2))), ICS, "")
                End If
                
                ' aggiorna i dati nel rsDisco
                rsDisco(Index).AddNew
                For i = 0 To rsDisco(Index).Fields.count - 1
                    rsDisco(Index).Fields(i) = rsEventi.Fields(i)
                Next i
                rsDisco(Index).Update
                
                rsEventi.MoveNext
            Loop
        End With
    End If
    Set rsEventi = Nothing
    flxGriglia(Index).Row = 0
End Sub

Private Function CalcolaTotale(riga As Integer) As Integer
    With flxGriglia(0)
        If .TextMatrix(riga, 1) <> "" And .TextMatrix(riga, 2) <> "" Then
            CalcolaTotale = CDate(.TextMatrix(riga, 2)) - CDate(.TextMatrix(riga, 1))
        End If
    End With
End Function

'' Carica tutte le sottoschede
Private Sub CaricaScheda()
    Dim i As Integer
    For i = 0 To 3
        Call CaricaFlx(i)
        objAnnulla(i).Refresh
    Next i
End Sub

Private Sub cmdAnnulla_Click(Index As Integer)
    Dim Dato As String
    Dim Col As Integer
    Dim RowKey As Integer
    Dim i As Integer
    Dato = objAnnulla(Index).Dato
    Col = objAnnulla(Index).Col
    RowKey = objAnnulla(Index).Row
    ' cerca la riga con il key memorizzato in rowkey
    With flxGriglia(Index)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = RowKey Then
                ' annulla
                .TextMatrix(i, Col) = Dato
                objAnnulla(Index).Remove
                ' modifica anche il db
                vRow = i
                vCol = Col
                Call SalvaModifiche(Index)
                If objAnnulla(Index).Vuoto = True Then
                    cmdAnnulla(Index).Enabled = False
                End If
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub cmdInserisci_Click(Index As Integer)
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim num As Integer                  ' il key del nuovo record
    Dim nomeTabella As String
    
    If intPazientiKey = 0 Then Exit Sub
    Unload frmInput
    tInput.Tipo = tpIRICOVERI + Index ' sfrutto il fatto che sono consecutivi
    frmInput.Show 1
    If Not (tInput.v_valori(1) = "") Then
        Select Case Index
            Case 0:
                nomeTabella = "RICOVERI"
                v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DAL", "AL", "NOTE")
            Case 1:
                nomeTabella = "EPISODI"
                v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "NOTE")
            Case 2:
                nomeTabella = "TRASFUSIONI"
                v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "TIPO_TRASFUSIONE")
            Case 3:
                nomeTabella = "SIEROCONVERSIONI"
                v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "HBV", "HCV")
        End Select
        num = GetNumero(nomeTabella)
        If Index = 0 Or Index = 3 Then
            If Index = 3 Then
                v_Val = Array(num, intPazientiKey, CDate(tInput.v_valori(1)), CBool(tInput.v_valori(2)), CBool(tInput.v_valori(3)))
            Else
                v_Val = Array(num, intPazientiKey, CDate(tInput.v_valori(1)), IIf(tInput.v_valori(2) <> "", tInput.v_valori(2), Null), tInput.v_valori(3))
            End If
        Else
            v_Val = Array(num, intPazientiKey, tInput.v_valori(1), tInput.v_valori(2))
        End If
        Set rsEventi = New Recordset
        rsEventi.Open nomeTabella, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsEventi.AddNew v_Nomi, v_Val
        rsEventi.Update
        Set rsEventi = Nothing
        
        ' aggiorna i dati nel rsDisco
        rsDisco(Index).AddNew v_Nomi, v_Val
        rsDisco(Index).Update
        
        ' aggiorna la flx
        Call CaricaFlx(Index)
        
        ' si posiziona sul record e lo seleziona
        flxGriglia(Index).Row = Esiste(flxGriglia(Index), 0, vRow, num)
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
        If flxGriglia(Index).Row > 8 Then
            flxGriglia(Index).TopRow = flxGriglia(Index).Row
        End If
        
'        MsgBox "Inserimento effettuato", vbInformation, "Inserimento"
    End If
End Sub

Private Sub cmdStampa_Click(Index As Integer)
    Dim codiceId As Integer
    Dim strSql As String
    Dim i As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If

      
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDialisi("COGNOME") & " " & rsDialisi("NOME")
    structIntestazione.sDataPaziente = rsDialisi("DATA_NASCITA")
    codiceId = rsDialisi("CODICE_ID")
    Set rsDialisi = Nothing

    strSql = "SHAPE APPEND  NEW adLongVarChar  as NOTE_RICOVERI, " & _
                    "       NEW adLongVarChar as DATA_INIZIO, " & _
                    "       NEW adLongVarChar as DATA_FINE, " & _
                    "       NEW adLongVarChar as TOTALE_GIORNI, "
    strSql = strSql & _
                    "       NEW adLongVarChar as DATA_EDEMA_POLMONARE, " & _
                    "       NEW adLongVarChar as NOTE_EDEMA_POLMONARE, "
    strSql = strSql & _
                    "       NEW adLongVarChar as DATA_TRASFUSIONE, " & _
                    "       NEW adLongVarChar as NOTE_TRASFUSIONE, "
    strSql = strSql & _
                    "       NEW adLongVarChar as DATA_SIEROCONVERSIONI, " & _
                    "       NEW adLongVarChar as SIEROCONVERSIONI_HBV, " & _
                    "       NEW adLongVarChar as SIEROCONVERSIONI_HCV "

                 
         
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Set rsDialisi = New Recordset
        

    With rsMain
        .AddNew
       
        For i = 1 To flxGriglia(0).Rows - 1            ' Ricoveri
            .Fields("NOTE_RICOVERI") = .Fields("NOTE_RICOVERI") & vbCrLf & flxGriglia(0).TextMatrix(i, 4)
            .Fields("DATA_INIZIO") = .Fields("DATA_INIZIO") & vbCrLf & flxGriglia(0).TextMatrix(i, 1)
            .Fields("DATA_FINE") = .Fields("DATA_FINE") & vbCrLf & flxGriglia(0).TextMatrix(i, 2)
            .Fields("TOTALE_GIORNI") = .Fields("TOTALE_GIORNI") & vbCrLf & flxGriglia(0).TextMatrix(i, 3)
        Next i
        
        For i = 1 To flxGriglia(1).Rows - 1             ' Episodi Edema Polmonare
            .Fields("DATA_EDEMA_POLMONARE") = .Fields("DATA_EDEMA_POLMONARE") & vbCrLf & flxGriglia(1).TextMatrix(i, 1)
            .Fields("NOTE_EDEMA_POLMONARE") = .Fields("NOTE_EDEMA_POLMONARE") & vbCrLf & flxGriglia(1).TextMatrix(i, 2)
        Next i
        
        For i = 1 To flxGriglia(2).Rows - 1             ' Trasfusioni
            .Fields("DATA_TRASFUSIONE") = .Fields("DATA_TRASFUSIONE") & vbCrLf & flxGriglia(2).TextMatrix(i, 1)
            .Fields("NOTE_TRASFUSIONE") = .Fields("NOTE_TRASFUSIONE") & vbCrLf & flxGriglia(2).TextMatrix(i, 2)
        Next i
        
        For i = 1 To flxGriglia(3).Rows - 1             ' Sieroconversioni
            .Fields("DATA_SIEROCONVERSIONI") = .Fields("DATA_SIEROCONVERSIONI") & vbCrLf & flxGriglia(3).TextMatrix(i, 1)
            .Fields("SIEROCONVERSIONI_HBV") = .Fields("SIEROCONVERSIONI_HBV") & vbCrLf & flxGriglia(3).TextMatrix(i, 2)
            .Fields("SIEROCONVERSIONI_HCV") = .Fields("SIEROCONVERSIONI_HCV") & vbCrLf & flxGriglia(3).TextMatrix(i, 3)
        Next i
                               
    End With

    Set rptEventi.DataSource = rsMain
    rptEventi.TopMargin = 0
    rptEventi.BottomMargin = 0
    rptEventi.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptEventi.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptEventi.PrintReport True, rptRangeAllPages
    
    Set rsDialisi = Nothing
End Sub

Private Sub cmdChiudi_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    Call CaricaPaziente
End Sub

Private Sub flxGriglia_Click(Index As Integer)
    Dim i As Integer
    vCol = flxGriglia(Index).Col
    flxGriglia(Index).SetFocus
    If VerificaClickFlx(flxGriglia(Index)) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1, True)
        ' annulla le row e col
        flxGriglia(Index).Row = 0
        flxGriglia(Index).Col = 0
    Else
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
        flxGriglia(Index).Col = vCol
        vRow = flxGriglia(Index).Row
        ' discolora le altre flx
        For i = 0 To 2
            If i <> Index Then
                ' discolora
                Call ColoraFlx(flxGriglia(i), flxGriglia(i).Cols - 1, True)
                ' annulla le row e col
                flxGriglia(i).Row = 0
                flxGriglia(i).Col = 0
            End If
        Next i
    End If
End Sub

Private Sub flxGriglia_Scroll(Index As Integer)
    If txtAppo(Index).Visible Then
        txtAppo(Index).Top = flxGriglia(Index).rowPos(flxGriglia(Index).Row) + flxGriglia(Index).Top + 45
    End If
    If cboTrasfusioni.Visible Then
        cboTrasfusioni.Top = flxGriglia(Index).rowPos(flxGriglia(Index).Row) + flxGriglia(Index).Top + 45
    End If
End Sub

Private Sub flxGriglia_DblClick(Index As Integer)
    ' fase di modifica
    Dim ris As Integer
    
    If VerificaClickFlx(flxGriglia(Index)) = False Then Exit Sub
    With flxGriglia(Index)
        If Index = 0 Then
            If vCol = 3 Then Exit Sub
            If vCol = 4 Then
                txtAppo(Index).Left = .colPos(.Col) + .Left + 45
                txtAppo(Index).Top = .rowPos(.Row) + .Top + 45
                txtAppo(Index).Width = .ColWidth(.Col)
                txtAppo(Index).Text = .TextMatrix(.Row, .Col)
                txtAppo(Index).Visible = True
                txtAppo(Index).SetFocus
            Else
                frmCalendario.Show 1
                If laData <> "" Then
                    If laData <> .TextMatrix(.Row, .Col) Then
                        Call objAnnulla(Index).Add(flxGriglia(Index).TextMatrix(vRow, vCol), vCol, vRow)
                        cmdAnnulla(Index).Enabled = True
                        .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                        Call SalvaModifiche(Index)
                        ' cambia colonna per evitave di ricaricare il calendario
                        .Col = 0
                        ris = CalcolaTotale(.Row)
                        .TextMatrix(.Row, 3) = Space(21 - Len(CStr(ris)) + 1) & ris
                    End If
                End If
            End If
        ElseIf Index = 1 Then
            If vCol = 2 Then
                txtAppo(Index).Left = .colPos(.Col) + .Left + 45
                txtAppo(Index).Top = .rowPos(.Row) + .Top + 45
                txtAppo(Index).Width = .ColWidth(.Col)
                txtAppo(Index).Text = .TextMatrix(.Row, .Col)
                txtAppo(Index).Visible = True
                txtAppo(Index).SetFocus
            Else
                frmCalendario.Show 1
                Call objAnnulla(Index).Add(flxGriglia(Index).TextMatrix(vRow, vCol), vCol, vRow)
                cmdAnnulla(Index).Enabled = True
                .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                Call SalvaModifiche(Index)
                ' cambia colonna per evitave di ricaricare il calendario
                .Col = 0
            End If
        ElseIf Index = 2 Then
            If vCol = 2 Then
                cboTrasfusioni.Left = .colPos(.Col) + .Left + 45
                cboTrasfusioni.Top = .rowPos(.Row) + .Top + 45
                cboTrasfusioni.Text = .TextMatrix(.Row, .Col)
                cboTrasfusioni.Visible = True
                cboTrasfusioni.SetFocus
            Else
                frmCalendario.Show 1
                Call objAnnulla(Index).Add(flxGriglia(Index).TextMatrix(vRow, vCol), vCol, vRow)
                cmdAnnulla(Index).Enabled = True
                .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                Call SalvaModifiche(Index)
                ' cambia colonna per evitave di ricaricare il calendario
                .Col = 0
            End If
        Else
            Call objAnnulla(Index).Add(flxGriglia(Index).TextMatrix(vRow, vCol), vCol, vRow)
            cmdAnnulla(Index).Enabled = True
            If vCol = 1 Then
                frmCalendario.Show 1
                .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
                ' cambia colonna per evitave di ricaricare il calendario
                .Col = 0
            Else
                If .TextMatrix(.Row, .Col) = ICS Then
                    .TextMatrix(.Row, .Col) = ""
                Else
                    .TextMatrix(.Row, .Col) = ICS
                End If
            End If
            Call SalvaModifiche(Index)
        End If
    End With
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    Dim i As Integer
'    For i = 0 To 3
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
'----------------------------------

Private Sub txtAppo_LostFocus(Index As Integer)
    If UCase(flxGriglia(Index).TextMatrix(vRow, vCol)) <> UCase(txtAppo(Index)) Then
        Call objAnnulla(Index).Add(flxGriglia(Index).TextMatrix(vRow, vCol), vCol, vRow)
        cmdAnnulla(Index).Enabled = True
        flxGriglia(Index).TextMatrix(vRow, vCol) = txtAppo(Index).Text
        Call SalvaModifiche(Index)
    End If
    txtAppo(Index).Visible = False
End Sub

Private Sub txtAppo_GotFocus(Index As Integer)
    txtAppo(Index).SelStart = 0
    txtAppo(Index).SelLength = Len(txtAppo(Index))
End Sub

Private Sub txtAppo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia(Index).SetFocus
    End If
End Sub

Private Sub cboTrasfusioni_Click()
    flxGriglia(2).TextMatrix(flxGriglia(2).Row, flxGriglia(2).Col) = cboTrasfusioni.Text
    cboTrasfusioni.Visible = False
End Sub

Private Sub cboTrasfusioni_LostFocus()
    Call SalvaModifiche(2)
    cboTrasfusioni.Visible = False
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
    ' carica la scheda
    Call CaricaScheda
End Sub

