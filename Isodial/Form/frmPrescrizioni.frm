VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmPrescrizioni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " GESTIONE RICETTE"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPrestazioni 
      Caption         =   "Prestazioni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   39
      Top             =   4770
      Width           =   12015
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Inserisci prescrizione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   9720
         TabIndex        =   21
         Top             =   1220
         Width           =   2055
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "&Elimina prescrizione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   7440
         TabIndex        =   20
         Top             =   1220
         Width           =   1935
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
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboPrescrizioni 
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
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   5535
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
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   360
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   975
         Left            =   120
         TabIndex        =   32
         Top             =   250
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   1720
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   $"frmPrescrizioni.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPrescrizioni.frx":00B8
      End
      Begin VB.Label lblImpegnate 
         Alignment       =   2  'Center
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
         Left            =   3720
         TabIndex        =   64
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "- Impegnabili"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4200
         TabIndex        =   63
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "- Impegnate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2400
         TabIndex        =   62
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lblSedute 
         Alignment       =   2  'Center
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
         Left            =   2040
         TabIndex        =   61
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Sedute Registrate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblImpegnabili 
         Alignment       =   2  'Center
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
         Left            =   5760
         TabIndex        =   59
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   49
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmPrescrizioni.frx":0212
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
         Left            =   2160
         TabIndex        =   55
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
         Left            =   6720
         TabIndex        =   54
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
         Left            =   11040
         TabIndex        =   53
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
         Left            =   10440
         TabIndex        =   52
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
         Left            =   6000
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dati Ricetta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1845
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   12015
      Begin VB.CheckBox chkPresenzaBarCode 
         Caption         =   "Barcode Cod.Fisc. su ricetta"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   1416
         Width           =   3615
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
         ItemData        =   "frmPrescrizioni.frx":066B
         Left            =   5520
         List            =   "frmPrescrizioni.frx":066D
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
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmPrescrizioni.frx":066F
         Left            =   2280
         List            =   "frmPrescrizioni.frx":0671
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtProgressivoAnnuale 
         Alignment       =   2  'Center
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
         Left            =   8600
         MaxLength       =   4
         TabIndex        =   7
         Top             =   652
         Width           =   615
      End
      Begin VB.CheckBox chkStampaPC 
         Caption         =   "Stampa PC"
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
         Left            =   8640
         TabIndex        =   5
         Top             =   1416
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtProgressivoRicetta 
         Alignment       =   2  'Center
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
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   4
         Top             =   652
         Width           =   615
      End
      Begin VB.TextBox txtMazzettaPrimo 
         Alignment       =   2  'Center
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
         Left            =   10560
         MaxLength       =   3
         TabIndex        =   9
         Top             =   652
         Width           =   615
      End
      Begin VB.TextBox txtMazzettaSecondo 
         Alignment       =   2  'Center
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
         Left            =   11320
         MaxLength       =   5
         TabIndex        =   10
         Top             =   652
         Width           =   615
      End
      Begin VB.ComboBox cboTipoPrescrizione 
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
         ItemData        =   "frmPrescrizioni.frx":0673
         Left            =   6720
         List            =   "frmPrescrizioni.frx":0683
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1034
         Width           =   1215
      End
      Begin VB.TextBox txtNumeroRicetta 
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
         MaxLength       =   15
         TabIndex        =   3
         Top             =   652
         Width           =   1695
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   8
         Top             =   1350
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
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero ricetta"
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
         Left            =   120
         TabIndex        =   69
         Top             =   652
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N°Mazzetta"
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
         Left            =   9350
         TabIndex        =   68
         Top             =   652
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N°Progr.Ricetta"
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
         Left            =   4560
         TabIndex        =   67
         Top             =   652
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N°Progr.Interno"
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
         Left            =   6960
         TabIndex        =   66
         Top             =   652
         Width           =   1590
      End
      Begin VB.Label Label1 
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
         Index           =   32
         Left            =   4800
         TabIndex        =   48
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mese di fatturazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   46
         Top             =   270
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "/"
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
         Left            =   11220
         TabIndex        =   45
         Top             =   652
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di prescrizione"
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
         Left            =   4560
         TabIndex        =   44
         Top             =   1034
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data ricetta"
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
         TabIndex        =   42
         Top             =   1034
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data prenotazione"
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
         TabIndex        =   34
         Top             =   1416
         Width           =   1920
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Medico Prescrittore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1000
      Left            =   120
      TabIndex        =   35
      Top             =   2670
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   360
         Picture         =   "frmPrescrizioni.frx":06A9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Seleziona il medico"
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblCognomeMedico 
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
         TabIndex        =   58
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblNomeMedico 
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
         Left            =   6240
         TabIndex        =   57
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblCodiceTimbroMedico 
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
         Left            =   10800
         TabIndex        =   56
         Top             =   480
         Width           =   975
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
         Index           =   13
         Left            =   960
         TabIndex        =   47
         Top             =   480
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
         Index           =   33
         Left            =   5520
         TabIndex        =   37
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Regionale"
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
         Index           =   37
         Left            =   9600
         TabIndex        =   36
         Top             =   360
         Width           =   1170
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipologia Esenzione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1125
      Left            =   120
      TabIndex        =   38
      Top             =   3660
      Width           =   12015
      Begin VB.ComboBox cboTipoErogazione 
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
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   5055
      End
      Begin VB.CheckBox chkEsenzioneDoppia 
         Caption         =   "Esenzione E05"
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
         Left            =   6000
         TabIndex        =   16
         Top             =   375
         Width           =   3615
      End
      Begin VB.ComboBox cboEsenzione 
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
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkEsenteReddito 
         Caption         =   "Esente per reddito"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di erogazione della prestazione"
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
         Left            =   120
         TabIndex        =   65
         Top             =   720
         Width           =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esenzione"
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
         TabIndex        =   41
         Top             =   375
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   240
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   40
      Top             =   6480
      Width           =   12015
      Begin VB.CommandButton cmdCancellaRicetta 
         Caption         =   "&Cancella"
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
         Left            =   3720
         TabIndex        =   23
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
         Left            =   10440
         TabIndex        =   27
         Top             =   240
         Width           =   1335
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
         Left            =   8760
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCarica 
         Caption         =   "&Ricerca"
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
         Left            =   7080
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuovaRicetta 
         Caption         =   "&Nuova"
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
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvalidaNumeroRicetta 
         Caption         =   "&Sostituisci numero ricetta"
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
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraPazientiEsteri 
      Caption         =   "Tipologia Ricetta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   70
      Top             =   4770
      Width           =   12015
      Begin VB.TextBox txtNumeroIdentificazioneTessera 
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
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1125
         Width           =   2535
      End
      Begin VB.TextBox txtNumeroIdentificazionePersonale 
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
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   765
         Width           =   2175
      End
      Begin VB.TextBox txtCodiceIstituzioneCompetente 
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
         Left            =   3360
         MaxLength       =   28
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   765
         Width           =   2535
      End
      Begin VB.ComboBox cboTipoRicetta 
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
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   6615
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   2
         Left            =   9600
         TabIndex        =   76
         Top             =   1080
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Scadenza Tessera Europea"
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
         Index           =   19
         Left            =   6000
         TabIndex        =   75
         Top             =   1155
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Num. Identificazione Tessera"
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
         Index           =   18
         Left            =   120
         TabIndex        =   74
         Top             =   1125
         Width           =   3030
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Num. Identificazione Personale"
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
         Index           =   17
         Left            =   6000
         TabIndex        =   73
         Top             =   765
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Istituzione Competente"
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
         Left            =   120
         TabIndex        =   72
         Top             =   765
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo ricetta"
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
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmPrescrizioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPrescrizioni.frm
'
' <b>Descrizione</b>: Scheda Prescrizioni associata alla tab RICETTE
'
' @remarks
'
' @author
'
' @date 07/08/2011 11.41

Option Explicit

'' indica se si è in fase di modifica
Dim modifica As Boolean
'' rs della scheda
Dim rsRicette As Recordset
'' il key utilizzato in fase di modifica
Dim keyId As Integer
'' il vecchio key per la cancellazione
Dim keyIdVecchio As Integer
Dim vRow As Integer
Dim vCol As Integer
Dim lettera As String
Dim codiceDistretto As Integer
Dim stoPulendo As Boolean
Dim stoCaricando As Boolean
Dim intPazientiKey As Integer
Dim intMedicoKey As Integer
Dim blnPazienteEstero As Boolean
Dim ControlloCodiceFiscalePazienteEstero As String
Dim MImpegnabili As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("TIPOLOGIE_ESENZIONE", "CODICE", cboEsenzione)
    Call RicaricaComboBox("TIPI_EROGAZIONE", "NOME", cboTipoErogazione)
    If cboTipoErogazione.ListIndex = -1 Then cboTipoErogazione.ListIndex = GetCboListIndex(0, cboTipoErogazione)
    Call RicaricaComboBox("Select Codice + ' - ' + Nome as Nome, Key From TIPIRICETTA Where Key<>1", "NOME", cboTipoRicetta)
    If cboTipoRicetta.ListIndex = -1 Then cboTipoRicetta.ListIndex = GetCboListIndex(1, cboTipoRicetta)
    Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboCodici)
    Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "NOME", cboPrescrizioni)
        
    If intPazientiKey = 0 Then
        cmdTrova_Click (0)
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    MImpegnabili = 0
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    oData(0).data = date
    oData(1).data = date
    For i = 0 To 2
        oData(i).ConnectionString = strConnectionStringCentro
    Next
    With flxGriglia
        .MousePointer = 99
        .ColWidth(0) = 0
        .ColWidth(8) = 0
        .Rows = 1
        .Row = 0
        For i = 1 To 7
            .Col = i
            .CellFontBold = True
        Next i
    End With
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    cboTipoRicetta.ListIndex = GetCboListIndex(2, cboTipoRicetta)
    stoCaricando = False
    laData = date
    cboTipoPrescrizione.ListIndex = 0
    For i = 1 To 12
        cboMese.AddItem UCase(MonthName(i))
    Next i
    Call SettaProgressivi
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
    intMedicoKey = 0
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'--------------------------------------------


'' Propone il progressivo di un campo
'
' @param nomeCampo nome del campo per cui proporre il progressivo
Private Function CaricaProgressivo(nomeCampo As String) As Integer
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT MAX(" & nomeCampo & ") AS MASSIMO FROM RICETTE WHERE ANNO=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsDataset("MASSIMO")) Then
        CaricaProgressivo = rsDataset("MASSIMO") + 1
    Else
        CaricaProgressivo = 1
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Function

'' Propone i numeri di mazzetta
'
' @param mazzetta1 variabile in output
' @param mazzetta2 variabile in output
Private Sub CaricaMazzette(ByRef mazzetta1 As Integer, ByRef mazzetta2 As Integer)
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "SELECT    MAX(MAZZETTA1) AS MASSIMO " & _
             "FROM      (RICETTE " & _
             "          INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
             "WHERE     ANNO=" & cboAnno.Text & " AND " & _
             "          MESE=" & cboMese.ListIndex + 1 & " AND " & _
             "          CODICE_DISTRETTO=" & codiceDistretto
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsDataset("MASSIMO")) Then
        mazzetta1 = rsDataset("MASSIMO")
    Else
        mazzetta1 = 1
        mazzetta2 = 1
    End If
    rsDataset.Close
    
    strSql = "SELECT    MAZZETTA2 " & _
             "FROM      (RICETTE " & _
             "          INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
             "WHERE     ANNO=" & cboAnno.Text & " AND " & _
             "          MAZZETTA1=" & mazzetta1 & " AND " & _
             "          MESE=" & cboMese.ListIndex + 1 & " AND " & _
             "          CODICE_DISTRETTO=" & codiceDistretto & " " & _
             "ORDER BY  MAZZETTA2 DESC"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        mazzetta2 = rsDataset("MAZZETTA2") + 1
        If mazzetta2 = 51 Then
            mazzetta1 = mazzetta1 + 1
            mazzetta2 = 1
        End If
    Else
        mazzetta2 = 1
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

'' Calcola i totali delle sedute registrate, impegnate e impegnabili
Private Sub CalcolaTotaliSeduteDialisi()
    Dim registrate As Integer
    Dim impegnate As Integer
    Dim impegnabili As Integer
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    If intPazientiKey = 0 Then Exit Sub
    
    strSql = "SELECT    COUNT(KEY) AS TOTALE " & _
             "FROM      SCHEDE_DIALISI " & _
             "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
             "          MONTH([DATA])=" & cboMese.ListIndex + 1 & " AND " & _
             "          YEAR([DATA])=" & cboAnno.Text & " AND " & _
             "          ERRATA=FALSE"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        registrate = rsDataset("TOTALE")
    Else
        registrate = 0
    End If
    rsDataset.Close
    
    strSql = "SELECT    SUM(QUANTITA) AS TOTALE " & _
             "FROM      (RICETTE " & _
             "          INNER JOIN PRESCRIZIONI ON PRESCRIZIONI.CODICE_RICETTA=RICETTE.KEY) " & _
             "WHERE     NOT FLAG=3 AND " & _
             "          CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
             "          ANNO=" & cboAnno.Text & " AND " & _
             "          MESE=" & cboMese.ListIndex + 1
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsDataset("TOTALE")) Then
        impegnate = rsDataset("TOTALE") + MImpegnabili
    Else
        impegnate = 0 + MImpegnabili
    End If
    rsDataset.Close
    
    impegnabili = registrate - impegnate
    
    lblSedute = registrate
    lblImpegnate = impegnate
    If impegnabili > 0 Then
        lblImpegnabili.ToolTipText = " PRESCRIZIONI MANCANTI "
        Timer1.Enabled = True
    ElseIf impegnabili < 0 Then
        lblImpegnabili.ToolTipText = " SEDUTE NON REGISTRATE "
        Timer1.Enabled = True
    End If
    lblImpegnabili = impegnabili
       
End Sub

Private Sub Timer1_Timer()
    If lblImpegnabili.ForeColor = vbRed Then
       lblImpegnabili.ForeColor = vbBlack
       lblImpegnabili.BackColor = vbRed
    Else
       lblImpegnabili.ForeColor = vbRed
       lblImpegnabili.BackColor = vbBlack
    End If
End Sub

'' Carica l'intera ricetta
Private Sub CaricaScheda()
    Set rsRicette = New Recordset
    
    rsRicette.Open "SELECT * FROM RICETTE WHERE KEY=" & keyId, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsRicette.EOF And rsRicette.BOF) Then
        
        oData(0).data = rsRicette("DATA_PRENOTAZIONE")
        oData(1).data = rsRicette("DATA_RICETTA")
        txtProgressivoRicetta = rsRicette("PROGRESSIVO_RICETTA")
        txtNumeroRicetta = rsRicette("NUMERO_RICETTA")
        cboMese.ListIndex = rsRicette("MESE") - 1
        stoCaricando = True
        cboAnno.ListIndex = IIf(cboAnno.List(0) = rsRicette("ANNO"), 0, 1)
        stoCaricando = False
        txtMazzettaPrimo = rsRicette("MAZZETTA1")
        txtMazzettaSecondo = rsRicette("MAZZETTA2")
        txtProgressivoAnnuale = rsRicette("PROGRESSIVO_ANNUALE")
        chkStampaPC.Value = IIf(CBool(rsRicette("STAMPATO_PC")), Checked, Unchecked)
        chkEsenteReddito.Value = IIf(CBool(rsRicette("ESENTE_REDDITO")), Checked, Unchecked)
        chkEsenzioneDoppia.Value = IIf(CBool(rsRicette("ESENZIONE_DOPPIA")), Checked, Unchecked)
        chkPresenzaBarCode.Value = IIf(CBool(rsRicette("PRESENZA_BARCODE")), Checked, Unchecked)
        cboEsenzione.ListIndex = GetCboListIndex(rsRicette("CODICE_ESENZIONE"), cboEsenzione)
        cboTipoErogazione.ListIndex = GetCboListIndex(rsRicette("CODICE_TIPO_EROGAZIONE"), cboTipoErogazione)
        cboTipoRicetta.ListIndex = GetCboListIndex(rsRicette("TIPIRICETTAID"), cboTipoRicetta)
        txtCodiceIstituzioneCompetente.Text = rsRicette("CodiceIstituzioneCompetente") & ""
        txtNumeroIdentificazionePersonale.Text = rsRicette("NumeroIdentificativoPersonale") & ""
        txtNumeroIdentificazioneTessera.Text = rsRicette("NumeroIdentificazioneTessera") & ""
        cboTipoPrescrizione.ListIndex = rsRicette("TIPOLOGIA_RICETTA")
        oData(2).data = rsRicette("DataScadenzaTessera") & ""
        If CBool(rsRicette("VALIDATA")) Or rsRicette("FLAG") = 2 Then
            cmdInvalidaNumeroRicetta.Visible = True
            txtNumeroRicetta.Enabled = False
            cboMese.Enabled = False
            cboAnno.Enabled = False
        Else
            cmdInvalidaNumeroRicetta.Visible = False
            txtNumeroRicetta.Enabled = True
            cboMese.Enabled = True
            cboAnno.Enabled = True
        End If
        
        intMedicoKey = rsRicette("CODICE_MEDICO")
        Call CaricaMedico
        Call CaricaPrescrizioni
        
        modifica = True
    Else
        modifica = False

    End If
    Set rsRicette = Nothing
End Sub

'' Setta i progressivi annuale, ricetta e le mazzette
Private Sub SettaProgressivi()
    Dim mazzetta1 As Integer
    Dim mazzetta2 As Integer
    
    txtProgressivoAnnuale = CaricaProgressivo("PROGRESSIVO_ANNUALE")
    txtProgressivoRicetta = CaricaProgressivo("PROGRESSIVO_RICETTA")
    Call CaricaMazzette(mazzetta1, mazzetta2)
    txtMazzettaPrimo = mazzetta1
    txtMazzettaSecondo = mazzetta2
End Sub

'' Carica tutte le prescrizioni associate alla ricetta
Private Sub CaricaPrescrizioni()
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "SELECT    PRESCRIZIONI.*, NOMENCLATORE_TARIFFARIO.CODICE, NOMENCLATORE_TARIFFARIO.NOME " & _
             "FROM      (PRESCRIZIONI " & _
             "          LEFT OUTER JOIN NOMENCLATORE_TARIFFARIO ON NOMENCLATORE_TARIFFARIO.KEY=PRESCRIZIONI.CODICE_PRESTAZIONE) " & _
             "WHERE     CODICE_RICETTA=" & keyId
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    flxGriglia.Rows = 1
    vRow = 0
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("CODICE")
            .TextMatrix(.Rows - 1, 2) = rsDataset("NOME")
            .TextMatrix(.Rows - 1, 3) = rsDataset("QUANTITA")
            .TextMatrix(.Rows - 1, 4) = VirgolaOrPunto(Format(rsDataset("IMPORTO"), "####.00"), ",")
            .TextMatrix(.Rows - 1, 5) = VirgolaOrPunto(Format(rsDataset("QUANTITA") * CSng(rsDataset("IMPORTO")), "####.00"), ",")
            .TextMatrix(.Rows - 1, 6) = rsDataset("DATA_INIZIO")
            .TextMatrix(.Rows - 1, 7) = rsDataset("DATA_FINE")
            .TextMatrix(.Rows - 1, 8) = VirgolaOrPunto(rsDataset("IMPORTO_SCONTATO"), ",")
        End With
        rsDataset.MoveNext
    Loop
    Set rsDataset = Nothing
    flxGriglia.Row = 0
End Sub

'' Pulisce la scheda
Private Sub Pulisci()
    lblImpegnabili.BackColor = &H8000000F
    lblImpegnabili.ForeColor = vbRed
    MImpegnabili = 0
    modifica = False
    stoPulendo = True
    keyId = 0
    codiceDistretto = 0
    keyIdVecchio = 0
    intMedicoKey = 0
    intPazientiKey = 0
    lblCognomeMedico = ""
    lblNomeMedico = ""
    lblCodiceTimbroMedico = ""
    lblSedute = ""
    lblImpegnate = ""
    lblImpegnabili = ""
    Call PulisciForm(Me)
    oData(0).Pulisci
    oData(1).Pulisci
    oData(2).Pulisci
    chkStampaPC.Value = Checked
    chkEsenteReddito.Value = Unchecked
    chkEsenzioneDoppia.Value = Unchecked
    chkPresenzaBarCode.Value = Unchecked
    cboTipoPrescrizione.ListIndex = 0
    cboTipoErogazione.ListIndex = GetCboListIndex(0, cboTipoErogazione)
    cboTipoRicetta.ListIndex = GetCboListIndex(1, cboTipoRicetta)
    txtNumeroIdentificazionePersonale.Text = ""
    txtNumeroIdentificazioneTessera.Text = ""
    txtCodiceIstituzioneCompetente.Text = ""
    txtNumeroRicetta.Enabled = True
    cboMese.Enabled = True
    cboAnno.Enabled = True
    stoCaricando = True
    cboAnno.ListIndex = 0
    cboTipoRicetta.ListIndex = GetCboListIndex(2, cboTipoRicetta)
    stoCaricando = False
    cmdInvalidaNumeroRicetta.Visible = False
    txtProgressivoAnnuale.Enabled = True
    txtProgressivoRicetta.Enabled = True
    txtMazzettaPrimo.Enabled = True
    txtMazzettaSecondo.Enabled = True
    cboMese.ListIndex = Month(Now) - 1
'    oData(0).data = date
    flxGriglia.Rows = 1
    Call SettaProgressivi
    oData(0).data = date
    stoPulendo = False
End Sub

'' Verifica se i progressivi inseriti siano disponibili
Private Function verificaProgressivi() As Boolean
    Dim rsDataset As New Recordset
    Dim strCondizione As String
    Dim mazzetta1 As Integer
    Dim mazzetta2 As Integer
    Dim strSql As String
    
    If modifica Then
        strCondizione = " AND NOT RICETTE.KEY=" & keyId
    Else
        strCondizione = ""
    End If
    
    strSql = "SELECT    * " & _
             "FROM      RICETTE " & _
             "WHERE     NOT FLAG=3 AND " & _
             "          NOT KEY=" & keyIdVecchio & " AND " & _
             "          ANNO=" & cboAnno.Text & " AND " & _
             "          PROGRESSIVO_ANNUALE=" & txtProgressivoAnnuale & strCondizione
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        MsgBox "Valore di progressivo interno già in uso" & vbCrLf & "Valore disponibile: " & CaricaProgressivo("PROGRESSIVO_ANNUALE"), vbCritical, "Attenzione"
        verificaProgressivi = False
        Exit Function
    End If
    rsDataset.Close
    
    strSql = "SELECT    * " & _
             "FROM      RICETTE " & _
             "WHERE     NOT FLAG=3 AND " & _
             "          NOT KEY=" & keyIdVecchio & " AND " & _
             "          ANNO=" & cboAnno.Text & " AND " & _
             "          PROGRESSIVO_RICETTA=" & txtProgressivoRicetta & " " & strCondizione
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        MsgBox "Valore di progressivo ricetta già in uso" & vbCrLf & "Valore disponibile: " & CaricaProgressivo("PROGRESSIVO_RICETTA"), vbCritical, "Attenzione"
        verificaProgressivi = False
        Exit Function
    End If
    rsDataset.Close
    
    strSql = "SELECT    RICETTE.* " & _
             "FROM      (RICETTE " & _
             "          INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
             "WHERE     NOT FLAG=3 AND " & _
             "          NOT RICETTE.KEY=" & keyIdVecchio & " AND " & _
             "          ANNO=" & cboAnno.Text & " AND " & _
             "          MAZZETTA1=" & txtMazzettaPrimo & " AND " & _
             "          MAZZETTA2=" & txtMazzettaSecondo & " AND " & _
             "          MESE=" & cboMese.ListIndex + 1 & " AND " & _
             "          CODICE_DISTRETTO=" & codiceDistretto & strCondizione
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        Call CaricaMazzette(mazzetta1, mazzetta2)
        MsgBox "Valori di mazzetta già in uso" & vbCrLf & "Valori disponibili: " & mazzetta1 & "/" & mazzetta2, vbCritical, "Attenzione"
        verificaProgressivi = False
        Exit Function
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
    verificaProgressivi = True
End Function

'' Salva le modifiche effettuate sulle prescrizioni nell flx
Private Sub SalvaModifiche()
    Dim valore() As Variant
    Dim nome() As Variant
    If Not modifica Then Exit Sub
    ReDim valore(0)
    ReDim nome(0)
    Select Case vCol
        Case 1, 2
            ReDim valore(2)
            ReDim nome(2)
            nome(0) = "CODICE_PRESTAZIONE"
            valore(0) = GetNumeroDaNome("NOMENCLATORE_TARIFFARIO", "NOME", flxGriglia.TextMatrix(vRow, 2))
            nome(1) = "IMPORTO"
            valore(1) = flxGriglia.TextMatrix(flxGriglia.Row, 4)
            nome(2) = "IMPORTO_SCONTATO"
            valore(2) = flxGriglia.TextMatrix(flxGriglia.Row, 8)
        Case 3
            nome(0) = "QUANTITA"
            valore(0) = flxGriglia.TextMatrix(vRow, vCol)
        Case 6
            nome(0) = "DATA_INIZIO"
            valore(0) = flxGriglia.TextMatrix(vRow, vCol)
        Case 7
            nome(0) = "DATA_FINE"
            valore(0) = flxGriglia.TextMatrix(vRow, vCol)
    End Select
    
    Set rsRicette = New Recordset
    rsRicette.Open "SELECT * FROM PRESCRIZIONI WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsRicette.Update nome, valore
    Set rsRicette = Nothing
End Sub

'' Salva le prescrizioni tutte assieme durante la fase di salvataggio ricetta
Private Function SalvaPrescrizioni() As Boolean
    On Error GoTo gestione
    Dim rsDataset As New Recordset
    Dim cmCommand As New Command
    Dim i As Integer
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    
    If modifica Then
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
        cmCommand.CommandText = "DELETE * FROM PRESCRIZIONI WHERE CODICE_RICETTA=" & keyId
        cmCommand.Execute
    End If
    
    v_Nomi = Array("KEY", "CODICE_PRESTAZIONE", "CODICE_RICETTA", "QUANTITA", "DATA_INIZIO", "DATA_FINE", "IMPORTO", "IMPORTO_SCONTATO")
    rsDataset.Open "PRESCRIZIONI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    With flxGriglia
        For i = 1 To .Rows - 1
            v_Val = Array(GetNumero("PRESCRIZIONI"), GetNumeroDaNome("NOMENCLATORE_TARIFFARIO", "NOME", .TextMatrix(i, 2)), keyId, .TextMatrix(i, 3), .TextMatrix(i, 6), .TextMatrix(i, 7), VirgolaOrPunto(.TextMatrix(i, 4), ","), VirgolaOrPunto(.TextMatrix(i, 8), ","))
            rsDataset.AddNew v_Nomi, v_Val
            rsDataset.Update
        Next i
    End With
    
    Set rsDataset = Nothing
    SalvaPrescrizioni = True
    Exit Function
    
gestione:
    MsgBox "Descrizione: Valore non valido", vbCritical, "Errore n°: " & Err.Number
    cnPrinc.RollbackTrans
    SalvaPrescrizioni = False
End Function

'' Verifica che la scheda sia completa prima di memorizzarla
Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData(0).data = "" Then
        MsgBox "Inserire la data di prenotazione", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData(1).data = "" Then
        MsgBox "Inserire la data di ricetta", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtNumeroRicetta = "" Then
        MsgBox "Inserire il numero di ricetta", vbCritical, "Attenzione"
        Exit Function
    End If
    If intMedicoKey = 0 Then
        MsgBox "Selezionare il medico", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtProgressivoRicetta = "" Then
        MsgBox "Inserire il progressivo di ricetta", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtProgressivoAnnuale = "" Then
        MsgBox "Inserire il progressivo interno", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtMazzettaPrimo = "" Or txtMazzettaSecondo = "" Then
        MsgBox "Inserire il numero di mazzetta", vbCritical, "Attenzione"
        Exit Function
    End If
    If Len(txtNumeroRicetta) <> txtNumeroRicetta.MaxLength Then
        MsgBox "Inserire il numero ricetta per intero", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboTipoPrescrizione.ListIndex = -1 Then
        MsgBox "Selezionare la tipologia della ricetta", vbCritical, "Attenzione"
        Exit Function
    End If
    If flxGriglia.Rows = 1 Then
        MsgBox "Inserire almeno una prescrizione", vbCritical, "Attenzione"
        Exit Function
    End If
    If esisteNumeroRicetta Then
        MsgBox "Numero di ricetta già esistente", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboTipoErogazione.ListIndex = -1 Then
        MsgBox "Selezionare il tipo di erogazione", vbCritical, "Attenzione"
        Exit Function
    End If
    If blnPazienteEstero Then
        Dim intTipoRicettaID As Integer
        intTipoRicettaID = cboTipoRicetta.ItemData(cboTipoRicetta.ListIndex)
        If intTipoRicettaID >= 3 Then
            If txtCodiceIstituzioneCompetente.Text = "" Then
                MsgBox "Inserire il codice di istituzione competente", vbCritical, "Attenzione"
                Exit Function
            End If
            If oData(2).data = "" Then
                MsgBox "Inserire la data di scadenza tessera europea", vbCritical, "Attenzione"
                Exit Function
            End If
        End If
        If intTipoRicettaID = 3 Or intTipoRicettaID = 5 Then
            If txtNumeroIdentificazionePersonale.Text = "" Then
                MsgBox "Inserire il numero di identificazione personale", vbCritical, "Attenzione"
                Exit Function
            End If
            If txtNumeroIdentificazioneTessera.Text = "" Then
                MsgBox "Inserire il numero di identificazione tessera", vbCritical, "Attenzione"
                Exit Function
            End If
        End If
        If UCase(Mid(ControlloCodiceFiscalePazienteEstero, 1, 3)) <> "STP" And cboTipoRicetta.ListIndex = 0 Then
            MsgBox "Opzione ST ammessa solo in presenza del codice STP", vbCritical, "Attenzione"
            Exit Function
        End If
    End If
    Completo = True
End Function

'' Verifica se il num ricetta è gia presente in archivio
Private Function esisteNumeroRicetta() As Boolean
    Dim rsDataset As New Recordset
    Dim strCondizione As String
    
    If modifica Then
        strCondizione = " AND NOT KEY=" & keyId
    Else
        strCondizione = ""
    End If
    
    rsDataset.Open "SELECT * FROM RICETTE WHERE (NUMERO_RICETTA='" & txtNumeroRicetta & "' " & strCondizione & ")", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        esisteNumeroRicetta = True
    Else
        esisteNumeroRicetta = False
    End If
    rsDataset.Close
End Function

Private Sub flxGriglia_Click()
    vCol = flxGriglia.Col
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
        vRow = 0
    Else
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        vRow = flxGriglia.Row
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
      .SetFocus
      Select Case .Col
        Case 1  ' codici
            cboCodici.Left = .colPos(.Col) + .Left + 45
            cboCodici.Top = .rowPos(.Row) + .Top + 45
            cboCodici.ListIndex = GetIndex(cboCodici, .TextMatrix(.Row, .Col))
            cboCodici.Visible = True
            cboCodici.SetFocus
        Case 2      ' prestazioni
            cboPrescrizioni.Left = .colPos(.Col) + .Left + 45
            cboPrescrizioni.Top = .rowPos(.Row) + .Top + 45
            cboPrescrizioni.ListIndex = GetIndex(cboPrescrizioni, .TextMatrix(.Row, .Col))
            cboPrescrizioni.Visible = True
            cboPrescrizioni.SetFocus
        Case 3      ' quantità
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        Case 6, 7      ' data inizio o fine
            frmCalendario.Show 1
            If laData = "" Then Exit Sub
            If .Col = 6 Then
                If .TextMatrix(.Row, 7) <> "" Then
                    If laData <= CDate(.TextMatrix(.Row, 7)) Then
                        .TextMatrix(.Row, .Col) = laData
                        .Col = 0
                        Call SalvaModifiche
                    Else
                        MsgBox "La data di inizio prestazione non può essere successiva alla data di fine prestazione", vbCritical, "Attenzione"
                    End If
                Else
                    .TextMatrix(.Row, .Col) = laData
                    .Col = 0
                    Call SalvaModifiche
                End If
            Else
                If .TextMatrix(.Row, 6) <> "" Then
                    If laData >= CDate(.TextMatrix(.Row, 6)) Then
                        If laData >= CDate(oData(1).data) Then
                            If laData >= CDate(oData(0).data) Then
                                .TextMatrix(.Row, .Col) = laData
                                .Col = 0
                                Call SalvaModifiche
                            Else
                                MsgBox "La data di fine prestazione non può essere antecedente alla data di prenotazione", vbCritical, "Attenzione"
                            End If
                        Else
                            MsgBox "La data di fine prestazione non può essere antecedente alla data di ricetta", vbCritical, "Attenzione"
                        End If
                    Else
                        MsgBox "La data di fine prestazione non può essere antecedente alla data di inizio prestazione", vbCritical, "Attenzione"
                    End If
                Else
                    .TextMatrix(.Row, .Col) = laData
                    .Col = 0
                    Call SalvaModifiche
                End If
            End If
      End Select
    End With
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
    If cboPrescrizioni.Visible Then
        cboPrescrizioni.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
    If cboCodici.Visible Then
        cboCodici.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

'' Inserisce una nuova prescrizione
Private Sub cmdInserisci_Click()
    Dim rsDataset As New Recordset
    Dim keyPrescrizione As Integer
    
    If intPazientiKey = 0 Then Exit Sub
    If oData(1).data = "" Then
        MsgBox "Inserire la data della ricetta", vbCritical, "Attenzione"
        Exit Sub
    End If
    If oData(0).data = "" Then
        MsgBox "Inserire la data della prenotazione", vbCritical, "Attenzione"
        Exit Sub
    End If
    Unload frmInput
    tInput.Tipo = tpIPRESCRIZIONI
    frmInput.Show 1
    If Not (tInput.v_valori(1) = -1) Then
        rsDataset.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE KEY=" & tInput.v_valori(1), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = 0
            .TextMatrix(.Rows - 1, 1) = rsDataset("CODICE")
            .TextMatrix(.Rows - 1, 2) = rsDataset("NOME")
            .TextMatrix(.Rows - 1, 3) = tInput.v_valori(2)
            .TextMatrix(.Rows - 1, 4) = VirgolaOrPunto(Format(rsDataset("IMPORTO"), "####.00"), ",")
            .TextMatrix(.Rows - 1, 5) = VirgolaOrPunto(Format(tInput.v_valori(2) * CSng(rsDataset("IMPORTO")), "####.00"), ",")
            .TextMatrix(.Rows - 1, 6) = tInput.v_valori(3)
            .TextMatrix(.Rows - 1, 7) = tInput.v_valori(4)
            .TextMatrix(.Rows - 1, 8) = VirgolaOrPunto(Format(rsDataset("IMPORTO_SCONTATO"), "####.00"), ",")
            rsDataset.Close
            If modifica Then
                rsDataset.Open "PRESCRIZIONI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                rsDataset.AddNew
                keyPrescrizione = GetNumero("PRESCRIZIONI")
                rsDataset("KEY") = keyPrescrizione
                rsDataset("CODICE_PRESTAZIONE") = tInput.v_valori(1)
                rsDataset("CODICE_RICETTA") = keyId
                rsDataset("QUANTITA") = tInput.v_valori(2)
                rsDataset("DATA_INIZIO") = tInput.v_valori(3)
                rsDataset("DATA_FINE") = tInput.v_valori(4)
                rsDataset("IMPORTO") = .TextMatrix(.Rows - 1, 4)
                rsDataset("IMPORTO_SCONTATO") = .TextMatrix(.Rows - 1, 8)
                rsDataset.Update
                rsDataset.Close
                flxGriglia.TextMatrix(flxGriglia.Rows - 1, 0) = keyPrescrizione
            End If
        End With
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = flxGriglia.Rows - 1
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        laData = Day(date) & "/" & 10 & "/" & Year(date)
        Call CaricaPso
        MImpegnabili = MImpegnabili + tInput.v_valori(2)
        Call CalcolaTotaliSeduteDialisi
    End If
    
    Set rsDataset = Nothing
End Sub

'' Elimina la ricetta e scala tutti i numeri progressivi
Private Sub cmdCancellaRicetta_Click()
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    Dim cmCommand As New Command
    Dim strSql As String
    
    If keyId = 0 Then
        MsgBox "Impossibile cancellare la ricetta", vbCritical, "Attenzione"
        Exit Sub
    End If
    rsDataset.Open "SELECT * FROM RICETTE WHERE KEY=" & keyId, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If CBool(rsDataset("VALIDATA")) Then
        MsgBox "Impossibile cancellare la ricetta" & vbCrLf & "Ricetta validata con il precedente flusso XML", vbCritical, "Attenzione"
        Exit Sub
    End If
    rsDataset.Close
    If MsgBox("Sicuro di voler cancellare la ricetta?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancellazione") = vbYes Then
        
        ' riodina i progressivi
        strSql = "SELECT    * " & _
                 "FROM      RICETTE " & _
                 "WHERE     KEY>" & keyId & " " & _
                 "ORDER BY  KEY DESC"
        rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        Do While Not rsDataset.EOF
            strSql = "SELECT    * " & _
                     "FROM      RICETTE " & _
                     "WHERE     KEY=" & rsDataset("KEY") - 1
            rsAppo.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                rsDataset("PROGRESSIVO_ANNUALE") = rsAppo("PROGRESSIVO_ANNUALE")
                rsDataset("PROGRESSIVO_RICETTA") = rsAppo("PROGRESSIVO_RICETTA")
            End If
            rsAppo.Close
            rsDataset.MoveNext
        Loop
        rsDataset.Close
        
        ' riordina le mazzette tenendo il conto il mese di riferimento, l'anno e il distretto
        strSql = "SELECT    RICETTE.* " & _
                 "FROM      RICETTE " & _
                 "WHERE     KEY IN (" & _
                 "              SELECT  RICETTE.KEY " & _
                 "              FROM    (RICETTE " & _
                 "                      INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
                 "              WHERE   MESE=" & cboMese.ListIndex + 1 & " AND " & _
                 "                      ANNO=" & cboAnno.Text & " AND " & _
                 "                      CODICE_DISTRETTO=" & codiceDistretto & _
                 "                  ) AND " & _
                 "          KEY>" & keyId & " " & _
                 "ORDER BY  KEY DESC"
        rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        Do While Not rsDataset.EOF
            strSql = "SELECT    RICETTE.* " & _
                     "FROM      RICETTE " & _
                     "WHERE     KEY IN (" & _
                     "              SELECT  RICETTE.KEY " & _
                     "              FROM    (RICETTE " & _
                     "                      INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
                     "              WHERE   MESE=" & cboMese.ListIndex + 1 & " AND " & _
                     "                      ANNO=" & cboAnno.Text & " AND " & _
                     "                      CODICE_DISTRETTO=" & codiceDistretto & _
                     "                  ) AND " & _
                     "          NOT KEY=" & rsDataset("KEY") & " AND " & _
                     "          KEY<" & rsDataset("KEY") & " " & _
                     "ORDER BY  KEY DESC"
            rsAppo.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If Not (rsAppo.EOF And rsAppo.BOF) Then
                rsDataset("MAZZETTA1") = rsAppo("MAZZETTA1")
                rsDataset("MAZZETTA2") = rsAppo("MAZZETTA2")
            End If
            rsAppo.Close
            rsDataset.MoveNext
        Loop
        rsDataset.Close
    
        ' elimina la ricetta
        rsDataset.Open "SELECT * FROM RICETTE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsDataset.Delete
        rsDataset.Update
        rsDataset.Close
        
        ' elimina le prestazioni
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
        cmCommand.CommandText = "DELETE * FROM PRESCRIZIONI WHERE CODICE_RICETTA=" & keyId
        cmCommand.Execute

        Call Pulisci
        MsgBox "La ricetta è stata cancellata", vbInformation, "Cancellazione"
        cmdTrova_Click (0)
    End If
End Sub

'' Invalida la ricetta
Private Sub cmdInvalidaNumeroRicetta_Click()
    If MsgBox("SEI SICURO DI SOSTITUIRE IL NUMERO DELLA RICETTA ?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancellazione") = vbYes Then
        keyIdVecchio = keyId
        txtNumeroRicetta.Enabled = True
        cboMese.Enabled = True
        cboAnno.Enabled = True
        txtNumeroRicetta.Text = ""
        modifica = False
        cmdInvalidaNumeroRicetta.Visible = False
        txtProgressivoAnnuale.Enabled = False
        txtProgressivoRicetta.Enabled = False
        txtMazzettaPrimo.Enabled = False
        txtMazzettaSecondo.Enabled = False
        txtNumeroRicetta.SetFocus
    End If

End Sub

'' Carica una ricetta passata
Private Sub cmdCarica_Click()
    Unload frmPrescrizioniPassate
    Load frmPrescrizioniPassate
    frmPrescrizioniPassate.LetkeyPaziente = intPazientiKey
    frmPrescrizioniPassate.Show 1
    keyId = frmPrescrizioniPassate.getkeyReturn
    If keyId <> 0 Then
        Call CaricaScheda
    End If
End Sub

'' Elimina la prescrizione
Private Sub cmdElimina_Click()
    Dim rsDataset As New Recordset
    Dim EliminaImpegnabili As Integer
    
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare la prescrizione da eliminare", vbCritical, "Attenzione"
    Else
        If flxGriglia.Rows = 2 Then
            MsgBox "Impossibile eliminare tutte le prescrizioni di una ricetta", vbCritical, "Attenzione"
            Exit Sub
        End If
        If flxGriglia.TextMatrix(flxGriglia.Row, 0) <> 0 Then
            rsDataset.Open "SELECT * FROM PRESCRIZIONI WHERE KEY=" & flxGriglia.TextMatrix(flxGriglia.Row, 0), cnPrinc, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                rsDataset.Delete
            End If
            rsDataset.Close
        End If
        If flxGriglia.Rows = 2 Then
            flxGriglia.Rows = 1
        Else
            EliminaImpegnabili = flxGriglia.TextMatrix(flxGriglia.Row, 3)
            flxGriglia.RemoveItem (flxGriglia.Row)
        End If
        vRow = 0
        flxGriglia.Row = 0
    End If
    MImpegnabili = MImpegnabili - EliminaImpegnabili
    Call CalcolaTotaliSeduteDialisi
End Sub

Private Sub cmdNuovaRicetta_Click()
    Dim rsDataset As New Recordset
    modifica = False
    keyId = 0
    oData(1).Pulisci
    oData(2).Pulisci
    txtNumeroRicetta = ""
    cmdInvalidaNumeroRicetta.Visible = False
    chkStampaPC.Value = Checked
    chkPresenzaBarCode.Value = Unchecked
    cboTipoPrescrizione.ListIndex = 0
    cboTipoErogazione.ListIndex = GetCboListIndex(0, cboTipoErogazione)
    cboTipoRicetta.ListIndex = GetCboListIndex(0, cboTipoRicetta)
    txtNumeroIdentificazionePersonale.Text = ""
    txtNumeroIdentificazioneTessera.Text = ""
    txtCodiceIstituzioneCompetente.Text = ""
    cboMese.ListIndex = Month(Now) - 1
    cboAnno.ListIndex = 0
    cboTipoRicetta.ListIndex = GetCboListIndex(2, cboTipoRicetta)
    oData(0).data = date
    flxGriglia.Rows = 1
    rsDataset.Open "SELECT ESENZIONE_REDDITO, T.CODICE, T.KEY FROM (PAZIENTI P INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=P.CODICE_ESENZIONE) WHERE P.KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        cboEsenzione.ListIndex = GetCboListIndex(rsDataset("KEY"), cboEsenzione)
        chkEsenteReddito.Value = IIf(CBool(rsDataset("ESENZIONE_REDDITO")), Checked, Unchecked)
    End If
    rsDataset.Close
    Call SettaProgressivi
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    ' pulisce per evitare problemi
    Timer1.Enabled = False
    lblImpegnabili.BackColor = &H8000000F
    lblImpegnabili.ToolTipText = ""
    
    If Index = 0 Then
        Call Pulisci
        tTrova.Tipo = tpPAZIENTE
    Else
        tTrova.Tipo = tpMEDICOBASE
    End If
    
    tTrova.condizione = ""
    tTrova.condStato = ""
    Unload frmTrova
    frmTrova.Show 1
    If Index = 0 Then
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
    Else
        If tTrova.keyReturn <> 0 Then
            intMedicoKey = tTrova.keyReturn
            Call CaricaMedico
        End If
    End If
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val(1 To 26) As Variant
    Dim v_Nomi(1 To 26) As Variant
    
    If Completo Then
        If Not verificaProgressivi Then
            Exit Sub
        End If
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "DATA_PRENOTAZIONE"
        v_Nomi(3) = "DATA_RICETTA"
        v_Nomi(4) = "PROGRESSIVO_ANNUALE"
        v_Nomi(5) = "PROGRESSIVO_RICETTA"
        v_Nomi(6) = "MAZZETTA1"
        v_Nomi(7) = "MAZZETTA2"
        v_Nomi(8) = "NUMERO_RICETTA"
        v_Nomi(9) = "STAMPATO_PC"
        v_Nomi(10) = "CODICE_PAZIENTE"
        v_Nomi(11) = "CODICE_ESENZIONE"
        v_Nomi(12) = "ESENTE_REDDITO"
        v_Nomi(13) = "TIPOLOGIA_RICETTA"
        v_Nomi(14) = "PRESENZA_BARCODE"
        v_Nomi(15) = "MESE"
        v_Nomi(16) = "ANNO"
        v_Nomi(17) = "VALIDATA"
        v_Nomi(18) = "FLAG"
        v_Nomi(19) = "CODICE_MEDICO"
        v_Nomi(20) = "ESENZIONE_DOPPIA"
        v_Nomi(21) = "CODICE_TIPO_EROGAZIONE"
        v_Nomi(22) = "TIPIRICETTAID"
        v_Nomi(23) = "CodiceIstituzioneCompetente"
        v_Nomi(24) = "NumeroIdentificativoPersonale"
        v_Nomi(25) = "NumeroIdentificazioneTessera"
        v_Nomi(26) = "DataScadenzaTessera"
                
        
        keyId = IIf(modifica, keyId, GetNumeroNuovo("RICETTE"))
        v_Val(1) = keyId
        v_Val(2) = oData(0).data
        v_Val(3) = oData(1).data
        v_Val(4) = txtProgressivoAnnuale
        v_Val(5) = txtProgressivoRicetta
        v_Val(6) = txtMazzettaPrimo
        v_Val(7) = txtMazzettaSecondo
        v_Val(8) = txtNumeroRicetta
        v_Val(9) = IIf(chkStampaPC.Value = Checked, True, False)
        v_Val(10) = intPazientiKey
        If cboEsenzione.ListIndex = -1 Then
            v_Val(11) = -1
        Else
            v_Val(11) = cboEsenzione.ItemData(cboEsenzione.ListIndex)
        End If
        v_Val(12) = IIf(chkEsenteReddito.Value = Checked, True, False)
        v_Val(13) = cboTipoPrescrizione.ListIndex
        v_Val(14) = IIf(chkPresenzaBarCode.Value = Checked, True, False)
        v_Val(15) = cboMese.ListIndex + 1
        v_Val(16) = cboAnno.Text
        v_Val(17) = False
        If modifica And cmdInvalidaNumeroRicetta.Visible = True Then
            v_Val(18) = 2
        Else
            v_Val(18) = 1
        End If
        v_Val(19) = intMedicoKey
        v_Val(20) = IIf(chkEsenzioneDoppia.Value = Checked, True, False)
        v_Val(21) = cboTipoErogazione.ItemData(cboTipoErogazione.ListIndex)
        If Not blnPazienteEstero Then
            v_Val(22) = 1
        Else
            v_Val(22) = cboTipoRicetta.ItemData(cboTipoRicetta.ListIndex)
        End If
        v_Val(23) = txtCodiceIstituzioneCompetente.Text
        v_Val(24) = txtNumeroIdentificazionePersonale.Text
        v_Val(25) = txtNumeroIdentificazioneTessera.Text
        v_Val(26) = IIf(oData(2).data = "", Null, oData(2).data)

        
        cnPrinc.BeginTrans
        Set rsRicette = New Recordset
        If modifica Then
            rsRicette.Open "SELECT * FROM RICETTE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsRicette.Update v_Nomi, v_Val
        Else
            rsRicette.Open "RICETTE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsRicette.AddNew v_Nomi, v_Val
            rsRicette.Update
        End If
        rsRicette.Close
        If Not SalvaPrescrizioni Then Exit Sub
        
        If keyIdVecchio <> 0 Then
            rsRicette.Open "SELECT * FROM RICETTE WHERE KEY=" & keyIdVecchio, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            rsRicette("FLAG") = 3
            rsRicette("VALIDATA") = False
            rsRicette.Update
            rsRicette.Close
        End If
        Set rsRicette = Nothing
        cnPrinc.CommitTrans
        
        Call Pulisci
        MsgBox "La ricetta è stata memorizzata", vbInformation, "Salvataggio"
        
        cmdTrova_Click (0)
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

'' Se cambia il mese ricarica le mazzette
Private Sub cboMese_Click()
    Timer1.Enabled = False
    lblImpegnabili.ToolTipText = ""
    lblImpegnabili.BackColor = &H8000000F
    lblImpegnabili.ForeColor = vbRed
    
    Dim mazzetta1 As Integer
    Dim mazzetta2 As Integer
    
    If stoPulendo Then Exit Sub
    Call CaricaMazzette(mazzetta1, mazzetta2)
    txtMazzettaPrimo = mazzetta1
    txtMazzettaSecondo = mazzetta2
    Call CalcolaTotaliSeduteDialisi
End Sub

Private Sub cboEsenzione_Click()
    If cboEsenzione.ListIndex = 0 Or cboEsenzione.Text = "E05" Then
        chkEsenzioneDoppia.Enabled = False
        chkEsenzioneDoppia.Value = Unchecked
    Else
        chkEsenzioneDoppia.Enabled = True
    End If
End Sub

'' Se cambia l'anno ricarica i progressivi
Private Sub cboAnno_Click()
    Timer1.Enabled = False
    lblImpegnabili.ToolTipText = ""
    lblImpegnabili.BackColor = &H8000000F
    lblImpegnabili.ForeColor = vbRed
    If stoCaricando Or stoPulendo Then Exit Sub
    Call SettaProgressivi
    Call CalcolaTotaliSeduteDialisi
End Sub

Private Sub cboPrescrizioni_DropDown()
    Call SetComboWidth(cboPrescrizioni, 500)
End Sub

Private Sub cboPrescrizioni_Click()
    cboPrescrizioni.Visible = False
End Sub

Private Sub cboPrescrizioni_LostFocus()
    Dim rsDataset As New Recordset
    If flxGriglia.TextMatrix(vRow, vCol) <> cboPrescrizioni.Text Then
        rsDataset.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE NOME='" & cboPrescrizioni.Text & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        flxGriglia.TextMatrix(flxGriglia.Row, 2) = cboPrescrizioni.Text
        flxGriglia.TextMatrix(flxGriglia.Row, 1) = rsDataset("CODICE")
        flxGriglia.TextMatrix(flxGriglia.Row, 4) = VirgolaOrPunto(Format(rsDataset("IMPORTO"), "####.00"), ",")
        flxGriglia.TextMatrix(flxGriglia.Row, 5) = VirgolaOrPunto(Format(flxGriglia.TextMatrix(flxGriglia.Row, 3) * CSng(rsDataset("IMPORTO")), "####.00"), ",")
        flxGriglia.TextMatrix(flxGriglia.Row, 8) = VirgolaOrPunto(Format(rsDataset("IMPORTO_SCONTATO"), "####.00"), ",")
    End If
    Set rsDataset = Nothing
    Call SalvaModifiche
    cboPrescrizioni.Visible = False
End Sub

Private Sub cboCodici_Click()
    cboCodici.Visible = False
End Sub

Private Sub cboCodici_LostFocus()
    Dim rsDataset As New Recordset
    If flxGriglia.TextMatrix(vRow, vCol) <> cboCodici.Text Then
        rsDataset.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO WHERE CODICE='" & cboCodici.Text & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        flxGriglia.TextMatrix(flxGriglia.Row, 1) = cboCodici.Text
        flxGriglia.TextMatrix(flxGriglia.Row, 2) = rsDataset("NOME")
        flxGriglia.TextMatrix(flxGriglia.Row, 4) = VirgolaOrPunto(Format(rsDataset("IMPORTO"), "####.00"), ",")
        flxGriglia.TextMatrix(flxGriglia.Row, 5) = VirgolaOrPunto(Format(flxGriglia.TextMatrix(flxGriglia.Row, 3) * CSng(rsDataset("IMPORTO")), "####.00"), ",")
        flxGriglia.TextMatrix(flxGriglia.Row, 8) = VirgolaOrPunto(Format(rsDataset("IMPORTO_SCONTATO"), "####.00"), ",")
    End If
    Set rsDataset = Nothing
    Call SalvaModifiche
    cboCodici.Visible = False
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    Dim i As Integer
    Dim corretto As Boolean
    
    If flxGriglia.Rows = 2 Then Exit Sub
    If oData(Index).data <> "" Then
        corretto = True
        If Index = 0 Then
            If oData(1).data <> "" Then
                If CDate(oData(0).data) < CDate(oData(1).data) Then
                    MsgBox "La data di prenotazione non può essere antecedente la data ricetta", vbCritical, "Attenzione"
                    corretto = False
                End If
            End If
            If flxGriglia.Rows <> 1 And corretto Then
                For i = 1 To flxGriglia.Rows - 1
                    If CDate(oData(0).data) >= flxGriglia.TextMatrix(i, 7) Then
                        MsgBox "La data di prenotazione non può essere successiva o uguale la data di fine prestazione", vbCritical, "Attenzione"
                        corretto = False
                    End If
                Next i
            End If
            If Not corretto Then oData(0).Pulisci
        ElseIf Index = 1 Then
            If oData(0).data <> "" Then
                If CDate(oData(1).data) > CDate(oData(0).data) Then
                    MsgBox "La data di ricetta non può essere successiva la data prenotazione", vbCritical, "Attenzione"
                    corretto = False
                End If
            End If
            If flxGriglia.Rows <> 1 And corretto Then
                For i = 1 To flxGriglia.Rows - 1
                    If CDate(oData(1).data) >= flxGriglia.TextMatrix(i, 7) Then
                        MsgBox "La data di ricetta non può essere successiva o uguale la data di fine prestazione", vbCritical, "Attenzione"
                        corretto = False
                    End If
                Next i
            End If
            If Not corretto Then oData(1).Pulisci
        End If
    End If
End Sub

'' Carica i dati del medico
Private Sub CaricaMedico()
    Dim rsDataset As Recordset
    If intMedicoKey = 0 Then Exit Sub
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME, CODICE FROM MEDICI_BASE WHERE KEY=" & intMedicoKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognomeMedico = rsDataset("COGNOME") & ""
    lblNomeMedico = rsDataset("NOME") & ""
    lblCodiceTimbroMedico = rsDataset("CODICE") & ""
    Set rsDataset = Nothing
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    Dim nome As String
    Dim strSql As String
    
    If intPazientiKey = 0 Then
        Exit Sub
    End If
    Set rsDataset = New Recordset
    
    strSql = "SELECT    PAZIENTI.COGNOME, PAZIENTI.NOME, DATA_NASCITA, CODICE_FISCALE, CODICE_MEDICO, CODICE_REGIONE, CODICE_DISTRETTO, ESENZIONE_REDDITO, CODICE_ESENZIONE, Nazioni.Nome as NazioniNome, " & _
             "          MEDICI_BASE.KEY AS MEDICI_BASEKEY, MEDICI_BASE.CODICE AS MEDICI_BASECODICE, CODICE_TIPO_MEDICO, PRESENZA_BARCODE  " & _
             "FROM      ((PAZIENTI " & _
             "          LEFT OUTER JOIN MEDICI_BASE ON MEDICI_BASE.KEY=PAZIENTI.CODICE_MEDICO) " & _
             "          LEFT OUTER JOIN NAZIONI ON NAZIONI.KEY=PAZIENTI.NAZIONIID) " & _
             "WHERE     PAZIENTI.KEY=" & intPazientiKey
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    ControlloCodiceFiscalePazienteEstero = rsDataset("CODICE_FISCALE")
    
    If rsDataset("CODICE_MEDICO") = 0 Then
        nome = "Al paziente non è stato definito il campo MEDICO DI BASE"
    ElseIf rsDataset("MEDICI_BASECODICE") = "" Or IsNull(rsDataset("MEDICI_BASECODICE")) Then
        nome = "Al paziente non è stato definito il campo CODICE REGIONALE DEL MEDICO DI BASE"
    ElseIf rsDataset("CODICE_REGIONE") = 16 And rsDataset("CODICE_DISTRETTO") = 0 Then
        nome = "Al paziente non è stato definito il campo DISTRETTO"
    ElseIf rsDataset("CODICE_TIPO_MEDICO") = -1 And structIntestazione.sCodiceAsl = 1 Then
        nome = "Non è stata definita la TIPOLOGIA del medico associato al paziente"
    End If
    If nome <> "" Then
        MsgBox "IMPOSSIBILE PROSEGUIRE NELLA COMPILAZIONE!!!" & vbCrLf & nome, vbCritical, "Attenzione"
        cmdTrova_Click (0)
        Exit Sub
    End If
    If UCase(rsDataset("NAZIONINOME")) = UCase("iTALIA") Then
        blnPazienteEstero = False
        fraPazientiEsteri.Visible = False
        fraPazientiEsteri.Enabled = False
        fraPrestazioni.Top = fraPazientiEsteri.Top
        fraPulsanti.Top = fraPrestazioni.Top + fraPrestazioni.Height - 100
        Me.Height = fraPulsanti.Top + fraPulsanti.Height + 400
    Else
        blnPazienteEstero = True
        fraPazientiEsteri.Visible = True
        fraPazientiEsteri.Enabled = True
        fraPrestazioni.Top = fraPazientiEsteri.Top + fraPazientiEsteri.Height - 20
        fraPrestazioni.ZOrder 1
        fraPulsanti.Top = fraPrestazioni.Top + fraPrestazioni.Height - 100
        fraPulsanti.ZOrder 1
        Me.Height = fraPulsanti.Top + fraPulsanti.Height + 400
    End If
    
    cboEsenzione.ListIndex = GetCboListIndex(rsDataset("CODICE_ESENZIONE"), cboEsenzione)
    chkEsenteReddito.Value = IIf(CBool(rsDataset("ESENZIONE_REDDITO")), Checked, Unchecked)
    chkPresenzaBarCode.Value = IIf(CBool(rsDataset("PRESENZA_BARCODE")), Checked, Unchecked)
    intMedicoKey = rsDataset("MEDICI_BASEKEY")
    Call CaricaMedico
    codiceDistretto = rsDataset("CODICE_DISTRETTO")
    Call SettaProgressivi
    Call CalcolaTotaliSeduteDialisi
    
    txtNumeroRicetta.SetFocus
    Set rsDataset = Nothing
End Sub

Private Sub txtAppo_Change()
    If Not lettera = "" Then
        Call OnlyNumber(txtAppo, lettera)
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim OldImpegnabili As Integer
        If txtAppo <> "" Then
        If txtAppo > 0 And txtAppo < 99 Then
            OldImpegnabili = flxGriglia.TextMatrix(vRow, 3)
            flxGriglia.TextMatrix(vRow, 3) = txtAppo.Text
            ' ricalcola il totale complessivo
            flxGriglia.TextMatrix(flxGriglia.Row, 5) = Format(flxGriglia.TextMatrix(vRow, 3) * CSng(VirgolaOrPunto(flxGriglia.TextMatrix(vRow, 4), ".")), "####.00")
            MImpegnabili = MImpegnabili + txtAppo.Text - OldImpegnabili
            Call CalcolaTotaliSeduteDialisi
            Call SalvaModifiche
        Else
            MsgBox "Inserire un valore compreso tra 1 e 99", vbCritical, "Attenzione"
        End If
    End If
    txtAppo.Visible = False
End Sub

Private Sub txtMazzettaPrimo_GotFocus()
    txtMazzettaPrimo.BackColor = colArancione
End Sub

Private Sub txtMazzettaPrimo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMazzettaPrimo_LostFocus()
    txtMazzettaPrimo.BackColor = vbWhite
End Sub

Private Sub txtMazzettaSecondo_GotFocus()
    txtMazzettaSecondo.BackColor = colArancione
End Sub

Private Sub txtMazzettaSecondo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMazzettaSecondo_LostFocus()
    txtMazzettaSecondo.BackColor = vbWhite
End Sub

Private Sub txtNumeroRicetta_GotFocus()
    txtNumeroRicetta.BackColor = colArancione
End Sub

Private Sub txtNumeroRicetta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumeroRicetta_LostFocus()
    txtNumeroRicetta.BackColor = vbWhite
End Sub

Private Sub txtProgressivoAnnuale_GotFocus()
    txtProgressivoAnnuale.BackColor = colArancione
End Sub

Private Sub txtProgressivoAnnuale_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtProgressivoAnnuale_LostFocus()
    txtProgressivoAnnuale.BackColor = vbWhite
End Sub

Private Sub txtProgressivoRicetta_GotFocus()
    txtProgressivoRicetta.BackColor = colArancione
End Sub

Private Sub txtProgressivoRicetta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtProgressivoRicetta_LostFocus()
    txtProgressivoRicetta.BackColor = vbWhite
End Sub

