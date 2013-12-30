VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmSchedeSorveglianzaFAV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schede Sorveglianza FAV"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPortataRicircolo 
      Caption         =   "Valutazione Portate e Ricircolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   47
      Top             =   4680
      Width           =   12855
      Begin VB.TextBox txtPortataIndicatori 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   53
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtRicircoloIndicatori 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   52
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtPortataParametri 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   51
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtRicircoloParametri 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   50
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtPortataTollAccettate 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   49
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtRicircoloTollAccettate 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   48
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Toll. Accettate"
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
         Left            =   1320
         TabIndex        =   55
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Portata:"
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
         Index           =   23
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ricircolo:"
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
         Index           =   22
         Left            =   6960
         TabIndex        =   60
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indicatori"
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
         Index           =   21
         Left            =   1800
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indicatori"
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
         Index           =   20
         Left            =   8160
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parametri"
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
         Left            =   1800
         TabIndex        =   57
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parametri"
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
         Left            =   8160
         TabIndex        =   56
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Toll. Accettate"
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
         Left            =   7680
         TabIndex        =   54
         Top             =   1080
         Width           =   1515
      End
   End
   Begin VB.Frame fraRilevazione 
      Caption         =   "Rilevazione Pressione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Width           =   12855
      Begin VB.TextBox txtRientroTollAccettate 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   46
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtAspirazioneTollAccettate 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   45
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtRientroParametri 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   42
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtAspirazioneParametri 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   41
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtRientroIndicatore 
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
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   38
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtAspirazioneIndicatore 
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
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   37
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Toll. Accettate"
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
         Left            =   7680
         TabIndex        =   44
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Toll. Accettate"
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
         Left            =   1320
         TabIndex        =   43
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parametri"
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
         Left            =   8160
         TabIndex        =   40
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parametri"
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
         Left            =   1800
         TabIndex        =   39
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indicatori"
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
         Left            =   8160
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indicatori"
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
         Left            =   1800
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In Rientro:"
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
         Left            =   6960
         TabIndex        =   34
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In Aspirazione:"
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
         TabIndex        =   33
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2415
      Left            =   120
      TabIndex        =   77
      Top             =   840
      Width           =   3735
      Begin VB.Frame Frame9 
         Height          =   855
         Left            =   120
         TabIndex        =   81
         Top             =   1320
         Width           =   735
         Begin VB.OptionButton optNoAccessoVascolare 
            Caption         =   "No"
            Height          =   255
            Left            =   0
            TabIndex        =   83
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton optSiAccessoVascolare 
            Caption         =   "Si"
            Height          =   255
            Left            =   0
            TabIndex        =   82
            Top             =   120
            Width           =   495
         End
      End
      Begin DataTimeBox.uDataTimeBox oDataNuovoAccessoVascolare 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   78
         Top             =   1380
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label lblNomeCognomeUtente 
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   85
         Top             =   480
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Si è reso necessario eseguire un nuovo accesso vascolare?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   28
         Left            =   120
         TabIndex        =   84
         Top             =   840
         Width           =   3465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipUtente 
         Height          =   255
         Index           =   27
         Left            =   840
         TabIndex        =   80
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Utente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   79
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame frmSegni 
      Caption         =   "Segni e Sintomi locali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   9135
      Begin VB.Frame Frame7 
         Height          =   495
         Left            =   2280
         TabIndex        =   66
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton optSiEritema 
            Caption         =   "Si"
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
            Left            =   1320
            TabIndex        =   68
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optNoEritema 
            Caption         =   "No"
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
            TabIndex        =   67
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   2280
         TabIndex        =   65
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton optSiDolore 
            Caption         =   "Si"
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
            Left            =   1320
            TabIndex        =   70
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optNoDolore 
            Caption         =   "No"
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
            TabIndex        =   69
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   2280
         TabIndex        =   64
         Top             =   1680
         Width           =   2175
         Begin VB.OptionButton optSiPresenzaFremiti 
            Caption         =   "Si"
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
            Left            =   1320
            TabIndex        =   76
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optNoPresenzaFremiti 
            Caption         =   "No"
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
            TabIndex        =   75
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   2280
         TabIndex        =   63
         Top             =   1320
         Width           =   2175
         Begin VB.OptionButton optSiInfiltrazione 
            Caption         =   "Si"
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
            Left            =   1320
            TabIndex        =   74
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optNoInfiltrazione 
            Caption         =   "No"
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
            TabIndex        =   73
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   2280
         TabIndex        =   62
         Top             =   960
         Width           =   2175
         Begin VB.OptionButton optSiGonfiore 
            Caption         =   "Si"
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
            Left            =   1320
            TabIndex        =   72
            Top             =   120
            Width           =   735
         End
         Begin VB.OptionButton optNoGonfiore 
            Caption         =   "No"
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
            TabIndex        =   71
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CheckBox chkEritemaMedio 
         Caption         =   "Medio"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkPresenzaFremitiGrave 
         Caption         =   "Grave"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkInfiltrazioneGrave 
         Caption         =   "Grave"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkGonfioreGrave 
         Caption         =   "Grave"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDoloreGrave 
         Caption         =   "Grave"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkEritemaGrave 
         Caption         =   "Grave"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkPresenzaFremitiMedio 
         Caption         =   "Medio"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkInfiltrazioneMedio 
         Caption         =   "Medio"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkGonfioreMedio 
         Caption         =   "Medio"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDoloreMedio 
         Caption         =   "Medio"
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
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkPresenzaFremitiLieve 
         Caption         =   "Lieve"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkInfiltrazioneLieve 
         Caption         =   "Lieve"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkGonfioreLieve 
         Caption         =   "Lieve"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkDoloreLieve 
         Caption         =   "Lieve"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkEritemaLieve 
         Caption         =   "Lieve"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eritema"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dolore"
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
         TabIndex        =   15
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gonfiore"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Infiltrazione"
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
         TabIndex        =   13
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Presenza fremiti"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmSchedeSorveglianzaIAV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   450
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
         Left            =   11280
         TabIndex        =   7
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
         Left            =   6480
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   1005
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
         Left            =   11880
         TabIndex        =   4
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
         Left            =   7200
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
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   6360
      Width           =   7815
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
         CausesValidation=   0   'False
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
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
         CausesValidation=   0   'False
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
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSchedeSorveglianzaFAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDataset As Recordset
Dim PazienteKey As Integer
'Dim ColRosso As Long
'Dim ColNero As Long
'ColRosso = &HFF&
'ColNero = &H80000012
Dim keyId As Integer
Dim modifica As Boolean

Private Sub chkDoloreGrave_GotFocus()
    chkDoloreLieve.Value = Unchecked
    chkDoloreMedio.Value = Unchecked
End Sub

Private Sub chkDoloreLieve_GotFocus()
    chkDoloreMedio.Value = Unchecked
    chkDoloreGrave.Value = Unchecked
End Sub

Private Sub chkDoloreMedio_GotFocus()
    chkDoloreLieve.Value = Unchecked
    chkDoloreGrave.Value = Unchecked
End Sub

Private Sub chkEritemaGrave_GotFocus()
    chkEritemaLieve.Value = Unchecked
    chkEritemaMedio.Value = Unchecked
End Sub

Private Sub chkEritemaLieve_GotFocus()
    chkEritemaMedio.Value = Unchecked
    chkEritemaGrave.Value = Unchecked
End Sub

Private Sub chkEritemaMedio_GotFocus()
    chkEritemaLieve.Value = Unchecked
    chkEritemaGrave.Value = Unchecked
End Sub

Private Sub chkGonfioreGrave_GotFocus()
    chkGonfioreLieve.Value = Unchecked
    chkGonfioreMedio.Value = Unchecked
End Sub

Private Sub chkGonfioreLieve_GotFocus()
    chkGonfioreMedio.Value = Unchecked
    chkGonfioreGrave.Value = Unchecked
End Sub

Private Sub chkGonfioreMedio_Click()
    chkGonfioreLieve.Value = Unchecked
    chkGonfioreGrave.Value = Unchecked
End Sub


Private Sub chkInfiltrazioneGrave_GotFocus()
    chkInfiltrazioneLieve.Value = Unchecked
    chkInfiltrazioneMedio.Value = Unchecked
End Sub



Private Sub chkInfiltrazioneLieve_GotFocus()
    chkInfiltrazioneMedio.Value = Unchecked
    chkInfiltrazioneGrave.Value = Unchecked
End Sub



Private Sub chkInfiltrazioneMedio_GotFocus()
    chkInfiltrazioneLieve.Value = Unchecked
    chkInfiltrazioneGrave.Value = Unchecked
End Sub

Private Sub chkPresenzaFremitiGrave_GotFocus()
    chkPresenzaFremitiLieve.Value = Unchecked
    chkPresenzaFremitiMedio.Value = Unchecked
End Sub

Private Sub chkPresenzaFremitiLieve_GotFocus()
    chkPresenzaFremitiMedio.Value = Unchecked
    chkPresenzaFremitiGrave.Value = Unchecked
End Sub


Private Sub chkPresenzaFremitiMedio_GotFocus()
    chkPresenzaFremitiLieve.Value = Unchecked
    chkPresenzaFremitiGrave.Value = Unchecked
End Sub

Private Sub cmdChiudi_Click()
    Unload frmSchedeSorveglianzaFAV
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant

    
    If Not modifica Then
        keyId = GetNumero("SCHEDA_SORV_FAV")
    End If
    
    v_Nomi = Array("KEY", "KEY_PAZIENTE", "ERI_SI_NO", "ERI_VALORE", _
            "DOL_SI_NO", "DOL_VALORE", _
            "GON_SI_NO", "GON_VALORE", _
            "INF_SI_NO", "INF_VALORE", _
            "PRE_FRE_SI_NO", "PRE_FRE_VALORE", _
            "ASP_INDICATORI", "ASP_PARAMETRI", "ASP_TOLL_ACCET", _
            "RIE_INDICATORI", "RIE_PARAMETRI", "RIE_TOLL_ACCET", _
            "POR_INDICATORI", "POR_PARAMETRI", "P0R_TOLL_ACCET", _
            "RIC_INDICATORI", "RIC_PARAMETRI", "RIC_TOLL_ACCET")

    v_Val = Array(keyId, PazienteKey, GestisciSiNoEritema, GestisciOptEritema, _
            GestisciSiNoDolore, GestisciOptDolore, _
            GestisciSiNoGonfiore, GestisciOptGonfiore, _
            GestisciSiNoInfiltrazione, GestisciOptInfiltrazione, _
            GestisciSiNoPresenzaFremiti, GestisciOptPresenzaFremiti, _
            txtAspirazioneIndicatore, txtAspirazioneParametri, txtAspirazioneTollAccettate, _
            txtRientroIndicatore, txtRientroParametri, txtRientroTollAccettate, _
            txtPortataIndicatori, txtPortataParametri, txtPortataTollAccettate, _
            txtRicircoloIndicatori, txtRicircoloParametri, txtRicircoloTollAccettate)
        
    Set rsDataset = New Recordset
        If modifica = False Then
            rsDataset.Open "SCHEDA_SORV_FAV", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.AddNew v_Nomi, v_Val
            rsDataset.Update
            modifica = True
        Else
            rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDataset.Update v_Nomi, v_Val
        End If
    Set rsDataset = Nothing

    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    
End Sub

Private Function GestisciSiNoPresenzaFremiti() As String
    If optNoPresenzaFremiti.Value = True Then
        GestisciSiNoPresenzaFremiti = "NO"
    Else
        GestisciSiNoPresenzaFremiti = "SI"
    End If
End Function

Private Function GestisciOptPresenzaFremiti() As String
    If chkPresenzaFremitiLieve.Value = Checked Then
        GestisciOptPresenzaFremiti = "LIEVE"
    ElseIf chkPresenzaFremitiMedio.Value = Checked Then
        GestisciOptPresenzaFremiti = "MEDIO"
    ElseIf chkPresenzaFremitiGrave.Value = Checked Then
        GestisciOptPresenzaFremiti = "GRAVE"
    End If
End Function

Private Function GestisciSiNoInfiltrazione() As String
    If optNoInfiltrazione.Value = True Then
        GestisciSiNoInfiltrazione = "NO"
    Else
        GestisciSiNoInfiltrazione = "SI"
    End If
End Function

Private Function GestisciOptInfiltrazione() As String
    If chkInfiltrazioneLieve.Value = Checked Then
        GestisciOptInfiltrazione = "LIEVE"
    ElseIf chkInfiltrazioneMedio.Value = Checked Then
        GestisciOptInfiltrazione = "MEDIO"
    ElseIf chkInfiltrazioneGrave.Value = Checked Then
        GestisciOptInfiltrazione = "GRAVE"
    End If
End Function

Private Function GestisciSiNoGonfiore() As String
    If optNoGonfiore.Value = True Then
        GestisciSiNoGonfiore = "NO"
    Else
        GestisciSiNoGonfiore = "SI"
    End If
End Function

Private Function GestisciOptGonfiore() As String
    If chkGonfioreLieve.Value = Checked Then
        GestisciOptGonfiore = "LIEVE"
    ElseIf chkGonfioreMedio.Value = Checked Then
        GestisciOptGonfiore = "MEDIO"
    ElseIf chkGonfioreGrave.Value = Checked Then
        GestisciOptGonfiore = "GRAVE"
    End If
End Function

Private Function GestisciSiNoDolore() As String
    If optNoDolore.Value = True Then
        GestisciSiNoDolore = "NO"
    Else
        GestisciSiNoDolore = "SI"
    End If
End Function

Private Function GestisciOptDolore() As String
    If chkDoloreLieve.Value = Checked Then
        GestisciOptDolore = "LIEVE"
    ElseIf chkDoloreMedio.Value = Checked Then
        GestisciOptDolore = "MEDIO"
    ElseIf chkDoloreGrave.Value = Checked Then
        GestisciOptDolore = "GRAVE"
    End If
End Function

Private Function GestisciSiNoEritema() As String
    If optNoEritema.Value = True Then
        GestisciSiNoEritema = "NO"
    Else
        GestisciSiNoEritema = "SI"
    End If
End Function

Private Function GestisciOptEritema() As String
    If chkEritemaLieve.Value = Checked Then
        GestisciOptEritema = "LIEVE"
    ElseIf chkEritemaMedio.Value = Checked Then
        GestisciOptEritema = "MEDIO"
    ElseIf chkEritemaGrave.Value = Checked Then
        GestisciOptEritema = "GRAVE"
    End If
End Function

Private Sub Form_Activate()
    If PazienteKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Pulisci()
    'Eritema
    optNoEritema.Value = False
    optSiEritema.Value = False
    chkEritemaLieve.Value = Unchecked
    chkEritemaMedio.Value = Unchecked
    chkEritemaGrave.Value = Unchecked
    chkEritemaLieve.Enabled = False
    chkEritemaMedio.Enabled = False
    chkEritemaGrave.Enabled = False
    
    'Dolore
    optNoDolore.Value = False
    optSiDolore.Value = False
    chkDoloreLieve.Value = Unchecked
    chkDoloreMedio.Value = Unchecked
    chkDoloreGrave.Value = Unchecked
    chkDoloreLieve.Enabled = False
    chkDoloreMedio.Enabled = False
    chkDoloreGrave.Enabled = False
    
    'Gonfiore
    optNoGonfiore.Value = False
    optSiGonfiore.Value = False
    chkGonfioreLieve.Value = Unchecked
    chkGonfioreMedio.Value = Unchecked
    chkGonfioreGrave.Value = Unchecked
    chkGonfioreLieve.Enabled = False
    chkGonfioreMedio.Enabled = False
    chkGonfioreGrave.Enabled = False
    
    'Infiltrazione
    optNoInfiltrazione.Value = False
    optSiInfiltrazione.Value = False
    chkInfiltrazioneLieve.Value = Unchecked
    chkInfiltrazioneMedio.Value = Unchecked
    chkInfiltrazioneGrave.Value = Unchecked
    chkInfiltrazioneLieve.Enabled = False
    chkInfiltrazioneMedio.Enabled = False
    chkInfiltrazioneGrave.Enabled = False
    
    'Presenza fremiti
    optNoPresenzaFremiti.Value = False
    optSiPresenzaFremiti.Value = False
    chkPresenzaFremitiLieve.Value = Unchecked
    chkPresenzaFremitiMedio.Value = Unchecked
    chkPresenzaFremitiGrave.Value = Unchecked
    chkPresenzaFremitiLieve.Enabled = False
    chkPresenzaFremitiMedio.Enabled = False
    chkPresenzaFremitiGrave.Enabled = False
    
    'Rilevazioni
    txtAspirazioneIndicatore.Text = ""
    txtAspirazioneParametri.Text = ""
    txtAspirazioneTollAccettate.Text = ""
    txtRientroIndicatore.Text = ""
    txtRientroParametri.Text = ""
    txtRientroTollAccettate.Text = ""
    
    'Portate e Ricircolo
    txtPortataIndicatori.Text = ""
    txtPortataParametri.Text = ""
    txtPortataTollAccettate.Text = ""
    txtRicircoloIndicatori.Text = ""
    txtRicircoloParametri.Text = ""
    txtRicircoloTollAccettate.Text = ""
End Sub

Private Sub cmdTrova_Click()
    Call Pulisci
    lblTipUtente(27).Caption = Choose(tAccesso.Tipo, "Medico", "Infermiere", "Contabile", "Amministratore")
    lblNomeCognomeUtente(0).Caption = tAccesso.cognome & " " & tAccesso.nome
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        If PazienteKey = tTrova.keyReturn Then
            PazienteKey = 0
            Call CaricaPaziente
            PazienteKey = tTrova.keyReturn
            Call CaricaPaziente
        Else
            PazienteKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub CaricaPaziente()
    
    If PazienteKey = 0 Then
    
    Else
        ' carica i dati del paziente
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & PazienteKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
       
        ' cerca i riferimenti al paziente
        Call CaricaValori
    
    End If
End Sub

Private Sub CaricaValori()
    
    Set rsDataset = New Recordset
    
    rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        Call CaricaSiNoEritema
        Call CaricaSiNoDolore
        Call CaricaSiNoGonfiore
        Call CaricaSiNoInfiltrazione
        Call CaricaSiNoPresenzaFremito
        Call CaricaValoreEritema
        Call CaricaValoreDolore
        Call CaricaValoreGonfiore
        Call CaricaValoreInfiltrazione
        Call CaricaValorePresenzaFremiti
        
        Call CaricaRilevazionePressione
        Call CaricaPortataRicircolo
        modifica = True
    Else
        'Se non trova il paziente associato vuol dire che è in fase di inserimento
        modifica = False
    End If
    
    Set rsDataset = Nothing
    
End Sub

Private Sub CaricaPortataRicircolo()
    txtPortataIndicatori.Text = rsDataset("POR_INDICATORI") & ""
    txtPortataParametri.Text = rsDataset("POR_PARAMETRI") & ""
    txtPortataTollAccettate.Text = rsDataset("P0R_TOLL_ACCET") & ""
    txtRicircoloIndicatori.Text = rsDataset("RIC_INDICATORI") & ""
    txtRicircoloParametri.Text = rsDataset("RIC_PARAMETRI") & ""
    txtRicircoloTollAccettate.Text = rsDataset("RIC_TOLL_ACCET") & ""
End Sub

Private Sub CaricaRilevazionePressione()
    txtAspirazioneIndicatore.Text = rsDataset("ASP_INDICATORI") & ""
    txtAspirazioneParametri.Text = rsDataset("ASP_PARAMETRI") & ""
    txtAspirazioneTollAccettate.Text = rsDataset("ASP_TOLL_ACCET") & ""
    txtRientroIndicatore.Text = rsDataset("RIE_INDICATORI") & ""
    txtRientroParametri.Text = rsDataset("RIE_PARAMETRI") & ""
    txtRientroTollAccettate.Text = rsDataset("RIE_TOLL_ACCET") & ""
End Sub

Private Sub CaricaValoreEritema()
    If rsDataset("ERI_VALORE") = "LIEVE" Then
        chkEritemaLieve.Value = Checked
    ElseIf rsDataset("ERI_VALORE") = "MEDIO" Then
        chkEritemaMedio.Value = Checked
    ElseIf rsDataset("ERI_VALORE") = "GRAVE" Then
        chkEritemaGrave.Value = Checked
    End If
End Sub

Private Sub CaricaValoreDolore()
    If rsDataset("DOL_VALORE") = "LIEVE" Then
        chkDoloreLieve.Value = Checked
    ElseIf rsDataset("DOL_VALORE") = "MEDIO" Then
        chkDoloreMedio.Value = Checked
    ElseIf rsDataset("DOL_VALORE") = "GRAVE" Then
        chkDoloreGrave.Value = Checked
    End If
End Sub

Private Sub CaricaValoreGonfiore()
    If rsDataset("GON_VALORE") = "LIEVE" Then
        chkGonfioreLieve.Value = Checked
    ElseIf rsDataset("GON_VALORE") = "MEDIO" Then
        chkGonfioreMedio.Value = Checked
    ElseIf rsDataset("GON_VALORE") = "GRAVE" Then
        chkGonfioreGrave.Value = Checked
    End If
End Sub

Private Sub CaricaValoreInfiltrazione()
    If rsDataset("INF_VALORE") = "LIEVE" Then
        chkInfiltrazioneLieve.Value = Checked
    ElseIf rsDataset("INF_VALORE") = "MEDIO" Then
        chkInfiltrazioneMedio.Value = Checked
    ElseIf rsDataset("INF_VALORE") = "GRAVE" Then
        chkInfiltrazioneGrave.Value = Checked
    End If
End Sub

Private Sub CaricaValorePresenzaFremiti()
    If rsDataset("PRE_FRE_VALORE") = "LIEVE" Then
        chkPresenzaFremitiLieve.Value = Checked
    ElseIf rsDataset("PRE_FRE_VALORE") = "MEDIO" Then
        chkPresenzaFremitiMedio.Value = Checked
    ElseIf rsDataset("PRE_FRE_VALORE") = "GRAVE" Then
        chkPresenzaFremitiGrave.Value = Checked
    End If
End Sub

Private Sub CaricaSiNoEritema()
    If rsDataset("ERI_SI_NO") = "NO" Then
        optNoEritema.Value = True
    ElseIf rsDataset("ERI_SI_NO") = "SI" Then
        optSiEritema.Value = True
        chkEritemaLieve.Enabled = True
        chkEritemaMedio.Enabled = True
        chkEritemaGrave.Enabled = True
    End If
End Sub

Private Sub CaricaSiNoDolore()
    If rsDataset("DOL_SI_NO") = "NO" Then
        optNoDolore.Value = True
    ElseIf rsDataset("DOL_SI_NO") = "SI" Then
        optSiDolore.Value = True
        chkDoloreLieve.Enabled = True
        chkDoloreMedio.Enabled = True
        chkDoloreGrave.Enabled = True
    End If
End Sub

Private Sub CaricaSiNoGonfiore()
    If rsDataset("GON_SI_NO") = "NO" Then
        optNoGonfiore.Value = True
    ElseIf rsDataset("GON_SI_NO") = "SI" Then
        optSiGonfiore.Value = True
        chkGonfioreLieve.Enabled = True
        chkGonfioreMedio.Enabled = True
        chkGonfioreGrave.Enabled = True
    End If
End Sub

Private Sub CaricaSiNoInfiltrazione()
    If rsDataset("INF_SI_NO") = "NO" Then
        optNoInfiltrazione.Value = True
    ElseIf rsDataset("INF_SI_NO") = "SI" Then
        optSiInfiltrazione.Value = True
        chkInfiltrazioneLieve.Enabled = True
        chkInfiltrazioneMedio.Enabled = True
        chkInfiltrazioneGrave.Enabled = True
    End If
End Sub

Private Sub CaricaSiNoPresenzaFremito()
    If rsDataset("PRE_FRE_SI_NO") = "NO" Then
        optNoPresenzaFremiti.Value = True
    ElseIf rsDataset("PRE_FRE_SI_NO") = "SI" Then
        optSiPresenzaFremiti.Value = True
        chkPresenzaFremitiLieve.Enabled = True
        chkPresenzaFremitiMedio.Enabled = True
        chkPresenzaFremitiGrave.Enabled = True
    End If
End Sub

Private Sub oDataNuovoAccessoVascolare_OnDataClick(Index As Integer)
    oDataNuovoAccessoVascolare(2).Pulisci
End Sub

Private Sub optNoDolore_GotFocus()
    chkDoloreLieve.Enabled = False
    chkDoloreMedio.Enabled = False
    chkDoloreGrave.Enabled = False
    chkDoloreLieve.Value = Unchecked
    chkDoloreMedio.Value = Unchecked
    chkDoloreGrave.Value = Unchecked
End Sub

Private Sub optNoEritema_GotFocus()
    chkEritemaLieve.Enabled = False
    chkEritemaMedio.Enabled = False
    chkEritemaGrave.Enabled = False
    chkEritemaLieve.Value = Unchecked
    chkEritemaMedio.Value = Unchecked
    chkEritemaGrave.Value = Unchecked
End Sub

Private Sub optNoGonfiore_GotFocus()
    chkGonfioreLieve.Enabled = False
    chkGonfioreMedio.Enabled = False
    chkGonfioreGrave.Enabled = False
    chkGonfioreLieve.Value = Unchecked
    chkGonfioreMedio.Value = Unchecked
    chkGonfioreGrave.Value = Unchecked
End Sub

Private Sub optNoInfiltrazione_GotFocus()
    chkInfiltrazioneLieve.Enabled = False
    chkInfiltrazioneMedio.Enabled = False
    chkInfiltrazioneGrave.Enabled = False
    chkInfiltrazioneLieve.Value = Unchecked
    chkInfiltrazioneMedio.Value = Unchecked
    chkInfiltrazioneGrave.Value = Unchecked
End Sub

Private Sub optNoPresenzaFremiti_GotFocus()
    chkPresenzaFremitiLieve.Enabled = False
    chkPresenzaFremitiMedio.Enabled = False
    chkPresenzaFremitiGrave.Enabled = False
    chkPresenzaFremitiLieve.Value = Unchecked
    chkPresenzaFremitiMedio.Value = Unchecked
    chkPresenzaFremitiGrave.Value = Unchecked
End Sub

Private Sub optSiDolore_GotFocus()
    chkDoloreLieve.Enabled = True
    chkDoloreMedio.Enabled = True
    chkDoloreGrave.Enabled = True
End Sub

Private Sub optSiEritema_GotFocus()
    chkEritemaLieve.Enabled = True
    chkEritemaMedio.Enabled = True
    chkEritemaGrave.Enabled = True
End Sub

Private Sub optSiGonfiore_GotFocus()
    chkGonfioreLieve.Enabled = True
    chkGonfioreMedio.Enabled = True
    chkGonfioreGrave.Enabled = True
End Sub

Private Sub optSiInfiltrazione_GotFocus()
    chkInfiltrazioneLieve.Enabled = True
    chkInfiltrazioneMedio.Enabled = True
    chkInfiltrazioneGrave.Enabled = True
End Sub

Private Sub optSiPresenzaFremiti_GotFocus()
    chkPresenzaFremitiLieve.Enabled = True
    chkPresenzaFremitiMedio.Enabled = True
    chkPresenzaFremitiGrave.Enabled = True
End Sub

Private Sub txtAspirazioneIndicatore_GotFocus()
    txtAspirazioneIndicatore.BackColor = colArancione
End Sub

Private Sub txtAspirazioneIndicatore_LostFocus()
    txtAspirazioneIndicatore.BackColor = vbWhite
End Sub

Private Sub txtAspirazioneParametri_GotFocus()
    txtAspirazioneParametri.BackColor = colArancione
End Sub

Private Sub txtAspirazioneParametri_LostFocus()
    txtAspirazioneParametri.BackColor = vbWhite
End Sub

Private Sub txtAspirazioneTollAccettate_GotFocus()
    txtAspirazioneTollAccettate.BackColor = colArancione
End Sub

Private Sub txtAspirazioneTollAccettate_LostFocus()
    txtAspirazioneTollAccettate.BackColor = vbWhite
End Sub

Private Sub txtPortataIndicatori_GotFocus()
    txtPortataIndicatori.BackColor = colArancione
End Sub

Private Sub txtPortataIndicatori_LostFocus()
    txtPortataIndicatori.BackColor = vbWhite
End Sub

Private Sub txtPortataParametri_GotFocus()
    txtPortataParametri.BackColor = colArancione
End Sub

Private Sub txtPortataParametri_LostFocus()
    txtPortataParametri.BackColor = vbWhite
End Sub

Private Sub txtPortataTollAccettate_GotFocus()
    txtPortataTollAccettate.BackColor = colArancione
End Sub

Private Sub txtPortataTollAccettate_LostFocus()
     txtPortataTollAccettate.BackColor = vbWhite
End Sub

Private Sub txtRicircoloIndicatori_GotFocus()
    txtRicircoloIndicatori.BackColor = colArancione
End Sub

Private Sub txtRicircoloIndicatori_LostFocus()
    txtRicircoloIndicatori.BackColor = vbWhite
End Sub

Private Sub txtRicircoloParametri_GotFocus()
    txtRicircoloParametri.BackColor = colArancione
End Sub

Private Sub txtRicircoloParametri_LostFocus()
    txtRicircoloParametri.BackColor = vbWhite
End Sub

Private Sub txtRicircoloTollAccettate_GotFocus()
    txtRicircoloTollAccettate.BackColor = colArancione
End Sub

Private Sub txtRicircoloTollAccettate_LostFocus()
    txtRicircoloTollAccettate.BackColor = vbWhite
End Sub

Private Sub txtRientroIndicatore_GotFocus()
    txtRientroIndicatore.BackColor = colArancione
End Sub

Private Sub txtRientroIndicatore_LostFocus()
    txtRientroIndicatore.BackColor = vbWhite
End Sub

Private Sub txtRientroParametri_GotFocus()
    txtRientroParametri.BackColor = colArancione
End Sub

Private Sub txtRientroParametri_LostFocus()
    txtRientroParametri.BackColor = vbWhite
End Sub

Private Sub txtRientroTollAccettate_GotFocus()
    txtRientroTollAccettate.BackColor = colArancione
End Sub

Private Sub txtRientroTollAccettate_LostFocus()
    txtRientroTollAccettate.BackColor = vbWhite
End Sub

