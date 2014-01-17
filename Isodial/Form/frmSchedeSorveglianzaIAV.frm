VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmSchedeSorveglianzaFAV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scheda Sorveglianza FAV"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRilevazione 
      Caption         =   "Rilevazione Pressione"
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
      Height          =   1845
      Left            =   120
      TabIndex        =   72
      Top             =   3400
      Width           =   10935
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   78
         Top             =   600
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   77
         Top             =   600
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   76
         Top             =   960
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   75
         Top             =   960
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   74
         Top             =   1320
         Width           =   3375
      End
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   73
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In Aspirazione"
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
         Height          =   240
         Index           =   7
         Left            =   2760
         TabIndex        =   86
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In Rientro"
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
         Height          =   240
         Index           =   8
         Left            =   8400
         TabIndex        =   85
         Top             =   240
         Width           =   1110
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
         Left            =   600
         TabIndex        =   84
         Top             =   600
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
         Index           =   10
         Left            =   6120
         TabIndex        =   83
         Top             =   600
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
         Index           =   12
         Left            =   600
         TabIndex        =   82
         Top             =   960
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
         Index           =   13
         Left            =   6120
         TabIndex        =   81
         Top             =   960
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
         Index           =   14
         Left            =   120
         TabIndex        =   80
         Top             =   1320
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
         Index           =   15
         Left            =   5640
         TabIndex        =   79
         Top             =   1320
         Width           =   1515
      End
   End
   Begin VB.Frame fraPortataRicircolo 
      Caption         =   "Valutazione Portata e Ricircolo"
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
      Height          =   1845
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   10935
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   18
         Top             =   600
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   17
         Top             =   600
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   16
         Top             =   960
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   15
         Top             =   960
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1320
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
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1320
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
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Portata"
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
         Height          =   240
         Index           =   23
         Left            =   3000
         TabIndex        =   26
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ricircolo"
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
         Height          =   240
         Index           =   22
         Left            =   8520
         TabIndex        =   25
         Top             =   240
         Width           =   1050
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
         Left            =   600
         TabIndex        =   24
         Top             =   600
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
         Left            =   6120
         TabIndex        =   23
         Top             =   600
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
         Left            =   600
         TabIndex        =   22
         Top             =   960
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
         Left            =   6120
         TabIndex        =   21
         Top             =   960
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
         Left            =   5640
         TabIndex        =   19
         Top             =   1320
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2415
      Left            =   120
      TabIndex        =   42
      Top             =   955
      Width           =   3735
      Begin VB.OptionButton optNoAccessoVascolare 
         Caption         =   "No"
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
         Left            =   2640
         TabIndex        =   69
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton optSiAccessoVascolare 
         Caption         =   "Si"
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
         Left            =   2040
         TabIndex        =   68
         Top             =   1560
         Width           =   495
      End
      Begin DataTimeBox.uDataTimeBox oDataNuovoAccessoVascolare 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   43
         Top             =   1965
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataScheda 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   88
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin VB.Label lblNomeUtente 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   89
         Top             =   1110
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "Scheda compilata il"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   87
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "COMPILATORE"
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
         Index           =   24
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Data Nuovo Accesso Vasc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblCognomeUtente 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   47
         Top             =   870
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E' stato necessario eseguire un nuovo accesso vascolare?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   28
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   3465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipoUtente 
         Height          =   255
         Index           =   27
         Left            =   1560
         TabIndex        =   45
         Top             =   630
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UTENTE"
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
         TabIndex        =   44
         Top             =   720
         Width           =   1425
         WordWrap        =   -1  'True
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
      Top             =   955
      Width           =   7215
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   6975
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
            Left            =   3720
            TabIndex        =   50
            Top             =   120
            Width           =   975
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
            Left            =   5880
            TabIndex        =   49
            Top             =   120
            Width           =   975
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
            Left            =   4770
            TabIndex        =   48
            Top             =   120
            Width           =   1095
         End
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
            Left            =   2880
            TabIndex        =   33
            Top             =   120
            Width           =   855
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
            Left            =   2040
            TabIndex        =   32
            Top             =   120
            Width           =   855
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
            TabIndex        =   63
            Top             =   120
            Width           =   810
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   615
         Width           =   6975
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
            Left            =   3720
            TabIndex        =   53
            Top             =   120
            Width           =   975
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
            Left            =   4770
            TabIndex        =   52
            Top             =   120
            Width           =   975
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
            Left            =   5880
            TabIndex        =   51
            Top             =   120
            Width           =   975
         End
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
            Left            =   2880
            TabIndex        =   35
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
            Left            =   2040
            TabIndex        =   34
            Top             =   120
            Width           =   855
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
            TabIndex        =   64
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   120
         TabIndex        =   27
         Top             =   1005
         Width           =   6975
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
            Left            =   3720
            TabIndex        =   56
            Top             =   120
            Width           =   975
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
            Left            =   4770
            TabIndex        =   55
            Top             =   120
            Width           =   975
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
            Left            =   5880
            TabIndex        =   54
            Top             =   120
            Width           =   975
         End
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
            Left            =   2880
            TabIndex        =   37
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
            Left            =   2040
            TabIndex        =   36
            Top             =   120
            Width           =   855
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
            TabIndex        =   65
            Top             =   120
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   1395
         Width           =   6975
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
            Left            =   3720
            TabIndex        =   59
            Top             =   120
            Width           =   975
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
            Left            =   4770
            TabIndex        =   58
            Top             =   120
            Width           =   975
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
            Left            =   5880
            TabIndex        =   57
            Top             =   120
            Width           =   975
         End
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
            Left            =   2880
            TabIndex        =   39
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
            Left            =   2040
            TabIndex        =   38
            Top             =   120
            Width           =   855
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
            TabIndex        =   66
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   1785
         Width           =   6975
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
            Left            =   3720
            TabIndex        =   62
            Top             =   120
            Width           =   975
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
            Left            =   4770
            TabIndex        =   61
            Top             =   120
            Width           =   975
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
            Left            =   5880
            TabIndex        =   60
            Top             =   120
            Width           =   975
         End
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
            Left            =   2880
            TabIndex        =   41
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
            Left            =   2040
            TabIndex        =   40
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Presenza Fremiti"
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
            TabIndex        =   67
            Top             =   120
            Width           =   1875
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   970
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   120
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
         Left            =   9480
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
         Left            =   5400
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
         Left            =   720
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
         Left            =   10080
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
         Left            =   6120
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   10935
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
         CausesValidation=   0   'False
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
         Height          =   495
         Left            =   6480
         TabIndex        =   90
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
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
         Left            =   7920
         TabIndex        =   10
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
         Left            =   9360
         TabIndex        =   9
         Top             =   240
         Width           =   1335
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

    If oDataScheda(0).data = "" Then
        MsgBox "Inserire la Data della Scheda", vbCritical, "ATTENZIONE!!!!"
        Exit Sub
    End If
    
    If Not modifica Then
        keyId = GetNumero("SCHEDA_SORV_FAV")
    End If
    
    v_Nomi = Array("KEY", "KEY_PAZIENTE", "DATA", _
            "ERI_SI_NO", "ERI_VALORE", _
            "DOL_SI_NO", "DOL_VALORE", _
            "GON_SI_NO", "GON_VALORE", _
            "INF_SI_NO", "INF_VALORE", _
            "PRE_FRE_SI_NO", "PRE_FRE_VALORE", _
            "ASP_INDICATORI", "ASP_PARAMETRI", "ASP_TOLL_ACCET", _
            "RIE_INDICATORI", "RIE_PARAMETRI", "RIE_TOLL_ACCET", _
            "POR_INDICATORI", "POR_PARAMETRI", "P0R_TOLL_ACCET", _
            "RIC_INDICATORI", "RIC_PARAMETRI", "RIC_TOLL_ACCET", _
            "ACC_VAS_SI_NO", "ACC_VAS_DATA", _
            "UTENTE_COMP_TIPO", "UTENTE_COMP_COGNOME", "UTENTE_COMP_NOME")

    v_Val = Array(keyId, PazienteKey, oDataScheda(0).data, _
            GestisciSiNoEritema, GestisciOptEritema, _
            GestisciSiNoDolore, GestisciOptDolore, _
            GestisciSiNoGonfiore, GestisciOptGonfiore, _
            GestisciSiNoInfiltrazione, GestisciOptInfiltrazione, _
            GestisciSiNoPresenzaFremiti, GestisciOptPresenzaFremiti, _
            txtAspirazioneIndicatore, txtAspirazioneParametri, txtAspirazioneTollAccettate, _
            txtRientroIndicatore, txtRientroParametri, txtRientroTollAccettate, _
            txtPortataIndicatori, txtPortataParametri, txtPortataTollAccettate, _
            txtRicircoloIndicatori, txtRicircoloParametri, txtRicircoloTollAccettate, _
            GestisciSiNoAccessoVascolare, IIf(oDataNuovoAccessoVascolare(2).data = "", Null, oDataNuovoAccessoVascolare(2).data), _
            GestisciTipoUtenteCompilatore, GestisciCognomeUtenteCompilatore, GestisciNomeUtenteCompilatore)
        
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
    cmdStampa.Enabled = True
End Sub

Private Function GestisciTipoUtenteCompilatore() As String
    If Choose(tAccesso.Tipo, "Medico", "Infermiere", "Contabile", "Amministratore") <> lblTipoUtente(27).Caption Then
        GestisciTipoUtenteCompilatore = Choose(tAccesso.Tipo, "Medico", "Infermiere", "Contabile", "Amministratore")
    Else
        GestisciTipoUtenteCompilatore = lblTipoUtente(27).Caption
    End If
End Function

Private Function GestisciCognomeUtenteCompilatore() As String
    If tAccesso.cognome <> lblCognomeUtente(0).Caption Then
        GestisciCognomeUtenteCompilatore = tAccesso.cognome
    Else
        GestisciCognomeUtenteCompilatore = lblCognomeUtente(0).Caption
    End If
End Function

Private Function GestisciNomeUtenteCompilatore() As String
    If tAccesso.nome <> lblNomeUtente(1).Caption Then
        GestisciNomeUtenteCompilatore = tAccesso.nome
    Else
        GestisciNomeUtenteCompilatore = lblNomeUtente(1).Caption
    End If
End Function

Private Function GestisciSiNoAccessoVascolare() As String
    If optNoAccessoVascolare.Value = True Then
        GestisciSiNoAccessoVascolare = "NO"
    ElseIf optSiAccessoVascolare.Value = True Then
        GestisciSiNoAccessoVascolare = "SI"
    End If
End Function

Private Function GestisciSiNoPresenzaFremiti() As String
    If optNoPresenzaFremiti.Value = True Then
        GestisciSiNoPresenzaFremiti = "NO"
    ElseIf optSiPresenzaFremiti.Value = True Then
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
    ElseIf optSiInfiltrazione.Value = True Then
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
    ElseIf optSiGonfiore.Value = True Then
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
    ElseIf optSiDolore.Value = True Then
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
    ElseIf optSiEritema.Value = True Then
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

Private Sub cmdStampa_Click()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim data As Date
    
    'CARICA IL PAZIENTE
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & PazienteKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDataset("COGNOME") & " " & rsDataset("NOME")
    Set rsDataset = Nothing
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(2) AS ERI_SI_NO, " & _
                "       NEW adVarChar(6) AS ERI_VALORE, " & _
                "       NEW adVarChar(2) AS DOL_SI_NO, " & _
                "       NEW adVarChar(6) AS DOL_VALORE, " & _
                "       NEW adVarChar(2) AS GON_SI_NO, " & _
                "       NEW adVarChar(6) AS GON_VALORE, " & _
                "       NEW adVarChar(2) AS INF_SI_NO, " & _
                "       NEW adVarChar(6) AS INF_VALORE, " & _
                "       NEW adVarChar(2) AS PRE_FRE_SI_NO, " & _
                "       NEW adVarChar(6) AS PRE_FRE_VALORE, " & _
                "       NEW adVarChar(30) AS ASP_INDICATORI, " & _
                "       NEW adVarChar(30) AS ASP_PARAMETRI, " & _
                "       NEW adVarChar(30) AS ASP_TOLL_ACCET, " & _
                "       NEW adVarChar(30) AS RIE_INDICATORI, " & _
                "       NEW adVarChar(30) AS RIE_PARAMETRI, " & _
                "       NEW adVarChar(30) AS RIE_TOLL_ACCET, "
    SQLString = SQLString & _
                "       NEW adVarChar(30) AS POR_INDICATORI, " & _
                "       NEW adVarChar(30) AS POR_PARAMETRI, " & _
                "       NEW adVarChar(30) AS P0R_TOLL_ACCET, " & _
                "       NEW adVarChar(30) AS RIC_INDICATORI, " & _
                "       NEW adVarChar(30) AS RIC_PARAMETRI, " & _
                "       NEW adVarChar(30) AS RIC_TOLL_ACCET, " & _
                "       NEW adVarChar(30) AS ACC_VAS_SI_NO_DATA "
                
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    ' CARICO LA DATA AMERICANA PER LA RICERCA
    data = oDataScheda(0).DataAmericana
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey & " AND DATA=#" & data & "#", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            Do While Not rsDataset.EOF
                .AddNew
                .Fields("ERI_SI_NO") = rsDataset("ERI_SI_NO")
                .Fields("ERI_VALORE") = rsDataset("ERI_VALORE")
                .Fields("DOL_SI_NO") = rsDataset("DOL_SI_NO")
                .Fields("DOL_VALORE") = rsDataset("DOL_VALORE")
                .Fields("GON_SI_NO") = rsDataset("GON_SI_NO")
                .Fields("GON_VALORE") = rsDataset("GON_VALORE")
                .Fields("INF_SI_NO") = rsDataset("INF_SI_NO")
                .Fields("INF_VALORE") = rsDataset("INF_VALORE")
                .Fields("PRE_FRE_SI_NO") = rsDataset("PRE_FRE_SI_NO")
                .Fields("PRE_FRE_VALORE") = rsDataset("PRE_FRE_VALORE")
                .Fields("ASP_INDICATORI") = rsDataset("ASP_INDICATORI")
                .Fields("ASP_PARAMETRI") = rsDataset("ASP_PARAMETRI")
                .Fields("ASP_TOLL_ACCET") = rsDataset("ASP_TOLL_ACCET")
                .Fields("RIE_INDICATORI") = rsDataset("RIE_INDICATORI")
                .Fields("RIE_PARAMETRI") = rsDataset("RIE_PARAMETRI")
                .Fields("RIE_TOLL_ACCET") = rsDataset("RIE_TOLL_ACCET")
                .Fields("POR_INDICATORI") = rsDataset("POR_INDICATORI")
                .Fields("POR_PARAMETRI") = rsDataset("POR_PARAMETRI")
                .Fields("P0R_TOLL_ACCET") = rsDataset("P0R_TOLL_ACCET")
                .Fields("RIC_INDICATORI") = rsDataset("RIC_INDICATORI")
                .Fields("RIC_PARAMETRI") = rsDataset("RIC_PARAMETRI")
                .Fields("RIC_TOLL_ACCET") = rsDataset("RIC_TOLL_ACCET")
                .Fields("ACC_VAS_SI_NO_DATA") = rsDataset("ACC_VAS_SI_NO") & "    in data " & rsDataset("ACC_VAS_DATA")
                rsDataset.MoveNext
            Loop
        End With
    End If
    
    If rsDataset.RecordCount = 0 Then
        MsgBox "Scheda NON presente con la data selezionata", vbInformation, "Informazione"
        Exit Sub
    End If
    Set rsDataset = Nothing
    
    Set rptSchedaSorveglianzaFav.DataSource = rsMain
    'rptStampaApparati.TopMargin = 0
    'rptStampaApparati.RightMargin = 0
    'rptStampaApparati.LeftMargin = 0
    rptSchedaSorveglianzaFav.Sections("Intestazione").Controls.Item("lblDataScheda").Caption = oDataScheda(0).data
    rptSchedaSorveglianzaFav.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptSchedaSorveglianzaFav.Sections("Intestazione").Controls.Item("lblTipoUtenteCompilatore").Caption = lblTipoUtente(27).Caption
    rptSchedaSorveglianzaFav.Sections("Intestazione").Controls.Item("lblCognomeNomeUtenteCompilatore").Caption = lblCognomeUtente(0).Caption & " " & lblNomeUtente(1).Caption
    rptSchedaSorveglianzaFav.PrintReport True, rptRangeAllPages

End Sub

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
    optNoEritema.Value = True
    optSiEritema.Value = False
    chkEritemaLieve.Value = Unchecked
    chkEritemaMedio.Value = Unchecked
    chkEritemaGrave.Value = Unchecked
    chkEritemaLieve.Enabled = False
    chkEritemaMedio.Enabled = False
    chkEritemaGrave.Enabled = False
    
    'Dolore
    optNoDolore.Value = True
    optSiDolore.Value = False
    chkDoloreLieve.Value = Unchecked
    chkDoloreMedio.Value = Unchecked
    chkDoloreGrave.Value = Unchecked
    chkDoloreLieve.Enabled = False
    chkDoloreMedio.Enabled = False
    chkDoloreGrave.Enabled = False
    
    'Gonfiore
    optNoGonfiore.Value = True
    optSiGonfiore.Value = False
    chkGonfioreLieve.Value = Unchecked
    chkGonfioreMedio.Value = Unchecked
    chkGonfioreGrave.Value = Unchecked
    chkGonfioreLieve.Enabled = False
    chkGonfioreMedio.Enabled = False
    chkGonfioreGrave.Enabled = False
    
    'Infiltrazione
    optNoInfiltrazione.Value = True
    optSiInfiltrazione.Value = False
    chkInfiltrazioneLieve.Value = Unchecked
    chkInfiltrazioneMedio.Value = Unchecked
    chkInfiltrazioneGrave.Value = Unchecked
    chkInfiltrazioneLieve.Enabled = False
    chkInfiltrazioneMedio.Enabled = False
    chkInfiltrazioneGrave.Enabled = False
    
    'Presenza fremiti
    optNoPresenzaFremiti.Value = True
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
    
    'Accesso Vascolare
    optNoAccessoVascolare.Value = True
    optSiAccessoVascolare.Value = False
    oDataNuovoAccessoVascolare(2).Pulisci
    Label2.Visible = False
    oDataNuovoAccessoVascolare(2).Visible = False
End Sub

Private Sub cmdTrova_Click()
    Call Pulisci
    oDataScheda(0).Pulisci
    lblTipoUtente(27).Caption = Choose(tAccesso.Tipo, "Medico", "Infermiere", "Contabile", "Amministratore")
    lblCognomeUtente(0).Caption = tAccesso.cognome
    lblNomeUtente(1).Caption = tAccesso.nome
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
       
        ' Di default vado a caricare la data di sistema
        oDataScheda(0).data = date
       
    End If
End Sub

Private Sub CaricaValori()
    Dim data As Date
    
    ' la data americana
    data = oDataScheda(0).DataAmericana
    
    Set rsDataset = New Recordset

    rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey & " AND DATA=#" & data & "#", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    
    Call Pulisci
    
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
        
        Call CaricaAccessoVascolare
        
        Call CaricaUtenteCompilatore
        
        modifica = True
        cmdStampa.Enabled = True
    Else
        'Se non trova il paziente associato vuol dire che è in fase di inserimento
        modifica = False
        cmdStampa.Enabled = False
    End If
    
    Set rsDataset = Nothing
    
End Sub

Private Sub CaricaUtenteCompilatore()
    lblTipoUtente(27).Caption = rsDataset("UTENTE_COMP_TIPO") & ""
    lblCognomeUtente(0).Caption = rsDataset("UTENTE_COMP_COGNOME") & ""
    lblNomeUtente(1).Caption = rsDataset("UTENTE_COMP_NOME") & ""
End Sub

Private Sub CaricaAccessoVascolare()
    If rsDataset("ACC_VAS_SI_NO") = "NO" Then
        optNoAccessoVascolare.Value = True
    ElseIf rsDataset("ACC_VAS_SI_NO") = "SI" Then
        optSiAccessoVascolare.Value = True
        Label2.Visible = True
        oDataNuovoAccessoVascolare(2).Visible = True
        oDataNuovoAccessoVascolare(2).data = rsDataset("ACC_VAS_DATA") & ""
    End If
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

Private Sub oDataScheda_OnDataChange(Index As Integer)
    If IsDate(oDataScheda(0).data) Then
        Call CaricaValori
    ElseIf oDataScheda(0).data = "" Then
        Call Pulisci
    End If
End Sub

Private Sub oDataScheda_OnDataClick(Index As Integer)
    oDataScheda(0).Pulisci
    Call Pulisci
End Sub

Private Sub oDataScheda_OnElencaClick(Index As Integer)
    tElenca.Tipo = tpSCHEDA_SORV_FAV
    tElenca.condizione = "WHERE KEY_PAZIENTE=" & PazienteKey
    frmElencaDate.Show 1
    If laData <> "" Then oDataScheda(0).data = laData
End Sub

Private Sub optNoAccessoVascolare_GotFocus()
    Label2.Visible = False
    oDataNuovoAccessoVascolare(2).Visible = False
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

Private Sub optSiAccessoVascolare_GotFocus()
    Label2.Visible = True
    oDataNuovoAccessoVascolare(2).Visible = True
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

