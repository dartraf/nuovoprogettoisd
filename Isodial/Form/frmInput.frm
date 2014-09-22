VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserimento Valori"
   ClientHeight    =   8388
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   12588
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8388
   ScaleWidth      =   12588
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraComuni 
      Height          =   1695
      Left            =   6480
      TabIndex        =   164
      Top             =   6480
      Width           =   6015
      Begin VB.ComboBox cboRegComuni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtComune 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   166
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtCodIstat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   165
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice ISTAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   45
         Left            =   120
         TabIndex        =   170
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label Comune 
         AutoSize        =   -1  'True
         Caption         =   "Comune"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   169
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   44
         Left            =   120
         TabIndex        =   168
         Top             =   1200
         Width           =   900
      End
   End
   Begin VB.Frame fraVoci 
      Height          =   2055
      Left            =   6480
      TabIndex        =   58
      Top             =   5880
      Width           =   6255
      Begin VB.CheckBox chkStampaVoce 
         Caption         =   "Stampa in box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   1110
         Width           =   1815
      End
      Begin VB.TextBox txtValoreMin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtValoreMax 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtUnita 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkPN 
         Caption         =   "Pos/Neg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtVoce 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   45
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
      Begin VB.CheckBox chkEsameDaStampare 
         Caption         =   "Esami da Stampare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   163
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valore min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valore max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unità di misura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraEsenzioni 
      Height          =   1185
      Left            =   6480
      TabIndex        =   153
      Top             =   5400
      Width           =   6015
      Begin VB.OptionButton optTicketRicetta 
         Caption         =   "Non Esente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   156
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCodiceEsenzione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   154
         Top             =   225
         Width           =   1575
      End
      Begin VB.OptionButton optTicketRicetta 
         Caption         =   "Esente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   155
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quota aggiuntiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   39
         Left            =   120
         TabIndex        =   158
         Top             =   720
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice esenzione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   120
         TabIndex        =   157
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Frame fraNomenclatore 
      Height          =   1695
      Left            =   6480
      TabIndex        =   125
      Top             =   4680
      Width           =   6015
      Begin VB.TextBox txtImportoScontato 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   129
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtImporto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   128
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCodicePrestazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   126
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNomePrestazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   127
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importo scontato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   33
         Left            =   2880
         TabIndex        =   133
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   32
         Left            =   120
         TabIndex        =   132
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Descrizione Prestazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   131
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   120
         TabIndex        =   130
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame fraTerapiaStraordinaria 
      Height          =   1695
      Left            =   6480
      TabIndex        =   145
      Top             =   4080
      Width           =   6015
      Begin VB.TextBox txtNoteStraordinarie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   149
         Top             =   1200
         Width           =   4455
      End
      Begin VB.ComboBox cboMedicinaliStraordinaria 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPosologiaStraordinaria 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   147
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkConfermaStraordinaria 
         Caption         =   "Conferma Somministazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   148
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   38
         Left            =   120
         TabIndex        =   152
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Farmaco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   151
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Posologia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   37
         Left            =   120
         TabIndex        =   150
         Top             =   780
         Width           =   1080
      End
   End
   Begin VB.Frame fraAsl 
      Height          =   1695
      Left            =   6480
      TabIndex        =   134
      Top             =   3480
      Width           =   6015
      Begin VB.TextBox txtCodiceAsl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   135
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNomeAsl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   136
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cboRegione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   120
         TabIndex        =   140
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome ASL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   139
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice ASL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   34
         Left            =   120
         TabIndex        =   138
         Top             =   300
         Width           =   1230
      End
   End
   Begin VB.Frame fraPrestazione 
      Height          =   2055
      Left            =   6480
      TabIndex        =   105
      Top             =   2880
      Width           =   6015
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   8
         Left            =   3360
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   110
         ToolTipText     =   "Cerca data"
         Top             =   1560
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   7
         Left            =   3360
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   109
         ToolTipText     =   "Cerca data"
         Top             =   1080
         Width           =   360
      End
      Begin VB.TextBox txtQuantita 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   108
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cboPrescrizioni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         TabIndex        =   106
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox cboCodicePrestazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         TabIndex        =   107
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fine prestazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   117
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   116
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inizio prestazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   115
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   114
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantità"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         TabIndex        =   113
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Prestazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   112
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   111
         Top             =   765
         Width           =   750
      End
   End
   Begin VB.Frame fraColture 
      Height          =   2415
      Left            =   6480
      TabIndex        =   90
      Top             =   2400
      Width           =   6015
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   6
         Left            =   3360
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   48
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2040
         TabIndex        =   95
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton optEsitoBagno 
            Caption         =   "Pos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   53
            Top             =   90
            Width           =   735
         End
         Begin VB.OptionButton optEsitoBagno 
            Caption         =   "Neg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   54
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.TextBox txtColtureBagno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   52
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtColtureAcqua 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   49
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton optEsitoAcqua 
         Caption         =   "Pos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   50
         Top             =   1170
         Width           =   735
      End
      Begin VB.OptionButton optEsitoAcqua 
         Caption         =   "Neg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   51
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   102
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   120
         TabIndex        =   101
         Top             =   280
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Colture su Acqua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   120
         TabIndex        =   94
         Top             =   720
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   120
         TabIndex        =   93
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Colture su Bagno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   120
         TabIndex        =   92
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   120
         TabIndex        =   91
         Top             =   2000
         Width           =   540
      End
   End
   Begin VB.Frame fraRene 
      Height          =   2655
      Left            =   6480
      TabIndex        =   68
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox txtNumeroRene 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   9
         Left            =   3480
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   43
         ToolTipText     =   "Cerca data"
         Top             =   2160
         Width           =   360
      End
      Begin VB.OptionButton optTipoRene 
         Caption         =   "NEG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   40
         Top             =   840
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTipoRene 
         Caption         =   "HBV+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optTipoRene 
         Caption         =   "HCV+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   38
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtPostazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtMatricola 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   42
         Top             =   1755
         Width           =   3855
      End
      Begin VB.TextBox txtTipoRene 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   41
         Top             =   1305
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero rene"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   43
         Left            =   3600
         TabIndex        =   162
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monitor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   42
         Left            =   120
         TabIndex        =   161
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data rottamazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   41
         Left            =   120
         TabIndex        =   160
         Top             =   2205
         Width           =   1905
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2160
         TabIndex        =   159
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matricola"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di rene"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame fraPassword 
      Height          =   2295
      Left            =   6480
      TabIndex        =   96
      Top             =   1200
      Width           =   6015
      Begin VB.TextBox txtCognomePass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   21
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtNomePass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   22
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtChiave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   23
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optUtente 
         Caption         =   "Medico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   25
         Top             =   1080
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optUtente 
         Caption         =   "Infermiere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optUtente 
         Caption         =   "Contabile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   27
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   100
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   99
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome utente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   98
         Top             =   1245
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   120
         TabIndex        =   97
         Top             =   1680
         Width           =   1035
      End
   End
   Begin VB.Frame fraEpisodi 
      Height          =   1215
      Left            =   120
      TabIndex        =   77
      Top             =   7080
      Width           =   6015
      Begin VB.TextBox txtNoteEpisodi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   35
         TabIndex        =   32
         Top             =   720
         Width           =   4935
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   3
         Left            =   2160
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   31
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   120
         TabIndex        =   80
         Top             =   750
         Width           =   510
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   79
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame fraSieroconversioni 
      Height          =   1215
      Left            =   120
      TabIndex        =   85
      Top             =   6480
      Width           =   6015
      Begin VB.CheckBox chkSieroconversioni 
         Caption         =   "HCV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   47
         Top             =   750
         Width           =   975
      End
      Begin VB.CheckBox chkSieroconversioni 
         Caption         =   "HBV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   46
         Top             =   750
         Width           =   975
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   5
         Left            =   3360
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   35
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   88
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sieroconversioni"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   120
         TabIndex        =   86
         Top             =   750
         Width           =   1770
      End
   End
   Begin VB.Frame fraTrasfusioni 
      Height          =   1215
      Left            =   120
      TabIndex        =   81
      Top             =   5880
      Width           =   6015
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   4
         Left            =   2760
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   33
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.ComboBox cboTrasfusioni 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trasfusione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   120
         TabIndex        =   84
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   82
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraDistretti 
      Height          =   1215
      Left            =   120
      TabIndex        =   118
      Top             =   5280
      Width           =   6015
      Begin VB.ComboBox cboAslAppartenenza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtNomeDistretto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   120
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCodiceDistretto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   119
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Distretto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   120
         TabIndex        =   124
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Distretto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3720
         TabIndex        =   123
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asl di riferimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   120
         TabIndex        =   122
         Top             =   795
         Width           =   1755
      End
   End
   Begin VB.Frame fraRicoveri 
      Height          =   1215
      Left            =   120
      TabIndex        =   72
      Top             =   4680
      Width           =   6015
      Begin VB.TextBox txtNoteRicoveri 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   35
         TabIndex        =   30
         Top             =   720
         Width           =   4935
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   5400
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   29
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   2160
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   28
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   120
         TabIndex        =   89
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   76
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   3480
         TabIndex        =   75
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   120
         TabIndex        =   74
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   73
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraTabelle 
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.OptionButton optMansione 
         Caption         =   "Coordinatore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   67
         Top             =   1200
         Width           =   1500
      End
      Begin VB.OptionButton optMansione 
         Caption         =   "Infermiere professionale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1200
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.ComboBox cboVoci 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmInput.frx":0000
         Left            =   2040
         List            =   "frmInput.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   240
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCognome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   57
      Top             =   1080
      Width           =   6015
      Begin VB.CheckBox chkInserisci 
         Caption         =   "Inserisci nell'organigramma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdAnnulla 
         Cancel          =   -1  'True
         Caption         =   "&Annulla"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   45
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Memorizza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   44
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.Frame fraTerapia 
      Height          =   735
      Left            =   6480
      TabIndex        =   63
      Top             =   600
      Width           =   6015
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   3480
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   172
         ToolTipText     =   "Cerca data"
         Top             =   4200
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.OptionButton optSomministrazione 
         Caption         =   "Postdialitica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   144
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton optSomministrazione 
         Caption         =   "Intradialitica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   143
         Top             =   2880
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Tutti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Do"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   19
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Sa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Ve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Gi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Ma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkGiorni 
         Caption         =   "Lu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtPosologia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   11
         Top             =   720
         Width           =   972
      End
      Begin VB.ComboBox cboMedicinali 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   4572
      End
      Begin VB.TextBox txtSomministrazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   1200
         Width           =   3735
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   171
         Top             =   675
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   656
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataFarmaco1 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   174
         Top             =   2520
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   656
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataFarmaco2 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   176
         Top             =   3000
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   656
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataFarmaco3 
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   178
         Top             =   3480
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   656
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label lblData3 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   48
         Left            =   120
         TabIndex        =   179
         Top             =   3525
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblData2 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   47
         Left            =   120
         TabIndex        =   177
         Top             =   3045
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblData1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   46
         Left            =   120
         TabIndex        =   175
         Top             =   2565
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   173
         Top             =   5700
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Frequenza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   142
         Top             =   1680
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Posologia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   36
         Left            =   3600
         TabIndex        =   141
         Top             =   756
         Width           =   1080
      End
      Begin VB.Label lblMedicinaleTerapie 
         AutoSize        =   -1  'True
         Caption         =   "Farmaco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   936
      End
      Begin VB.Label lblNoteTerapie 
         AutoSize        =   -1  'True
         Caption         =   "Somministrazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   1848
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Inizio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   64
         Top             =   750
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sec As String
Dim lettera As String
Dim rsDataset As Recordset
Private Type prestazioni
    codice As String
    nome As String
End Type
Dim v_prestazioni() As prestazioni

Private Sub Form_Activate()
    Select Case tInput.Tipo
    
        Case tpISINGOLO, tpIESAMI
            txtCognome.SetFocus
            If tInput.mantieniDati Then
                txtCognome = tInput.v_valori(1)
            End If
        
        Case tpICOMPOSTO:
            txtCognome.SetFocus
            If tInput.mantieniDati Then
                txtNome = tInput.v_valori(1)
                txtCognome = tInput.v_valori(2)
            End If
        
        Case tpIVOCI:
            txtVoce.SetFocus
            If tInput.mantieniDati Then
            txtVoce = tInput.v_valori(1)
                If tInput.v_valori(2) = True Then
                    chkPN.Value = Checked
                Else
                    chkPN.Value = Unchecked
                End If
                txtUnita = tInput.v_valori(3)
                txtValoreMin = tInput.v_valori(4)
                txtValoreMax = tInput.v_valori(5)
                chkStampaVoce.Value = Unchecked
                chkEsameDaStampare.Value = Unchecked
            End If
        
        Case tpITIPIESAMILAB
            cboVoci.SetFocus
            If tInput.mantieniDati Then
                cboVoci.ListIndex = GetCboListIndex(CInt(tInput.v_valori(1)), cboVoci)
            End If
        
        Case tpIPASSWORD
            txtCognomePass.SetFocus
            If tInput.mantieniDati Then
                txtChiave = tInput.v_valori(1)
                txtCognomePass = tInput.v_valori(2)
                txtNomePass = tInput.v_valori(3)
                txtPass = tInput.v_valori(4)
                If tInput.v_valori(5) = "1" Then
                    optUtente(0).Value = True
                ElseIf tInput.v_valori(5) = "2" Then
                    optUtente(1).Value = True
                ElseIf tInput.v_valori(5) = "3" Then
                    optUtente(2).Value = True
                End If
                If tInput.v_valori(6) <> "" Then
                    chkInserisci.Value = IIf(CBool(tInput.v_valori(6)), Checked, Unchecked)
                End If
            End If
        
        Case tpITERAPIADOMICILIARE, tpITERAPIADIALITICA
            cboMedicinali.SetFocus
        
        Case tpITERAPIESTRAORDINARIE
            cboMedicinaliStraordinaria.SetFocus
        
        Case tpICOLTURE
            txtColtureAcqua.SetFocus
        
        Case tpIPRESCRIZIONI
            cboPrescrizioni.SetFocus
        
        Case tpIDISTRETTI
            txtCodiceDistretto.SetFocus
            If tInput.mantieniDati Then
                txtCodiceDistretto = tInput.v_valori(1)
                txtNomeDistretto = tInput.v_valori(2)
                cboAslAppartenenza.ListIndex = GetCboListIndex(CInt(tInput.v_valori(3)), cboAslAppartenenza)
            End If
        
        Case tpINOMENCLATORE
            txtCodicePrestazione.SetFocus
            If tInput.mantieniDati Then
                txtCodicePrestazione = tInput.v_valori(1)
                txtNomePrestazione = tInput.v_valori(2)
                txtImporto = tInput.v_valori(3)
                txtImportoScontato = tInput.v_valori(4)
            End If
        
        Case tpICOMUNI
           txtCodIstat.SetFocus
           If tInput.mantieniDati Then
                txtCodIstat = tInput.v_valori(1)
                txtComune = tInput.v_valori(2)
                cboRegComuni.ListIndex = GetCboListIndex(CInt(tInput.v_valori(3)), cboRegComuni)
            End If
        
        Case tpIASL
            txtCodiceAsl.SetFocus
            If tInput.mantieniDati Then
                txtCodiceAsl = tInput.v_valori(1)
                txtNomeAsl = tInput.v_valori(2)
                cboRegione.ListIndex = GetCboListIndex(CInt(tInput.v_valori(3)), cboRegione)
            End If
        
        Case tpIESENZIONE
            txtCodiceEsenzione.SetFocus
            If tInput.mantieniDati Then
                txtCodiceEsenzione = tInput.v_valori(1)
                If tInput.v_valori(1) Then
                    optTicketRicetta(0).Value = True
                Else
                    optTicketRicetta(1).Value = True
                End If
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.Height = 2385
    Me.Width = 6315
    For i = 0 To 9
       lblData(i).BackColor = vbWhite
       picData(i).Picture = LoadResPicture("cal1", 0)
    Next i
    
    Select Case tInput.Tipo
        
        Case tpISINGOLO
            fraTabelle.Height = 700
            fraPulsanti.Top = fraTabelle.Height - 135
            If tTabelle = tpesame Then
                Label1(1) = "Esame"
            End If
        
        Case tpICOMPOSTO
            If tTabelle = tpRegioni Then
                Label1(1) = "Codice Regione"
                txtCognome.MaxLength = 3
                lblNome = "Regione"
                txtNome.MaxLength = 30
            ElseIf tTabelle = tpTIPOLOGIEMEDICO Then
                Label1(1) = "Codice"
                txtCognome.MaxLength = 1
                txtCognome.Width = txtCognome.Width / 8
                lblNome = "Tipologia Medico"
                txtNome.MaxLength = 50
            ElseIf tTabelle = tpEDTA Then
                Label1(1) = "Codice"
                txtCognome.MaxLength = 3
                lblNome = "E.D.T.A."
                txtNome.MaxLength = 150
            ElseIf tTabelle = tpEDTA_MORTE Then
                Label1(1) = "Codice"
                txtCognome.MaxLength = 3
                lblNome = "Causa Morte"
                txtNome.MaxLength = 150
            End If
        
        Case tpIESAMI
            fraTabelle.Height = 615
            fraPulsanti.Top = fraTabelle.Height - 135
            Label1(1) = "Gruppo Esami"
        
        Case tpIVOCI
            frmInput.Width = 6570
            fraPulsanti.Width = 6255
            cmdAnnulla.Left = 4920
            cmdInserisci.Left = 3360
            fraVoci.Top = fraTabelle.Top
            fraVoci.Left = fraTabelle.Left
            fraVoci.ZOrder
            fraPulsanti.Top = fraVoci.Height - 135
        
        Case tpITIPIESAMILAB
            fraTabelle.Height = 615
            fraPulsanti.Top = fraTabelle.Height - 135
            Label1(1) = "Esame"
            Me.Caption = "Inserimento Esami"
            cboVoci.Visible = True
            Call RicaricaComboBox("VOCI_ESAMI", "NOME", cboVoci)
        
        Case tpIPASSWORD
            fraPassword.Top = fraTabelle.Top
            fraPassword.Left = fraTabelle.Left
            fraPassword.ZOrder
            fraPulsanti.Top = fraPassword.Height - 135
            chkInserisci.Visible = False
        
        Case tpITERAPIADOMICILIARE, tpITERAPIADIALITICA
            fraTerapia.Height = 2550
            If tInput.Tipo = tpITERAPIADIALITICA Then
                fraTerapia.Height = 3972
                lblNoteTerapie = "Note"
                txtSomministrazione.Left = 840
                txtSomministrazione.Width = 4335
                lblData1(46).Visible = True
                oDataFarmaco1(1).Visible = True
                lblData2(47).Visible = True
                oDataFarmaco2(2).Visible = True
                lblData3(48).Visible = True
                oDataFarmaco3(3).Visible = True
            End If
            fraTerapia.Top = fraTabelle.Top
            fraTerapia.Left = fraTabelle.Left
            fraTerapia.ZOrder
            fraPulsanti.Top = fraTerapia.Height - 135
            Call RicaricaComboBox("MEDICINALI", "NOME", cboMedicinali)
        
        Case tpIPRESCRIZIONI
            fraPrestazione.Top = fraTabelle.Top
            fraPrestazione.Left = fraTabelle.Left
            fraPrestazione.ZOrder
            fraPulsanti.Top = fraPrestazione.Height - 135
            ' carica le prestazioni
            Set rsDataset = New Recordset
            rsDataset.Open "SELECT * FROM NOMENCLATORE_TARIFFARIO ORDER BY CODICE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            ReDim Preserve v_prestazioni(0)
            Do While Not rsDataset.EOF
                ReDim Preserve v_prestazioni(UBound(v_prestazioni) + 1)
                v_prestazioni(UBound(v_prestazioni)).codice = rsDataset("CODICE")
                v_prestazioni(UBound(v_prestazioni)).nome = rsDataset("NOME")
                cboCodicePrestazione.AddItem rsDataset("CODICE")
                cboCodicePrestazione.ItemData(cboCodicePrestazione.NewIndex) = rsDataset("KEY")
                cboPrescrizioni.AddItem rsDataset("NOME")
                cboPrescrizioni.ItemData(cboPrescrizioni.NewIndex) = rsDataset("KEY")
                rsDataset.MoveNext
            Loop
            rsDataset.Close
            Me.Caption = "Inserimento prescrizione"
            cmdInserisci.Caption = "&Inserisci"
            
            rsDataset.Open "SELECT CODICE_ASL, NOME FROM (INTESTAZIONE_STAMPA I INNER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            cboCodicePrestazione.ListIndex = GetIndex(cboCodicePrestazione, IIf(rsDataset("CODICE_ASL") = 2, "39.95.2", "39.95.4"))
            rsDataset.Close
            
            Set rsDataset = Nothing
        
        Case tpIRICOVERI
            fraRicoveri.Top = fraTabelle.Top
            fraRicoveri.Left = fraTabelle.Left
            fraRicoveri.ZOrder
            fraPulsanti.Top = fraRicoveri.Height - 135
        
        Case tpIEPISODI
            fraEpisodi.Top = fraTabelle.Top
            fraEpisodi.Left = fraTabelle.Left
            fraEpisodi.ZOrder
            fraPulsanti.Top = fraEpisodi.Height - 135
        
        Case tpITRASFUSIONI
            fraTrasfusioni.Top = fraTabelle.Top
            fraTrasfusioni.Left = fraTabelle.Left
            fraTrasfusioni.ZOrder
            fraPulsanti.Top = fraTrasfusioni.Height - 135
            Call RicaricaComboBox("TIPO_TRASFUSIONI", "NOME", cboTrasfusioni)
        
        Case tpISIEROCONVERSIONI
            fraSieroconversioni.Top = fraTabelle.Top
            fraSieroconversioni.Left = fraTabelle.Left
            fraSieroconversioni.ZOrder
            fraPulsanti.Top = fraSieroconversioni.Height - 135
        
        Case tpICOLTURE
            fraColture.Top = fraTabelle.Top
            fraColture.Left = fraTabelle.Left
            fraColture.ZOrder
            fraPulsanti.Top = fraColture.Height - 135
        
        Case tpITERAPIESTRAORDINARIE
            fraTerapiaStraordinaria.Top = fraTabelle.Top
            fraTerapiaStraordinaria.Left = fraTabelle.Left
            fraTerapiaStraordinaria.ZOrder
            fraPulsanti.Top = fraTerapiaStraordinaria.Height - 135
            Call RicaricaComboBox("MEDICINALI", "NOME", cboMedicinaliStraordinaria)
        
        Case tpIDISTRETTI
            fraDistretti.Top = fraTabelle.Top
            fraDistretti.Left = fraTabelle.Left
            fraDistretti.ZOrder
            fraPulsanti.Top = fraDistretti.Height - 135
            Call RicaricaComboBox("ASL", "NOME", cboAslAppartenenza)
        
        Case tpINOMENCLATORE
            fraNomenclatore.Top = fraTabelle.Top
            fraNomenclatore.Left = fraTabelle.Left
            fraNomenclatore.ZOrder
            fraPulsanti.Top = fraNomenclatore.Height - 135
        
        Case tpICOMUNI
            fraComuni.Top = fraTabelle.Top
            fraComuni.Left = fraTabelle.Left
            fraComuni.ZOrder
            fraPulsanti.Top = fraAsl.Height - 135
            Call RicaricaComboBox("REGIONI ORDER BY NOME", "NOME", cboRegComuni)
        
        Case tpIASL
            fraAsl.Top = fraTabelle.Top
            fraAsl.Left = fraTabelle.Left
            fraAsl.ZOrder
            fraPulsanti.Top = fraAsl.Height - 135
            Call RicaricaComboBox("REGIONI ORDER BY NOME", "NOME", cboRegione)
        
        Case tpIESENZIONE
            fraEsenzioni.Top = fraTabelle.Top
            fraEsenzioni.Left = fraTabelle.Left
            fraEsenzioni.ZOrder
            fraPulsanti.Top = fraEsenzioni.Height - 135
            
    End Select
    Me.Height = fraPulsanti.Top + fraPulsanti.Height + 480
    Call TakeCloseOff(Me.hWnd)
    If Not tInput.mantieniDati Then
        Call PulisciVar
    End If
End Sub

Private Function GestisciOpt() As Integer
    If optUtente(0).Value = True Then
        GestisciOpt = 1
    ElseIf optUtente(1).Value = True Then
        GestisciOpt = 2
    ElseIf optUtente(2).Value = True Then
        GestisciOpt = 3
    End If
End Function

Private Sub PulisciVar()
    Dim i As Integer
    For i = 1 To 5
        tInput.v_valori(i) = ""
    Next i
End Sub

Private Function GestisciOptTre(opt As Object) As Integer
    If opt(0).Value = False And opt(1).Value = False Then
        GestisciOptTre = 0
    Else
        If opt(0).Value Then
            GestisciOptTre = 1
        Else
            GestisciOptTre = 2
        End If
    End If
End Function

Private Function getIndexCodicePrestazione() As Integer
    ' trova la prestazione nel vettore
    Dim i As Integer
    Dim k As Integer
    
    For i = 1 To UBound(v_prestazioni)
        If cboPrescrizioni.Text = v_prestazioni(i).nome Then
            Exit For
        End If
    Next i
    ' trova il relativo codice nel cbo
    For k = 0 To cboCodicePrestazione.ListCount - 1
        If cboCodicePrestazione.List(k) = v_prestazioni(i).codice Then
            getIndexCodicePrestazione = k
            Exit For
        End If
    Next k
End Function

Private Function getIndexNomePrestazione() As Integer
    ' trova codice nel vettore
    Dim i As Integer
    Dim k As Integer
    
    For i = 1 To UBound(v_prestazioni)
        If cboCodicePrestazione.Text = v_prestazioni(i).codice Then
            Exit For
        End If
    Next i
    ' trova il relativo codice nel cbo
    For k = 0 To cboPrescrizioni.ListCount - 1
        If cboPrescrizioni.List(k) = v_prestazioni(i).nome Then
            getIndexNomePrestazione = k
            Exit For
        End If
    Next k
End Function

Private Sub cmdAnnulla_LostFocus()
    Select Case tInput.Tipo
        Case tpIVOCI
            txtVoce.SetFocus
        Case tpITERAPIADIALITICA, tpITERAPIADOMICILIARE
            cboMedicinali.SetFocus
        Case tpIPASSWORD
            txtCognomePass.SetFocus
        Case tpICOLTURE
            txtColtureAcqua.SetFocus
        Case tpIRICOVERI
            txtNoteRicoveri.SetFocus
        Case tpIDISTRETTI
            txtCodiceDistretto.SetFocus
        Case tpITRASFUSIONI
            cboTrasfusioni.SetFocus
        Case tpISIEROCONVERSIONI
            chkSieroconversioni(0).SetFocus
        Case tpIEPISODI
            txtNoteEpisodi.SetFocus
        Case tpIPRESCRIZIONI
            cboPrescrizioni.SetFocus
        Case tpICOMUNI
            txtCodIstat.SetFocus
        Case tpIASL
            txtCodiceAsl.SetFocus
        Case tpITERAPIESTRAORDINARIE
            cboMedicinaliStraordinaria.SetFocus
        Case tpINOMENCLATORE
            txtCodicePrestazione.SetFocus
        Case tpIESENZIONE
            txtCodiceEsenzione.SetFocus
    End Select
End Sub

Private Sub cmdAnnulla_Click()
    tInput.v_valori(1) = ""
    tInput.v_valori(2) = ""
    If tInput.Tipo = tpIPRESCRIZIONI Or tInput.Tipo = tpITIPIESAMILAB Or tInput.Tipo = tpITERAPIESTRAORDINARIE Then
        tInput.v_valori(1) = -1
        tInput.v_valori(1) = -1
    End If
    If tInput.Tipo = tpICOLTURE Then
        tInput.v_valori(1) = ""
        tInput.v_valori(2) = 0
        tInput.v_valori(3) = ""
        tInput.v_valori(4) = 0
    End If
    Unload Me
End Sub

Private Sub cmdInserisci_Click()
    Dim i As Integer
    Dim trovato As Boolean

    Select Case tInput.Tipo
    Case tpISINGOLO:
        tInput.v_valori(1) = UCase(txtCognome)                       ' obbligatorio per singolo
    
    Case tpICOMPOSTO:
        If txtNome = "" Or txtCognome = "" Then
            MsgBox "Inserire tutti i valori", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        Else
            tInput.v_valori(1) = txtNome                       ' obbligatorio per composto
            tInput.v_valori(2) = txtCognome
        End If
    
    Case tpIESAMI:
        tInput.v_valori(1) = txtCognome
    
    Case tpIVOCI:
        If txtValoreMin = "" Then
            txtValoreMin = 0
        End If
        If txtValoreMax = "" Then
            txtValoreMax = 0
        End If
        If val(txtValoreMax) < val(txtValoreMin) Then
            MsgBox "Il valore massimo non può essere inferiore al valore minimo", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        End If
        If tInput.v_valori(6) >= 3 And chkStampaVoce.Value = Checked Then
            MsgBox "IMPOSSIBILE stampare nel box - I tre esami sono già selezionati", vbCritical, "ATTENZIONE!!!"
            chkStampaVoce.Value = Unchecked
            Exit Sub
        End If
        If tInput.v_valori(7) >= 16 And chkEsameDaStampare.Value = Checked Then
            MsgBox "IMPOSSIBILE stampare nel box - I sedici esami sono già selezionati", vbCritical, "ATTENZIONE!!!"
            chkEsameDaStampare.Value = Unchecked
            Exit Sub
        End If
        tInput.v_valori(1) = txtVoce                       ' obbligatorio per voci
        tInput.v_valori(2) = IIf(chkPN.Value = Checked, True, False)
        tInput.v_valori(3) = txtUnita
        tInput.v_valori(4) = IIf(txtValoreMin = "", 0, txtValoreMin)
        tInput.v_valori(5) = IIf(txtValoreMax = "", 0, txtValoreMax)
        tInput.v_valori(6) = IIf(chkStampaVoce.Value = Checked, True, False)
        tInput.v_valori(7) = IIf(chkEsameDaStampare.Value = Checked, True, False)
    
    Case tpITIPIESAMILAB:
        If cboVoci.Text = "" Then
            MsgBox "Selezionare l' esame", vbInformation, "Informazione"
            Exit Sub
            Else
            tInput.v_valori(1) = cboVoci.ItemData(cboVoci.ListIndex)           ' obbligatorio per tipi esami lab
        End If
    
    Case tpIPASSWORD:
        If txtCognomePass = "" Then
            MsgBox "Il campo COGNOME è obbligatorio", vbInformation, "Informazione"
            txtCognomePass.SetFocus
            Exit Sub
        End If
        If txtNomePass = "" Then
            MsgBox "Il campo NOME è obbligatorio", vbInformation, "Informazione"
            txtNomePass.SetFocus
            Exit Sub
        End If
        If Len(txtChiave.Text) < 8 Then
            MsgBox "ATTENZIONE!!!: " & vbCrLf & "Il Nome Utente deve essere almeno di 8 caratteri!!!", vbInformation, "Informazione"
            txtChiave.SetFocus
            Exit Sub
        End If
        If Len(txtPass.Text) < 8 Then
            MsgBox "ATTENZIONE!!!: " & vbCrLf & "La Password deve essere almeno di 8 caratteri!!!", vbInformation, "Informazione"
            txtPass.SetFocus
            Exit Sub
        End If
        If ControlloDuplicato("LOGIN") Then
           txtCognomePass.SetFocus
           Exit Sub
        End If
        tInput.v_valori(1) = txtChiave
        tInput.v_valori(2) = UCase(txtCognomePass)
        tInput.v_valori(3) = UCase(txtNomePass)
        tInput.v_valori(4) = txtPass
        tInput.v_valori(5) = GestisciOpt
        tInput.v_valori(6) = (chkInserisci.Value = Checked)
        
    'Terapia Domiciliare e Dialitica
    Case tpITERAPIADOMICILIARE, tpITERAPIADIALITICA:
        If cboMedicinali.ListIndex = -1 Then
            MsgBox "Selezionare il farmaco", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        Else
            tInput.v_valori(2) = cboMedicinali.ListIndex
        End If
        If oData(0).data = "" Then
            MsgBox "Inserire la data di inizio terapia", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        Else
            tInput.v_valori(1) = oData(0).data
        End If
        If txtPosologia = "" Then
            MsgBox "Inserire la posologia", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        Else
            tInput.v_valori(3) = txtPosologia
        End If
        For i = 0 To 7
            If chkGiorni(i).Value = Checked Then
                trovato = True
                Exit For
            End If
        Next i
        If tInput.Tipo = tpITERAPIADIALITICA Then
            sec = "o definire una data"
        Else
            sec = ""
        End If
        If trovato = False And oDataFarmaco1(1).data = "" And oDataFarmaco2(2).data = "" And oDataFarmaco3(3).data = "" Then
            MsgBox "Indicare almeno un giorno della settimana " & sec, vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        ElseIf trovato And (oDataFarmaco1(1).data <> "" Or oDataFarmaco2(2).data <> "" Or oDataFarmaco3(3).data <> "") Then
            MsgBox "INDICAZIONE DOPPIA - Indicare un giorno della settimana o una data", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If tInput.Tipo = tpITERAPIADIALITICA Then
            tInput.v_valori(4) = IIf(optSomministrazione(0).Value, 1, 2)
        Else
            tInput.v_valori(4) = txtSomministrazione & ""
        End If
        For i = 0 To 7
            tInput.v_valori(5) = tInput.v_valori(5) & IIf(chkGiorni(i).Value = Checked, 1, 0) & "-"
        Next i
       ' tInput.v_valori(5) = tInput.v_valori(5) & IIf(chkConfermaSomministrazione.Value = Checked, 1, 0) & "-"
        tInput.v_valori(5) = Mid(tInput.v_valori(5), 1, Len(tInput.v_valori(5)) - 1)
        
        If tInput.Tipo = tpITERAPIADIALITICA Then
            tInput.v_valori(6) = txtSomministrazione & ""
            tInput.v_valori(7) = oDataFarmaco1(1).data
            tInput.v_valori(8) = oDataFarmaco2(2).data
            tInput.v_valori(9) = oDataFarmaco3(3).data
        End If
    
    Case tpIRICOVERI
        tInput.v_valori(3) = txtNoteRicoveri & ""
        If lblData(1) <> "" Then
            tInput.v_valori(1) = lblData(1)
        End If
        If lblData(2) <> "" Then
            tInput.v_valori(2) = lblData(2)
        End If
    
    Case tpIPRESCRIZIONI
        If lblData(7) = "" Or lblData(8) = "" Then
            MsgBox "Inserire le date", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If cboCodicePrestazione.ListIndex = -1 Then
            MsgBox "Selezionare la prescrizione", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If CDate(lblData(7)) > CDate(lblData(8)) Then
            MsgBox "Inserire le date correttamente", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If CDate(lblData(8)) < CDate(frmPrescrizioni.oData(0).data) Then
            MsgBox "La data di fine prestazione non può essere antecedente alla data prenotazione", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If CDate(lblData(8)) < CDate(frmPrescrizioni.oData(1).data) Then
            MsgBox "La data di fine prestazione non può essere antecedente alla data ricetta", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        End If
        tInput.v_valori(1) = cboPrescrizioni.ItemData(cboPrescrizioni.ListIndex)
        tInput.v_valori(2) = txtQuantita
        tInput.v_valori(3) = lblData(7)
        tInput.v_valori(4) = lblData(8)
    
    Case tpIEPISODI
        tInput.v_valori(1) = lblData(3)
        tInput.v_valori(2) = txtNoteEpisodi & ""
    
    Case tpITRASFUSIONI
        If cboTrasfusioni.Text <> "" Then
            Call GestisciNuovo("TIPO_TRASFUSIONI", cboTrasfusioni)
        End If
        tInput.v_valori(1) = lblData(4)
        tInput.v_valori(2) = cboTrasfusioni.ItemData(cboTrasfusioni.ListIndex)
    
    Case tpISIEROCONVERSIONI
        tInput.v_valori(1) = lblData(5)
        tInput.v_valori(2) = IIf(chkSieroconversioni(0).Value = Checked, True, False)
        tInput.v_valori(3) = IIf(chkSieroconversioni(1).Value = Checked, True, False)
    
    Case tpICOLTURE
        tInput.v_valori(1) = lblData(6)
        tInput.v_valori(2) = txtColtureAcqua
        tInput.v_valori(3) = GestisciOptTre(optEsitoAcqua)
        tInput.v_valori(4) = txtColtureBagno
        tInput.v_valori(5) = GestisciOptTre(optEsitoBagno)
    
    Case tpITERAPIESTRAORDINARIE
        If cboMedicinaliStraordinaria.ListIndex = -1 Then
            MsgBox "Selezionare il farmaco", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        Else
            tInput.v_valori(1) = cboMedicinaliStraordinaria.ListIndex
        End If
        tInput.v_valori(2) = txtPosologiaStraordinaria
        tInput.v_valori(3) = IIf(chkConfermaStraordinaria.Value = Checked, 1, 0)
        tInput.v_valori(4) = txtNoteStraordinarie
    
    Case tpIDISTRETTI
        If cboAslAppartenenza.ListIndex = -1 Or txtCodiceDistretto = "" Or txtNomeDistretto = "" Then
            MsgBox "Inserire tutti i valori", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        End If
        tInput.v_valori(1) = txtCodiceDistretto
        tInput.v_valori(2) = txtNomeDistretto
        tInput.v_valori(3) = cboAslAppartenenza.ItemData(cboAslAppartenenza.ListIndex)
    
    Case tpINOMENCLATORE
        If txtCodicePrestazione = "" Or txtNomePrestazione = "" Or txtImporto = "" Or txtImportoScontato = "" Then
            MsgBox "Inserire tutti i valori", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        End If
        tInput.v_valori(1) = txtCodicePrestazione
        tInput.v_valori(2) = txtNomePrestazione
        tInput.v_valori(3) = IIf(txtImporto = "", "0.00", txtImporto)
        tInput.v_valori(4) = IIf(txtImportoScontato = "", "0.00", txtImportoScontato)
    
    Case tpICOMUNI
        If cboRegComuni.ListIndex = -1 Or txtCodIstat = "" Or txtComune = "" Then
            MsgBox "Inserire tutti i valori", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If Not Len(txtCodIstat) = 6 Then
            MsgBox "Il codice ISTAT deve essere di 6 caratteri", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        tInput.v_valori(1) = txtCodIstat
        tInput.v_valori(2) = txtComune
        tInput.v_valori(3) = cboRegComuni.ItemData(cboRegComuni.ListIndex)
    
    Case tpIASL
        If cboRegione.ListIndex = -1 Or txtCodiceAsl = "" Or txtNomeAsl = "" Then
            MsgBox "Inserire tutti i valori", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        tInput.v_valori(1) = txtCodiceAsl
        tInput.v_valori(2) = txtNomeAsl
        tInput.v_valori(3) = cboRegione.ItemData(cboRegione.ListIndex)
    
    Case tpIESENZIONE
        tInput.v_valori(1) = txtCodiceEsenzione
        If txtCodiceEsenzione = "" Then
            MsgBox "Inserire il codice di esenzione", vbCritical, "ATTENZIONE!!!!!!"
            Exit Sub
        End If
        If optTicketRicetta(0).Value Then
            tInput.v_valori(2) = True
        Else
            tInput.v_valori(2) = False
        End If

    End Select
    Unload Me
End Sub

Private Sub cboAslAppartenenza_LostFocus()
    cmdInserisci.SetFocus
End Sub

Private Sub cboRegione_LostFocus()
    cmdInserisci.SetFocus
End Sub

Private Sub cboTrasfusioni_LostFocus()
    cmdInserisci.SetFocus
End Sub

Private Sub chkGiorni_LostFocus(Index As Integer)
    cmdInserisci.SetFocus
End Sub

Private Sub chkSieroconversioni_LostFocus(Index As Integer)
    cmdInserisci.SetFocus
End Sub

Private Sub chkStampaVoce_LostFocus()
    cmdInserisci.SetFocus
End Sub

Private Sub chkGiorni_Click(Index As Integer)
    Dim i As Integer
    Static cambio As Boolean
    If Index = 7 Then
        If chkGiorni(7).Value = Checked Then
            cambio = True
            For i = 0 To 6
                chkGiorni(i).Value = Unchecked
            Next i
            cambio = False
        End If
    Else
        If Not cambio Then chkGiorni(7).Value = Unchecked
    End If
End Sub

Private Sub cboCodicePrestazione_Click()
    cboPrescrizioni.ListIndex = getIndexNomePrestazione
End Sub

Private Sub cboPrescrizioni_Click()
    cboCodicePrestazione.ListIndex = getIndexCodicePrestazione
End Sub

Private Sub cboPrescrizioni_DropDown()
    Call SetComboWidth(cboPrescrizioni, 500)
End Sub

Private Sub cboTrasfusioni_DropDown()
    Call SetComboWidth(cboTrasfusioni, 300)
End Sub

Private Sub cboMedicinali_DropDown()
    Call SetComboWidth(cboMedicinali, 300)
End Sub

Private Sub cboVoci_DropDown()
    Call SetComboWidth(cboVoci, 300)
End Sub

Private Sub chkPN_Click()
    If chkPN.Value = Checked Then
        txtUnita = ""
        txtValoreMax = ""
        txtValoreMin = ""
        txtUnita.Enabled = False
        txtValoreMax.Enabled = False
        txtValoreMin.Enabled = False
    Else
        txtUnita.Enabled = True
        txtValoreMax.Enabled = True
        txtValoreMin.Enabled = True
    End If
End Sub

Private Sub lblData_Click(Index As Integer)
    lblData(Index) = ""
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

Private Sub optEsitoBagno_LostFocus(Index As Integer)
    cmdInserisci.SetFocus
End Sub

Private Sub optMansione_Click(Index As Integer)
    Call ColoraSel(optMansione, Index, 2)
End Sub

Private Sub optSomministrazione_Click(Index As Integer)
    Call ColoraSel(optSomministrazione, Index, 1)
End Sub

Private Sub optTicketRicetta_LostFocus(Index As Integer)
    cmdInserisci.SetFocus
End Sub

Private Sub optTipoRene_Click(Index As Integer)
    Call ColoraSel(optTipoRene, Index, 3)
End Sub

Private Sub optTicketRicetta_Click(Index As Integer)
    Call ColoraSel(optTicketRicetta, Index, 2)
End Sub

Private Sub optUtente_Click(Index As Integer)
    Call ColoraSel(optUtente, Index, 3)
End Sub

Private Sub optUtente_LostFocus(Index As Integer)
    cmdInserisci.SetFocus
End Sub

Private Sub picData_Click(Index As Integer)
    frmCalendario.Show 1
    If laData <> "" Then
        lblData(Index) = laData
        If Index = 7 And txtQuantita = 1 Then
            lblData(Index + 1) = laData
        End If
    End If
End Sub

Private Sub picData_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub txtChiave_GotFocus()
    txtChiave.SelStart = 0
    txtChiave.SelLength = Len(txtChiave.Text)
    txtChiave.BackColor = colArancione
End Sub

Private Sub txtNoteStraordinarie_LostFocus()
    cmdInserisci.SetFocus
End Sub

Private Sub txtSomministrazione_GotFocus()
    txtSomministrazione.BackColor = colArancione
End Sub

Private Sub txtSomministrazione_LostFocus()
    txtSomministrazione.BackColor = vbWhite
End Sub

Private Sub txtPosologia_GotFocus()
    txtPosologia.BackColor = colArancione
End Sub

Private Sub txtPosologia_LostFocus()
    txtPosologia.BackColor = vbWhite
End Sub

Private Sub txtQuantita_GotFocus()
    txtQuantita.BackColor = colArancione
End Sub

Private Sub txtQuantita_LostFocus()
    txtQuantita.BackColor = vbWhite
End Sub

Private Sub txtChiave_LostFocus()
    txtChiave.BackColor = vbWhite
End Sub

Private Sub txtCognomePass_GotFocus()
    txtCognomePass.BackColor = colArancione
End Sub

Private Sub txtCognomePass_LostFocus()
    txtCognomePass.BackColor = vbWhite
End Sub

Private Sub txtColtureAcqua_GotFocus()
    txtColtureAcqua.BackColor = colArancione
End Sub

Private Sub txtColtureAcqua_LostFocus()
    txtColtureAcqua.BackColor = vbWhite
End Sub

Private Sub txtColtureBagno_GotFocus()
    txtColtureBagno.BackColor = colArancione
End Sub

Private Sub txtColtureBagno_LostFocus()
    txtColtureBagno.BackColor = vbWhite
End Sub

Private Sub txtMatricola_GotFocus()
    txtMatricola.BackColor = colArancione
End Sub

Private Sub txtMatricola_LostFocus()
    txtMatricola.BackColor = vbWhite
    cmdInserisci.SetFocus
End Sub

Private Sub txtNome_GotFocus()
    txtNome.BackColor = colArancione
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = vbWhite
End Sub

Private Sub txtNomePass_GotFocus()
    txtNomePass.BackColor = colArancione
End Sub

Private Sub txtNomePass_LostFocus()
    txtNomePass.BackColor = vbWhite
End Sub

Private Sub txtNoteEpisodi_GotFocus()
    txtNoteEpisodi.BackColor = colArancione
End Sub

Private Sub txtNoteEpisodi_LostFocus()
    txtNoteEpisodi.BackColor = vbWhite
    cmdInserisci.SetFocus
End Sub

Private Sub txtNoteRicoveri_GotFocus()
    txtNoteRicoveri.BackColor = colArancione
End Sub

Private Sub txtNoteRicoveri_LostFocus()
    txtNoteRicoveri.BackColor = vbWhite
    cmdInserisci.SetFocus
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
    txtPass.BackColor = colArancione
End Sub

Private Sub txtPass_LostFocus()
    txtPass.BackColor = vbWhite
End Sub

Private Sub txtNumeroRene_Change()
    Call OnlyNumber(txtNumeroRene, lettera)
End Sub

Private Sub txtQuantita_Change()
    Call OnlyNumber(txtQuantita, lettera)
End Sub

Private Sub txtPostazione_GotFocus()
    txtpostazione.BackColor = colArancione
End Sub

Private Sub txtPostazione_LostFocus()
    txtpostazione.BackColor = vbWhite
End Sub

Private Sub txtNumeroRene_GotFocus()
    txtNumeroRene.BackColor = colArancione
End Sub

Private Sub txtNumeroRene_LostFocus()
    txtNumeroRene.BackColor = vbWhite
End Sub

Private Sub txtQuantita_Validate(Cancel As Boolean)
    If txtQuantita <> "" Then
        If txtQuantita < 1 Or txtQuantita > 99 Then
            MsgBox "Inserire un valore compreso tra 1 e 99", vbCritical, "ATTENZIONE!!!"
            Cancel = True
        Else
            Cancel = False
        End If
    Else
        Cancel = False
    End If
End Sub

Private Sub txtTipoRene_GotFocus()
    txtTipoRene.BackColor = colArancione
End Sub

Private Sub txtTipoRene_LostFocus()
    txtTipoRene.BackColor = vbWhite
End Sub

Private Sub txtUnita_GotFocus()
    txtUnita.BackColor = colArancione
End Sub

Private Sub txtUnita_LostFocus()
    txtUnita.BackColor = vbWhite
End Sub

Private Sub txtNomeDistretto_GotFocus()
    txtNomeDistretto.BackColor = colArancione
End Sub

Private Sub txtNomeDistretto_LostFocus()
    txtNomeDistretto.BackColor = vbWhite
End Sub

Private Sub txtCodiceDistretto_GotFocus()
    txtCodiceDistretto.BackColor = colArancione
End Sub

Private Sub txtCodiceDistretto_LostFocus()
    txtCodiceDistretto.BackColor = vbWhite
End Sub

Private Sub txtCodIstat_GotFocus()
    txtCodIstat.BackColor = colArancione
End Sub

Private Sub txtCodIstat_LostFocus()
    txtCodIstat.BackColor = vbWhite
End Sub

Private Sub txtComune_LostFocus()
    txtComune.BackColor = vbWhite
End Sub

Private Sub txtComune_GotFocus()
    txtComune.BackColor = colArancione
End Sub

Private Sub txtCodiceAsl_LostFocus()
    txtCodiceAsl.BackColor = vbWhite
End Sub

Private Sub txtCodiceAsl_GotFocus()
    txtCodiceAsl.BackColor = colArancione
End Sub

Private Sub txtNomeAsl_LostFocus()
    txtNomeAsl.BackColor = vbWhite
End Sub

Private Sub txtNomeAsl_GotFocus()
    txtNomeAsl.BackColor = colArancione
End Sub

Private Sub txtNomePrestazione_GotFocus()
    txtNomePrestazione.BackColor = colArancione
End Sub

Private Sub txtNomePrestazione_LostFocus()
    txtNomePrestazione.BackColor = vbWhite
End Sub

Private Sub txtCodicePrestazione_GotFocus()
    txtCodicePrestazione.BackColor = colArancione
End Sub

Private Sub txtCodicePrestazione_LostFocus()
    txtCodicePrestazione.BackColor = vbWhite
End Sub

Private Sub txtCodiceEsenzione_GotFocus()
    txtCodiceEsenzione.BackColor = colArancione
End Sub

Private Sub txtCodiceEsenzione_LostFocus()
    txtCodiceEsenzione.BackColor = vbWhite
End Sub

Private Sub txtImportoScontato_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtImportoScontato, lettera)
    End If
End Sub

Private Sub txtImportoScontato_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtImportoScontato_GotFocus()
    txtImportoScontato.BackColor = colArancione
End Sub

Private Sub txtImportoScontato_LostFocus()
    txtImportoScontato.BackColor = vbWhite
    cmdInserisci.SetFocus
End Sub

Private Sub txtImportoScontato_Validate(Cancel As Boolean)
    If txtImporto = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtImportoScontato.Text)
    End If
End Sub

Private Sub txtImporto_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtImporto, lettera)
    End If
End Sub

Private Sub txtImporto_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtImporto_GotFocus()
    txtImporto.BackColor = colArancione
End Sub

Private Sub txtImporto_LostFocus()
    txtImporto.BackColor = vbWhite
End Sub

Private Sub txtImporto_Validate(Cancel As Boolean)
    If txtImporto = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtImporto.Text)
    End If
End Sub

Private Sub txtValoreMax_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtValoreMax, lettera)
    End If
End Sub

Private Sub txtValoreMax_GotFocus()
    txtValoreMax.BackColor = colArancione
End Sub

Private Sub txtValoreMax_LostFocus()
    txtValoreMax.BackColor = vbWhite
End Sub

Private Sub txtValoreMax_Validate(Cancel As Boolean)
    If txtValoreMax = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtValoreMax.Text)
    End If
End Sub

Private Sub txtValoreMin_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtValoreMin, lettera)
    End If
End Sub

Private Sub txtValoreMin_GotFocus()
    txtValoreMin.BackColor = colArancione
End Sub

Private Sub txtValoreMin_LostFocus()
    txtValoreMin.BackColor = vbWhite
End Sub

Private Sub txtValoreMin_Validate(Cancel As Boolean)
    If txtValoreMin = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtValoreMin.Text)
    End If
End Sub

Private Sub txtVoce_GotFocus()
    txtVoce.BackColor = colArancione
End Sub

Private Sub txtVoce_LostFocus()
    txtVoce.BackColor = vbWhite
End Sub

' Gestione dei tab e degli invio

Private Sub txtPosologia_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtQuantita_KeyPress(KeyAscii As Integer)
    lettera = Chr(KeyAscii)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtPostazione_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNumeroRene_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtValoreMin_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtValoreMax_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub cmdAnnulla_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        Select Case tInput.Tipo
            Case tpICOMPOSTO, tpISINGOLO, tpIESAMI:
                txtCognome.SetFocus
            Case tpIVOCI:
                txtVoce.SetFocus
            Case tpITIPIESAMILAB
                cboVoci.SetFocus
            Case tpIPASSWORD
                txtCognomePass.SetFocus
            Case tpITERAPIADOMICILIARE Or tpITERAPIESTRAORDINARIE
                cboMedicinali.SetFocus
 '           Case tpIRENI
 '               txtTipoRene.SetFocus
            Case tpICOLTURE
                txtColtureAcqua.SetFocus
        End Select
    End If
End Sub

Private Sub chkSieroconversioni_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub optEsitoAcqua_KeyPress(Index As Integer, KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub optMansione_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub optEsitoBagno_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub cboTrasfusioni_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub optUtente_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtChiave_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtcogNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCognomePass_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtMatricola_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab) Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtNomePass_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNoteEpisodi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtNoteRicoveri_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        cmdInserisci.SetFocus
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Function ControlloDuplicato(strNomeTabella As String) As Boolean
    Dim rsDataset As New Recordset
    Dim strSql As String

    strSql = "Select    count(Key) as Totale " & _
            "From " & strNomeTabella & " " & _
            "Where      Cognome like '" & Apostrophe(UCase(txtCognomePass.Text)) & "' and" & _
            "           Nome like '" & Apostrophe(UCase(txtNomePass.Text)) & "' and Tipo= " & GestisciOpt & ""

    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly
    
    If rsDataset("Totale") = 0 Then
        ControlloDuplicato = False
    Else
        rsDataset.Close
         strSql = "Select * " & _
            "From " & strNomeTabella & " " & _
            "Where      Cognome like '" & Apostrophe(UCase(txtCognomePass.Text)) & "' and" & _
            "           Nome like '" & Apostrophe(UCase(txtNomePass.Text)) & "'and Tipo= " & GestisciOpt & ""
  
        rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic
        If rsDataset("Cognome") = UCase(txtCognomePass.Text) And rsDataset("Nome") = UCase(txtNomePass.Text) And rsDataset("Eliminato") = True Then
            If MsgBox(UCase(txtCognomePass.Text) & " " & UCase(txtNomePass.Text) & " è presente in archivio ma è disattivato. Vuoi riattivarlo?", vbYesNo + vbCritical + vbDefaultButton2, Me.Caption) = vbNo Then
                ControlloDuplicato = True
            Else
                rsDataset("Eliminato") = False
                rsDataset("Chiave") = txtChiave
                rsDataset("Password") = txtPass
                rsDataset.Update
                Unload Me
                ControlloDuplicato = False
            End If
        End If
    End If

    rsDataset.Close
    
    Set rsDataset = Nothing
End Function

