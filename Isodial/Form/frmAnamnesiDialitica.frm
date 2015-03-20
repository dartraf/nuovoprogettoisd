VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Object = "{EB7F7146-0A68-4457-8036-5793F0EB1EB8}#31.0#0"; "SuperTextBox.ocx"
Begin VB.Form frmAnamnesiDialitica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ANAMNESI SCHEDA DIALITICA"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleMode       =   0  'User
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   79
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmAnamnesiDialitica.frx":0000
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
         Top             =   360
         Width           =   465
      End
   End
   Begin TabDlg.SSTab tabSchede 
      Height          =   4335
      Left            =   120
      TabIndex        =   40
      Top             =   850
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   7646
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
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
      TabCaption(0)   =   "Scheda 1"
      TabPicture(0)   =   "frmAnamnesiDialitica.frx":0459
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRitmoDialitico"
      Tab(0).Control(1)=   "cboCodicePrestaione"
      Tab(0).Control(2)=   "picCaricaPesoSecco"
      Tab(0).Control(3)=   "txtSodio"
      Tab(0).Control(4)=   "txtGlucosio"
      Tab(0).Control(5)=   "cboTipoAgo(1)"
      Tab(0).Control(6)=   "cboTipoAgo(0)"
      Tab(0).Control(7)=   "txtCalcio"
      Tab(0).Control(8)=   "txtBicarbonato"
      Tab(0).Control(9)=   "txtMinuti"
      Tab(0).Control(10)=   "txtOre"
      Tab(0).Control(11)=   "txtAumentoPond"
      Tab(0).Control(12)=   "txtQuantita"
      Tab(0).Control(13)=   "txtPesoSecco"
      Tab(0).Control(14)=   "txtSedeAccesso"
      Tab(0).Control(15)=   "cboAccesso"
      Tab(0).Control(16)=   "cboTipoDialisi"
      Tab(0).Control(17)=   "txtPotassio"
      Tab(0).Control(18)=   "chkDiuresiResidua"
      Tab(0).Control(19)=   "cboTipoFiltro"
      Tab(0).Control(20)=   "picElenca(0)"
      Tab(0).Control(21)=   "picElenca(1)"
      Tab(0).Control(22)=   "cboTipoLinee"
      Tab(0).Control(23)=   "picElenca(2)"
      Tab(0).Control(24)=   "oData(0)"
      Tab(0).Control(25)=   "oData(1)"
      Tab(0).Control(26)=   "oData(2)"
      Tab(0).Control(27)=   "Label1(39)"
      Tab(0).Control(28)=   "Label1(38)"
      Tab(0).Control(29)=   "Label1(37)"
      Tab(0).Control(30)=   "Label1(36)"
      Tab(0).Control(31)=   "Label1(27)"
      Tab(0).Control(32)=   "Label1(17)"
      Tab(0).Control(33)=   "Label1(33)"
      Tab(0).Control(34)=   "Label1(21)"
      Tab(0).Control(35)=   "Label1(20)"
      Tab(0).Control(36)=   "Label1(2)"
      Tab(0).Control(37)=   "Label1(4)"
      Tab(0).Control(38)=   "Label1(5)"
      Tab(0).Control(39)=   "Label1(6)"
      Tab(0).Control(40)=   "Label1(7)"
      Tab(0).Control(41)=   "Label1(8)"
      Tab(0).Control(42)=   "Label1(9)"
      Tab(0).Control(43)=   "Label1(10)"
      Tab(0).Control(44)=   "Label1(18)"
      Tab(0).Control(45)=   "Label1(19)"
      Tab(0).Control(46)=   "Label1(22)"
      Tab(0).Control(47)=   "Label1(23)"
      Tab(0).Control(48)=   "Label1(24)"
      Tab(0).Control(49)=   "Label1(25)"
      Tab(0).ControlCount=   50
      TabCaption(1)   =   "Scheda 2"
      TabPicture(1)   =   "frmAnamnesiDialitica.frx":0475
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(16)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(14)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(13)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(26)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(28)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(29)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(30)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(31)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(35)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtDose(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtDose(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtDose(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtDose(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cboAnticoagulante(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboAnticoagulante(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFlusso"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cboSolDialitica"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cboSolInfusionale"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cboCartuccia"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtSolInfCc"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtFlussoSangue"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cboDosiUnitaMisura"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cboUnit‡ValoreInfusionale"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Scheda 3"
      TabPicture(2)   =   "frmAnamnesiDialitica.frx":0491
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(32)"
      Tab(2).Control(1)=   "lblUnitaMisura"
      Tab(2).Control(2)=   "Label1(34)"
      Tab(2).Control(3)=   "cboEPO"
      Tab(2).Control(4)=   "txtUI"
      Tab(2).Control(5)=   "txtNote"
      Tab(2).Control(6)=   "cmdEliminaEpo"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtRitmoDialitico 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -71100
         MaxLength       =   1
         TabIndex        =   96
         Top             =   480
         Width           =   300
      End
      Begin VB.ComboBox cboUnit‡ValoreInfusionale 
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
         ItemData        =   "frmAnamnesiDialitica.frx":04AD
         Left            =   9840
         List            =   "frmAnamnesiDialitica.frx":04B7
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   3390
         Width           =   615
      End
      Begin VB.ComboBox cboCodicePrestaione 
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
         ItemData        =   "frmAnamnesiDialitica.frx":04C3
         Left            =   -65160
         List            =   "frmAnamnesiDialitica.frx":04C5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   3360
         Width           =   1575
      End
      Begin VB.PictureBox picCaricaPesoSecco 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   -68520
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   87
         ToolTipText     =   "Carica Peso Secco"
         Top             =   960
         Width           =   360
      End
      Begin SuperTextBox.uSuperTextBox txtSodio 
         Height          =   285
         Left            =   -72120
         TabIndex        =   88
         Top             =   3840
         Width           =   615
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   5
      End
      Begin VB.TextBox txtGlucosio 
         Alignment       =   1  'Right Justify
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
         Left            =   -66720
         MaxLength       =   5
         TabIndex        =   15
         Top             =   3840
         Width           =   615
      End
      Begin VB.ComboBox cboDosiUnitaMisura 
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
         ItemData        =   "frmAnamnesiDialitica.frx":04C7
         Left            =   3240
         List            =   "frmAnamnesiDialitica.frx":04D1
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboTipoAgo 
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
         Left            =   -65040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cboTipoAgo 
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
         ItemData        =   "frmAnamnesiDialitica.frx":04DD
         Left            =   -67920
         List            =   "frmAnamnesiDialitica.frx":04DF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdEliminaEpo 
         Caption         =   "&Elimina EPO"
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
         Left            =   -69720
         TabIndex        =   33
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox txtFlussoSangue 
         Alignment       =   1  'Right Justify
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
         Left            =   7680
         MaxLength       =   5
         TabIndex        =   26
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCalcio 
         Alignment       =   1  'Right Justify
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
         Left            =   -68160
         MaxLength       =   5
         TabIndex        =   14
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtBicarbonato 
         Alignment       =   1  'Right Justify
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
         Left            =   -69345
         MaxLength       =   5
         TabIndex        =   13
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtMinuti 
         Alignment       =   1  'Right Justify
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
         Left            =   -63960
         MaxLength       =   2
         TabIndex        =   17
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtOre 
         Alignment       =   1  'Right Justify
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
         Left            =   -65160
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtSolInfCc 
         Alignment       =   1  'Right Justify
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
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   29
         Top             =   3390
         Width           =   615
      End
      Begin VB.ComboBox cboCartuccia 
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   3840
         Width           =   6255
      End
      Begin VB.ComboBox cboSolInfusionale 
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   3360
         Width           =   6255
      End
      Begin VB.ComboBox cboSolDialitica 
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox txtFlusso 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2400
         Width           =   615
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
         Height          =   2805
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   1320
         Width           =   11535
      End
      Begin VB.TextBox txtUI 
         Alignment       =   1  'Right Justify
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
         Left            =   -70560
         MaxLength       =   5
         TabIndex        =   32
         Top             =   495
         Width           =   735
      End
      Begin VB.ComboBox cboEPO 
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
         ItemData        =   "frmAnamnesiDialitica.frx":04E1
         Left            =   -73200
         List            =   "frmAnamnesiDialitica.frx":04F4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboAnticoagulante 
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
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   8055
      End
      Begin VB.ComboBox cboAnticoagulante 
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
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1440
         Width           =   8055
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
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
         Left            =   6840
         MaxLength       =   6
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   9840
         MaxLength       =   6
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtDose 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtAumentoPond 
         Alignment       =   1  'Right Justify
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
         Left            =   -64080
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtQuantita 
         Alignment       =   1  'Right Justify
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
         Left            =   -64080
         MaxLength       =   4
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPesoSecco 
         Alignment       =   1  'Right Justify
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
         Left            =   -72600
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtSedeAccesso 
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
         Left            =   -72600
         MaxLength       =   92
         TabIndex        =   10
         Top             =   2880
         Width           =   7935
      End
      Begin VB.ComboBox cboAccesso 
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
         Left            =   -72600
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2400
         Width           =   3735
      End
      Begin VB.ComboBox cboTipoDialisi 
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
         Left            =   -72600
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox txtPotassio 
         Alignment       =   1  'Right Justify
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
         Left            =   -70920
         MaxLength       =   5
         TabIndex        =   12
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox chkDiuresiResidua 
         Caption         =   "Diuresi Residua"
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
         Left            =   -70200
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cboTipoFiltro 
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
         ItemData        =   "frmAnamnesiDialitica.frx":051A
         Left            =   -72600
         List            =   "frmAnamnesiDialitica.frx":051C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   3735
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   -71895
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   51
         ToolTipText     =   "Elenca date"
         Top             =   930
         Width           =   360
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   -68760
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   52
         ToolTipText     =   "Elenca date"
         Top             =   1420
         Width           =   360
      End
      Begin VB.ComboBox cboTipoLinee 
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
         Left            =   -72600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   3735
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   -68760
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   53
         ToolTipText     =   "Elenca date"
         Top             =   1900
         Width           =   360
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   -70320
         TabIndex        =   89
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
         Left            =   -65400
         TabIndex        =   90
         Top             =   1440
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   2
         Left            =   -65400
         TabIndex        =   91
         Top             =   1920
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ritmo Dialisi Settimanale->Sedute"
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
         Index           =   39
         Left            =   -74760
         TabIndex        =   94
         Top             =   480
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Prestazione"
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
         Index           =   38
         Left            =   -67320
         TabIndex        =   93
         Top             =   3360
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gluc"
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
         Index           =   37
         Left            =   -67320
         TabIndex        =   86
         Top             =   3855
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ago A."
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
         Index           =   36
         Left            =   -68640
         TabIndex        =   39
         Top             =   2440
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ago V."
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
         Index           =   27
         Left            =   -65760
         TabIndex        =   38
         Top             =   2440
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flusso Sangue Qb (ml/min)"
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
         Index           =   35
         Left            =   4800
         TabIndex        =   78
         Top             =   2400
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ca"
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
         Left            =   -68520
         TabIndex        =   77
         Top             =   3840
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "HCO3-"
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
         Left            =   -70080
         TabIndex        =   76
         Top             =   3855
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minuti"
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
         Left            =   -64680
         TabIndex        =   75
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ore"
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
         Left            =   -65640
         TabIndex        =   74
         Top             =   3840
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartuccia"
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
         Index           =   31
         Left            =   240
         TabIndex        =   73
         Top             =   3870
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "valore"
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
         Index           =   30
         Left            =   9120
         TabIndex        =   72
         Top             =   3390
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Soluzione Infusionale"
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
         Index           =   29
         Left            =   240
         TabIndex        =   71
         Top             =   3390
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Soluzione Dialitica"
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
         Index           =   28
         Left            =   240
         TabIndex        =   70
         Top             =   2895
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flusso Dialisi Qd (ml/min)"
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
         Index           =   26
         Left            =   240
         TabIndex        =   69
         Top             =   2400
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   34
         Left            =   -74760
         TabIndex        =   68
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblUnitaMisura 
         AutoSize        =   -1  'True
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
         Left            =   -71160
         TabIndex        =   67
         Top             =   525
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eritropoietina"
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
         Left            =   -74760
         TabIndex        =   66
         Top             =   500
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anticoagulante"
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
         TabIndex        =   65
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose Iniziale"
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
         TabIndex        =   64
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose Intermedia"
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
         Left            =   4920
         TabIndex        =   63
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose Finale"
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
         Left            =   8400
         TabIndex        =   62
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Altro Anticoagulante"
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
         Left            =   240
         TabIndex        =   61
         Top             =   1470
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose"
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
         Left            =   240
         TabIndex        =   60
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantit‡ (ml/die)"
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
         Left            =   -65880
         TabIndex        =   44
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo Peso Secco"
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aumento Pond. Interdialitico (ml)"
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
         Left            =   -67560
         TabIndex        =   50
         Top             =   480
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di Filtro"
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
         Left            =   -74760
         TabIndex        =   46
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Accesso Vascolare"
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   2430
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sede di Accesso"
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
         Left            =   -74760
         TabIndex        =   59
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di Dialisi"
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
         Left            =   -74760
         TabIndex        =   58
         Top             =   3360
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bagno Dialisi"
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
         Left            =   -74760
         TabIndex        =   57
         Top             =   3840
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Na+"
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
         Left            =   -72600
         TabIndex        =   56
         Top             =   3840
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "K+"
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
         Left            =   -71280
         TabIndex        =   55
         Top             =   3840
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "in data"
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
         Left            =   -71160
         TabIndex        =   45
         Top             =   1005
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "in data"
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
         Left            =   -66225
         TabIndex        =   47
         Top             =   1485
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di linee"
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
         Index           =   24
         Left            =   -74760
         TabIndex        =   49
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "in data"
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
         Index           =   25
         Left            =   -66225
         TabIndex        =   48
         Top             =   1965
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   54
      Top             =   5040
      Width           =   12015
      Begin VB.CommandButton cmdStampaSintetica 
         Caption         =   "&Stampa Sintetica"
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
         Left            =   4560
         TabIndex        =   34
         Top             =   240
         Width           =   2175
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
         Left            =   7080
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   240
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
         Height          =   495
         Left            =   10560
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAnamnesiDialitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lettera As String
Dim rsAnamnesiDialitica As Recordset
Dim modifica As Boolean
Dim keyId As Integer            ' il key utilizzato in fase di modifica

Dim dataPeso As Date            ' data del peso caricato
Dim dataFiltro As Date          ' data del filtro caricato
Dim dataLinee As Date           ' data del linee caricato
Dim vecchioPeso As Single       ' il vecchio peso caricato
Dim vecchioFiltro As Integer    ' il vecchio filtro caricato
Dim vecchioLinee As Integer      ' il vecchio linee caricato
Dim epoAppo As Integer          ' epo di appoggio
Dim uiAppo As Long           ' ui di appogio

Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim blnModificato As Boolean


Private Sub Form_Activate()
    Dim blnModificatoAppo As Boolean
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    blnModificatoAppo = blnModificato
    ' ricarica le combo
    Call RicaricaComboBox("FILTRI", "NOME", cboTipoFiltro)
    Call RicaricaComboBox("LINEE", "NOME", cboTipoLinee)
    Call RicaricaComboBox("ANTICOAGULANTI", "NOME", cboAnticoagulante(0))
    Call RicaricaComboBox("ANTICOAGULANTI", "NOME", cboAnticoagulante(1))
    Call RicaricaComboBox("TIPI_DIALISI", "NOME", cboTipoDialisi)
    Call RicaricaComboBox("NOMENCLATORE_TARIFFARIO", "CODICE", cboCodicePrestaione)
    Call RicaricaComboBox("ACCESSI_VASCOLARI", "NOME", cboAccesso)
    Call RicaricaComboBox("AGO", "NOME", cboTipoAgo(0))
    Call RicaricaComboBox("AGO", "NOME", cboTipoAgo(1))
    Call RicaricaComboBox("SOL_DIALITICHE", "NOME", cboSolDialitica)
    Call RicaricaComboBox("SOL_INFUSIONALI", "NOME", cboSolInfusionale)
    Call RicaricaComboBox("CARTUCCE", "NOME", cboCartuccia)
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
    tabSchede.Tab = 0
    cboDosiUnitaMisura.ListIndex = 0
    
    For i = 0 To 2
        oData(i).ConnectionString = strConnectionStringCentro
        picElenca(i).Picture = LoadResPicture("storico1", 0)
    Next i
    picCaricaPesoSecco.Picture = LoadResPicture("PESO1", 0)
    
    Call ApriRsDisconnesso
    
    If structIntestazione.sCodiceSTS = "AD0082" Or (Environ$("COMPUTERNAME") = "MASTERMIO") Then
        cmdStampaSintetica.Visible = True
    End If
    
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

Private Sub ApriRsDisconnesso()
    ' apre il recordset disconnesso per la tracciatura
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ANAMNESI_DIALITICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
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
    Dim valAppo
    
    ReDim v_modifiche(0)
    For i = 0 To rsDisco.Fields.count - 1
        trovato = False
        If IsNull(rsDisco(i)) Or IsNull(rsAnamnesiDialitica(i)) Then
            If Not (IsNull(rsDisco(i)) And IsNull(rsAnamnesiDialitica(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(i) <> rsAnamnesiDialitica(i) Then
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
            If (IsNull(rsDisco.Fields((v_modifiche(i)))) Or rsDisco.Fields((v_modifiche(i))) = "") Then
                valAppo = "NULL"
            Else
                If IsNumeric(rsDisco.Fields((v_modifiche(i)))) And Not (rsDisco.Fields((v_modifiche(i))) = True Or rsDisco.Fields((v_modifiche(i))) = False) Then
                    valAppo = VirgolaOrPunto(rsDisco.Fields((v_modifiche(i))), ",")
                Else
                    valAppo = rsDisco.Fields((v_modifiche(i)))
                End If
            End If
            valori = valori & valAppo & "&-&"
            ' aggiorna il rsDisco
            rsDisco(v_modifiche(i)) = rsAnamnesiDialitica(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rsAnamnesiDialitica("KEY"), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_DIALITICHE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

Private Sub Completa()
    Dim i As Integer
    lettera = "0"
    If txtRitmoDialitico = "" Then txtRitmoDialitico = "0"
    If txtAumentoPond = "" Then txtAumentoPond = "0"
    If txtQuantita = "" Then txtQuantita = "0"
    If txtFlusso = "" Then txtFlusso = "0"
    If txtFlussoSangue = "" Then txtFlussoSangue = "0"
    If txtSolInfCc = "" Then txtSolInfCc = "0"
    If txtPesoSecco = "" Then txtPesoSecco = "0"
    If txtSodio.Text = "" Then txtSodio.Text = "0"
    If txtPotassio = "" Then txtPotassio = "0"
    If txtBicarbonato = "" Then txtBicarbonato = "0"
    If txtCalcio = "" Then txtCalcio = "0"
    If txtGlucosio = "" Then txtGlucosio = "0"
    For i = 0 To 3
        If txtDose(i) = "" Then txtDose(i) = "0"
    Next i
    If txtOre = "" Then txtOre = "0"
    If txtMinuti = "" Then txtMinuti = "0"
    If txtUI = "" Then txtUI = "0"
End Sub

Private Sub PredisponiDosi()
    If cboEPO.ListIndex <> -1 Then
        If cboDosiUnitaMisura.ListIndex = 0 Then
            laData = Day(date) & "/" & cboDosiUnitaMisura.ListIndex + 9 & "/" & Year(date)
            'Debug.Print "EPO: " & cboEPO.Text
            Call CaricaPso
        End If
    End If
End Sub

Private Function CompletoStorico() As Boolean
    Dim modPeso As Boolean          '  avverte se Ë stato modificato  il peso secco
    Dim modFiltro As Boolean        '  avverte se Ë stato modificato tipo di filtro
    Dim modLinee As Boolean         '  avverte se Ë stato modificato tipo di linee
    
    If txtPesoSecco <> vecchioPeso Then
        modPeso = True
    End If
    If cboTipoFiltro.ListIndex <> vecchioFiltro Then
        modFiltro = True
    End If
    If cboTipoLinee.ListIndex <> vecchioLinee Then
        modLinee = True
    End If
    
    If Not (modFiltro = False And modPeso = False And modLinee = False) Then
        If modFiltro Then
            If oData(1).data = "" Then
                MsgBox "Inserire la data dell'ultimo tipo di filtro", vbCritical, "ATTENZIONE!!!"
                CompletoStorico = False
                Exit Function
            End If
        End If
        If modPeso Then
            If oData(0).data = "" Then
                MsgBox "Inserire la data dell'ultimo peso secco", vbCritical, "ATTENZIONE!!!"
                CompletoStorico = False
                Exit Function
            End If
        End If
        If modLinee Then
            If oData(2).data = "" Then
                MsgBox "Inserire la data dell'ultimo tipo di linee", vbCritical, "ATTENZIONE!!!"
                CompletoStorico = False
                Exit Function
            End If
        End If
    End If
    CompletoStorico = True
End Function

Private Function GestisciStorico(numKey As Integer) As Integer
    On Error GoTo gestione
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim modPeso As Boolean          '  avverte se Ë stato modificato  il peso secco
    Dim modFiltro As Boolean        '  avverte se Ë stato modificato tipo di filtro
    Dim modLinee As Boolean         '  avverte se Ë stato modificato tipo di linee
        
    If txtPesoSecco <> vecchioPeso Then
        modPeso = True
    End If
    If cboTipoFiltro.ListIndex <> vecchioFiltro Then
        modFiltro = True
    End If
    If cboTipoLinee.ListIndex <> vecchioLinee Then
        modLinee = True
    End If
    
    If Not (modFiltro = False And modPeso = False And modLinee = False) Then
        If modFiltro Then
            v_Nomi = Array("KEY", "CODICE_SCHEDA", "DATA", "TIPO_FILTRO")
            v_Val = Array(GetNumero("STORICO_DIALISI_FILTRO"), numKey, dataFiltro, -1)
            If cboTipoFiltro.ListIndex <> -1 Then
                v_Val(3) = cboTipoFiltro.ItemData(vecchioFiltro)
            End If
            
            Set rsAnamnesiDialitica = New Recordset
            rsAnamnesiDialitica.Open "STORICO_DIALISI_FILTRO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAnamnesiDialitica.AddNew v_Nomi, v_Val
            rsAnamnesiDialitica.Update
            Set rsAnamnesiDialitica = Nothing
        End If
        If modPeso Then
            v_Nomi = Array("KEY", "CODICE_SCHEDA", "PESO", "DATA")
            v_Val = Array(GetNumero("STORICO_DIALISI_PESO"), numKey, vecchioPeso, dataPeso)
            Set rsAnamnesiDialitica = New Recordset
            rsAnamnesiDialitica.Open "STORICO_DIALISI_PESO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAnamnesiDialitica.AddNew v_Nomi, v_Val
            rsAnamnesiDialitica.Update
            Set rsAnamnesiDialitica = Nothing
        End If
        If modLinee Then
            v_Nomi = Array("KEY", "CODICE_SCHEDA", "DATA", "TIPO_LINEE")
            v_Val = Array(GetNumero("STORICO_DIALISI_LINEE"), numKey, dataLinee, -1)
            If cboTipoLinee.ListIndex <> -1 Then
                v_Val(3) = cboTipoLinee.ItemData(vecchioLinee)
            End If
            
            Set rsAnamnesiDialitica = New Recordset
            rsAnamnesiDialitica.Open "STORICO_DIALISI_LINEE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsAnamnesiDialitica.AddNew v_Nomi, v_Val
            rsAnamnesiDialitica.Update
            Set rsAnamnesiDialitica = Nothing
        End If
        GestisciStorico = 1
    Else
        GestisciStorico = 0
    End If
    
    Exit Function
gestione:
    MsgBox "Descrizione: Valore non valido", vbCritical, "Errore n∞: " & Err.Number
    cnPrinc.RollbackTrans
    GestisciStorico = -1
End Function

Private Sub PulisciTutto()
    ' pulisce l'intera scheda
    modifica = False
    Call PulisciForm(Me)
    intPazientiKey = 0
    oData(0).Pulisci
    oData(1).Pulisci
    oData(2).Pulisci
    cboDosiUnitaMisura.ListIndex = 0
    cboUnit‡ValoreInfusionale.ListIndex = 0
    chkDiuresiResidua.Value = Unchecked
    cmdTrova.SetFocus
    keyId = 0
    'lblRitmoDialisiSettimanale.Caption = ""
    blnModificato = False
End Sub

Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If txtPesoSecco = "" Then
        MsgBox "Inserire il peso secco", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If oData(0).data = "" Then
        MsgBox "Inserire la data del peso secco", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboTipoFiltro.ListIndex = -1 Then
        MsgBox "Selezionare il tipo di filtro", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If oData(1).data = "" Then
        MsgBox "Inserire la data del tipo di filtro", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboTipoLinee.ListIndex = -1 Then
        MsgBox "Inserire il tipo di linee", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If oData(2).data = "" Then
        MsgBox "Inserire la data del tipo di linee", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboAccesso.ListIndex = -1 Then
        MsgBox "Inserire il tipo di accesso vascolare", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboTipoAgo(0).ListIndex = -1 Then
        MsgBox "Inserire il tipo di ago arterioso", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboTipoAgo(1).ListIndex = -1 Then
        MsgBox "Inserire il tipo di ago venoso", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    If cboTipoDialisi.ListIndex = -1 Then
        MsgBox "Inserire il tipo di dialisi", vbCritical, "ATTENZIONE!!!"
        Exit Function
    End If
    Completo = True
End Function

Private Sub cmdStampaSintetica_Click()
    
If structIntestazione.sCodiceSTS = CODICESTS_SANT_ANDREA Or structIntestazione.sCodiceSTS = CODICESTS_SODAV Then
    
    Dim strSqlStampa As String
    Dim strSql As String
    Dim i As Integer
    'Dim intNumeroGiorni As Integer
    Dim strGiorni As String
    Dim intCodiceID As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "ATTENZIONE!!!"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "Memorizzare prima la scheda", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    End If
    
    strSqlStampa = "    NEW adVarChar (50) as COGNOME, " & _
                "       NEW adVarChar (50) as NOME, " & _
                "       NEW adVarChar(10) AS NATO_NATA, " & _
                "       NEW adDate as NATO_IL, " & _
                "       NEW adVarChar (100) AS LUOGO_NASCITA, " & _
                "       NEW adVarChar (100) as RESIDENZA, " & _
                "       NEW adVarChar (30) AS TELEFONO, " & _
                "       NEW adVarChar (30) AS CELLULARE, " & _
                "       NEW adVarChar (16) as CODICE_FISCALE, " & _
                "       NEW adVarChar (200) as AFFETTO_IRC, " & _
                "       NEW adDate as TERAPIA_DIALITCA_DAL, " & _
                "       NEW adVarChar (20) as EMOGRUPPO, " & _
                "       NEW adVarChar (10) as DATA_HBSAG, " & _
                "       NEW adVarChar (10) as DATA_HBSAB, " & _
                "       NEW adVarChar (10) as DATA_HBEAG, " & _
                "       NEW adVarChar (10) as DATA_HBEAB, " & _
                "       NEW adVarChar (10) as DATA_HBCAB, "
   strSqlStampa = strSqlStampa & _
                "       NEW adVarChar (10) as HBSAG, " & _
                "       NEW adVarChar (10) as HBSAB, " & _
                "       NEW adVarChar (10) as HBEAG, " & _
                "       NEW adVarChar (10) as HBEAB, " & _
                "       NEW adVarChar (10) as HBCAB, " & _
                "       NEW adVarChar (10) as DATA_HCV, " & _
                "       NEW adVarChar (10) as DATA_HCVRNA, " & _
                "       NEW adVarChar (10) as HCVAB, " & _
                "       NEW adVarChar (10) as HCVRN, " & _
                "       NEW adVarChar (10) as IMMUNITA_HIV, " & _
                "       NEW adVarChar (10) as HIV, " & _
                "       NEW adVarChar (10) as RITMO_SETTIMANALE, " & _
                "       NEW adVarChar (100) as GIORNI_DIALISI, "
    strSqlStampa = strSqlStampa & _
                "       NEW adVarChar (50) as DURATA_SEDUTA, " & _
                "       NEW adVarChar (50) as METODICA_DIALITICA, " & _
                "       NEW adVarChar (50) as ACCESSO_VASCOLARE, " & _
                "       NEW adVarChar (50) as SEDE_ACCESSO, " & _
                "       NEW adVarChar (100) as NUMERO_AGHI, " & _
                "       NEW adVarChar (50) as MONITOR_EMODIALISI, " & _
                "       NEW adVarChar (50) as DIALIZZATORE, " & _
                "       NEW adSingle as SODIO, " & _
                "       NEW adSingle as POTASSIO, " & _
                "       NEW adSingle as CALCIO, " & _
                "       NEW adSingle as GLUCOSIO, " & _
                "       NEW adVarChar (50) as EPARINIZZAZIONE, " & _
                "       NEW adVarChar (20) as DOSE_INIZIALE, " & _
                "       NEW adSingle as DOSE_INTERMEDIA, " & _
                "       NEW adSingle as QB, " & _
                "       NEW adSingle as QD, "
    strSqlStampa = strSqlStampa & _
                "       NEW adSingle  as PESO_SECCO, " & _
                "       NEW adDate as ULTIMA_DIALISI_DEL, " & _
                "       NEW adVarChar (50) as INTOLLERANZA_FARMACO, " & _
                "       NEW adVarChar (10) as PESO_PRE_DIALISI, " & _
                "       NEW adVarChar (10) as PESO_POST_DIALISI, " & _
                "       NEW adVarChar (10) as PRESS_ARTER_PRE_DIAL, " & _
                "       NEW adVarChar (10) as PRESS_ARTER_POST_DIAL, " & _
                "       NEW adLongVarChar as COMPLICANZE_INTRADIALITICHE, " & _
                "       NEW adLongVarChar as FARMACO_TERAPIA_POST, " & _
                "       NEW adLongVarChar as POS_TERAPIA_POST, " & _
                "       NEW adLongVarChar as GIORNI_TERAPIA_POST, " & _
                "       NEW adLongVarChar as NOTE_TERAPIA_POST, " & _
                "       NEW adLongVarChar as FARMACO_TERAPIA_DOMICILIARE, " & _
                "       NEW adLongVarChar as POS_TERAPIA_DOMICILIARE, " & _
                "       NEW adLongVarChar as GIORNI_TERAPIA_DOMICILIARE, " & _
                "       NEW adLongVarChar as NOTE_TERAPIA_DOMICILIARE "

      
                
    ' stringa di shape
    strSqlStampa = "SHAPE APPEND " & strSqlStampa
     
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    ' carica il recordset padre
    Set rsDataset = New Recordset
        
    With rsMain
        
        ' Dati di anamnesi dialitica e nefrologica
        strSql = "SELECT * " & _
            " FROM ((((((((((ANAMNESI_DIALITICHE AN_D " & _
            "       LEFT OUTER JOIN ANAMNESI_NEFROLOGICHE AN_N ON AN_N.CODICE_PAZIENTE=AN_D.CODICE_PAZIENTE) " & _
            "       LEFT OUTER JOIN EDTA EDTA ON EDTA.KEY=AN_N.CODICE_EDTA) " & _
            "       LEFT OUTER JOIN PAZIENTI P ON P.KEY=AN_D.CODICE_PAZIENTE) " & _
            "       LEFT OUTER JOIN COMUNI C ON P.CODICE_COMUNE_RESIDENZA=C.KEY) " & _
            "       LEFT OUTER JOIN FILTRI F ON F.KEY=AN_D.TIPO_FILTRO) " & _
            "       LEFT OUTER JOIN ANTICOAGULANTI ANT ON ANT.KEY=AN_D.ANTICOAGULANTE1) " & _
            "       LEFT OUTER JOIN ACCESSI_VASCOLARI ACC ON ACC.KEY=AN_D.ACCESSO_VASCOLARE) " & _
            "       LEFT OUTER JOIN AGO AGO1 ON AGO1.KEY=AN_D.AGO1) " & _
            "       LEFT OUTER JOIN AGO AGO2 ON AGO2.KEY=AN_D.AGO2) " & _
            "       LEFT OUTER JOIN TIPI_DIALISI ON TIPI_DIALISI.KEY=AN_D.TIPO_DIALISI) " & _
            " Where AN_D.CODICE_PAZIENTE = " & intPazientiKey
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            .AddNew
            intCodiceID = rsDataset("CODICE_ID")
            .Fields("COGNOME") = lblCognome.Caption
            .Fields("NOME") = lblNome.Caption
            .Fields("NATO_NATA") = IIf(rsDataset("SESSO") = "M", "Nato il", "Nata il")
            .Fields("NATO_IL") = rsDataset("DATA_NASCITA")
            .Fields("LUOGO_NASCITA") = rsDataset("CITTA_NASCITA")
            .Fields("RESIDENZA") = rsDataset("C.NOME")
            .Fields("TELEFONO") = rsDataset("TELEFONO")
            .Fields("CELLULARE") = rsDataset("CELLULARE")
            .Fields("CODICE_FISCALE") = rsDataset("CODICE_FISCALE")
            .Fields("AFFETTO_IRC") = rsDataset("EDTA.NOME")  'Right(rsDataset("EDTA.NOME"), Len(rsDataset("EDTA.NOME")) - 5)
            .Fields("TERAPIA_DIALITCA_DAL") = rsDataset("DATA1")
            If rsDataset("G_SANGUIGNO") <> -1 Then
                If rsDataset("RH") <> -1 Then
                    .Fields("EMOGRUPPO") = Choose(rsDataset("G_SANGUIGNO") + 1, "A", "B", "AB", "0") & " " & Choose(rsDataset("RH") + 1, "POSITIVO", "NEGATIVO")
                Else
                    .Fields("EMOGRUPPO") = Choose(rsDataset("G_SANGUIGNO") + 1, "A", "B", "AB", "0")
                End If
            Else
                .Fields("EMOGRUPPO") = "- -"
            End If
            
            If rsDataset("RITMO_DIALITICO") = 0 Then
                .Fields("RITMO_SETTIMANALE") = "- -"
            Else
                .Fields("RITMO_SETTIMANALE") = rsDataset("RITMO_DIALITICO") & " Sedute"
            End If
            
            .Fields("METODICA_DIALITICA") = rsDataset("TIPI_DIALISI.NOME")
            .Fields("DIALIZZATORE") = rsDataset("F.NOME")
            .Fields("SODIO") = rsDataset("SODIO")
            .Fields("POTASSIO") = rsDataset("POTASSIO")
            .Fields("CALCIO") = rsDataset("CALCIO")
            .Fields("GLUCOSIO") = rsDataset("GLUCOSIO")
            .Fields("EPARINIZZAZIONE") = rsDataset("ANT.NOME")
            .Fields("DOSE_INIZIALE") = rsDataset("DOSE1") & " " & cboDosiUnitaMisura.Text
            .Fields("DOSE_INTERMEDIA") = rsDataset("DOSE2")
            .Fields("QB") = rsDataset("FLUSSO_SANGUE")
            .Fields("QD") = rsDataset("FLUSSO")
            .Fields("ACCESSO_VASCOLARE") = rsDataset("ACC.NOME")
            .Fields("SEDE_ACCESSO") = rsDataset("SEDE_ACCESSO")
            .Fields("NUMERO_AGHI") = rsDataset("AGO1.NOME") & " " & rsDataset("AGO2.NOME")
            .Fields("DURATA_SEDUTA") = rsDataset("ORE") & " ore e " & rsDataset("MINUTI") & " minuti"
            .Fields("PESO_SECCO") = rsDataset("PESO_SECCO")
            .Fields("INTOLLERANZA_FARMACO") = rsDataset("ALLERGIA")
            
        End If
        rsDataset.Close
        
        strSql = "Select    Top 1 * " & _
                 "From      DIARI_CLINICI " & _
                 "          INNER JOIN TITOLI_DIARIO ON TITOLI_DIARIO.KEY=DIARI_CLINICI.CODICE_TITOLO " & _
                 "Where     CODICE_PAZIENTE=" & intPazientiKey & " AND TITOLI_DIARIO.NOME LIKE '%COMPLICANZE%' " & _
                 "Order By  DATA DESC"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            .Fields("COMPLICANZE_INTRADIALITICHE") = rsDataset("DATI")
        Else
            .Fields("COMPLICANZE_INTRADIALITICHE") = "- -"
        End If
        rsDataset.Close
        
        ' Dati degli esami di lab
        strSql = "Select    top 1 * " & _
                 "From      ((ANAMNESI_ESAMI AN " & _
                 "          INNER JOIN ESAMI_LAB ES ON ES.CODICE_ANAMNESI_ESAMI=AN.KEY) " & _
                 "          INNER JOIN VOCI_ESAMI V ON V.KEY=ES.CODICE_ESAME) " & _
                 "Where     AN.CODICE_PAZIENTE=" & intPazientiKey & _
                 "          AND V.NOME LIKE "
        
        rsDataset.Open strSql & "'%HBSAG%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HBSAG") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HBSAG") = "POS."
            Else
                .Fields("HBSAG") = rsDataset("VALORE")
            End If
            .Fields("DATA_HBSAG") = CStr(rsDataset("DATA"))
        Else
            .Fields("HBSAG") = "- -"
            .Fields("DATA_HBSAG") = "- -"
        End If
        rsDataset.Close
                                                        
        rsDataset.Open strSql & "'HBSAB%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HBSAB") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HBSAB") = "POS."
            Else
                .Fields("HBSAB") = "- -"
            End If
            .Fields("DATA_HBSAB") = CStr(rsDataset("DATA"))
       Else
            .Fields("HBSAB") = "- -"
            .Fields("DATA_HBSAB") = "- -"
        End If
        rsDataset.Close
        
        rsDataset.Open strSql & "'%HBEAG%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HBEAG") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HBEAG") = "POS."
            Else
                .Fields("HBEAG") = rsDataset("VALORE")
            End If
            .Fields("DATA_HBEAG") = CStr(rsDataset("DATA"))
        Else
            .Fields("HBEAG") = "- -"
            .Fields("DATA_HBEAG") = "- -"
        End If
        rsDataset.Close
        
        rsDataset.Open strSql & "'%HBEAB%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HBEAB") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HBEAB") = "POS."
            Else
                .Fields("HBEAB") = rsDataset("VALORE")
            End If
            .Fields("DATA_HBEAB") = CStr(rsDataset("DATA"))
        Else
            .Fields("HBEAB") = "- -"
            .Fields("DATA_HBEAB") = "- -"
        End If
        rsDataset.Close
        
        rsDataset.Open strSql & "'%HBCAB%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HBCAB") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HBCAB") = "POS."
            Else
                .Fields("HBCAB") = rsDataset("VALORE")
            End If
            .Fields("DATA_HBCAB") = CStr(rsDataset("DATA"))
        Else
            .Fields("HBCAB") = "- -"
            .Fields("DATA_HBCAB") = "- -"
        End If
        rsDataset.Close
           
        rsDataset.Open strSql & "'%HCVAB%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HCVAB") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HCVAB") = "POS."
            Else
                .Fields("HCVAB") = rsDataset("VALORE")
            End If
            .Fields("DATA_HCV") = CStr(rsDataset("DATA"))
        Else
            .Fields("HCVAB") = "- -"
            .Fields("DATA_HCV") = "- -"
        End If
        rsDataset.Close

        rsDataset.Open strSql & "'HCV-RNA qual%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HCVRN") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HCVRN") = "POS."
            Else
                .Fields("HCVRN") = rsDataset("VALORE")
            End If
            .Fields("DATA_HCVRNA") = CStr(rsDataset("DATA"))
        Else
            .Fields("HCVRN") = "- -"
            .Fields("DATA_HCVRNA") = "- -"
        End If
        rsDataset.Close

        rsDataset.Open strSql & "'%HIV%' order by Data desc", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("VALORE") = -2 Then
                .Fields("HIV") = "NEG."
            ElseIf rsDataset("VALORE") = -1 Then
                .Fields("HIV") = "POS."
            Else
                .Fields("HIV") = rsDataset("VALORE")
            End If
            .Fields("IMMUNITA_HIV") = CStr(rsDataset("DATA"))
        Else
            .Fields("HIV") = "- -"
            .Fields("IMMUNITA_HIV") = "- -"
        End If
        rsDataset.Close

        
        ' Dati delle terapie dialitica e domiciliare
        strSql = "  SELECT      * " & _
                "   FROM        (TERAPIE_DIALITICHE " & _
                "               INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
                "   WHERE       TERAPIE_DIALITICHE.CODICE_PAZIENTE=" & intPazientiKey & " AND SOSPESA=FALSE " & _
                "   ORDER BY    TERAPIE_DIALITICHE.DATA DESC"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            Do While Not rsDataset.EOF
                strGiorni = ""
                If CBool(rsDataset("TUTTI_GIORNI")) Then
                    strGiorni = "Tutti i giorni"
                Else
                    For i = 1 To 7
                        If CBool(rsDataset("GIORNO" & i)) Then
                            strGiorni = strGiorni & UCase(Mid(WeekdayName(i), 1, 1)) & Mid(WeekdayName(i), 2, Len(WeekdayName(i))) & ", "
                        End If
                    Next
                    If strGiorni <> "" Then strGiorni = Mid(strGiorni, 1, Len(strGiorni) - 2)
                End If
                
                .Fields("FARMACO_TERAPIA_POST") = .Fields("FARMACO_TERAPIA_POST") & vbCrLf & rsDataset("NOME")
                .Fields("POS_TERAPIA_POST") = .Fields("POS_TERAPIA_POST") & vbCrLf & rsDataset("POSOLOGIA")
                .Fields("GIORNI_TERAPIA_POST") = .Fields("GIORNI_TERAPIA_POST") & vbCrLf & strGiorni
                .Fields("NOTE_TERAPIA_POST") = .Fields("NOTE_TERAPIA_POST") & vbCrLf & rsDataset("NOTE")
                                                   
                rsDataset.MoveNext
            Loop
        Else
            .Fields("FARMACO_TERAPIA_POST") = "- -"
        End If
        rsDataset.Close
        
        strSql = "  SELECT      * " & _
                "   FROM        (TERAPIE_DOMICILIARI " & _
                "               INNER JOIN MEDICINALI ON TERAPIE_DOMICILIARI.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
                "   WHERE       TERAPIE_DOMICILIARI.CODICE_PAZIENTE=" & intPazientiKey & " AND SOSPESA=FALSE " & _
                "   ORDER BY    TERAPIE_DOMICILIARI.DATA DESC"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            Do While Not rsDataset.EOF
                strGiorni = ""
                If CBool(rsDataset("TUTTI_GIORNI")) Then
                    strGiorni = "Tutti i giorni"
                Else
                    For i = 1 To 7
                        If CBool(rsDataset("GIORNO" & i)) Then
                            strGiorni = strGiorni & UCase(Mid(WeekdayName(i), 1, 1)) & Mid(WeekdayName(i), 2, Len(WeekdayName(i))) & ", "
                        End If
                    Next
                    If strGiorni <> "" Then strGiorni = Mid(strGiorni, 1, Len(strGiorni) - 2)
                End If
                
                .Fields("FARMACO_TERAPIA_DOMICILIARE") = .Fields("FARMACO_TERAPIA_DOMICILIARE") & vbCrLf & rsDataset("NOME")
                .Fields("POS_TERAPIA_DOMICILIARE") = .Fields("POS_TERAPIA_DOMICILIARE") & vbCrLf & rsDataset("POSOLOGIA")
                .Fields("GIORNI_TERAPIA_DOMICILIARE") = .Fields("GIORNI_TERAPIA_DOMICILIARE") & vbCrLf & strGiorni
                .Fields("NOTE_TERAPIA_DOMICILIARE") = .Fields("NOTE_TERAPIA_DOMICILIARE") & vbCrLf & rsDataset("SOMMINISTRAZIONE")
                                                                   
                rsDataset.MoveNext
            Loop
        Else
            .Fields("FARMACO_TERAPIA_DOMICILIARE") = "- -"
        End If
        rsDataset.Close
        
        ' Dati delle dialisi giornaliere
        strSql = "  SELECT      TOP 1 * " & _
                "   FROM        ((SCHEDE_DIALISI " & _
                "               INNER JOIN TURNI ON SCHEDE_DIALISI.CODICE_PAZIENTE=TURNI.CODICE_PAZIENTE) " & _
                "               INNER JOIN APPARATI ON TURNI.CODICE_RENE=APPARATI.KEY) " & _
                "   WHERE       SCHEDE_DIALISI.CODICE_PAZIENTE=" & intPazientiKey & " AND ERRATA=FALSE " & _
                "   ORDER BY    DATA DESC"
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            .Fields("ULTIMA_DIALISI_DEL") = rsDataset("DATA")
            .Fields("PESO_PRE_DIALISI") = Left(rsDataset("PESO_INIZIO"), 6)
            .Fields("PESO_POST_DIALISI") = Left(rsDataset("PESO_FINE"), 6)
            .Fields("PRESS_ARTER_PRE_DIAL") = rsDataset("PA_MIN1") & "/" & rsDataset("PA_MAX1")
            .Fields("PRESS_ARTER_POST_DIAL") = rsDataset("PA_MIN5") & "/" & rsDataset("PA_MAX5")
            .Fields("MONITOR_EMODIALISI") = rsDataset("MODELLO")
                       
            'intNumeroGiorni = 0
            strGiorni = ""
            For i = 1 To 7
                If rsDataset("AM_INIZIO" & i) <> "" Or rsDataset("PM_INIZIO" & i) <> "" Or rsDataset("SR_INIZIO" & i) <> "" Then
            '        intNumeroGiorni = intNumeroGiorni + 1
                    strGiorni = strGiorni & UCase(Mid(WeekdayName(i), 1, 1)) & Mid(WeekdayName(i), 2, Len(WeekdayName(i))) & " "
                End If
            Next
            '.Fields("RITMO_SETTIMANALE") = intNumeroGiorni
            .Fields("GIORNI_DIALISI") = strGiorni
        Else
            .Fields("PESO_PRE_DIALISI") = "- -"
            .Fields("PESO_POST_DIALISI") = "- -"
            .Fields("PRESS_ARTER_PRE_DIAL") = "- -"
            .Fields("PRESS_ARTER_POST_DIAL") = "- -"
            .Fields("DIALIZZATORE") = "- -"
            '.Fields("RITMO_SETTIMANALE") = "- -"
            .Fields("GIORNI_DIALISI") = "- -"
        End If
        rsDataset.Close

    End With
    Set rsDataset = Nothing


    Set rptCartellaDialiticaMaddaloni = Nothing
    Set rptCartellaDialiticaMaddaloni.DataSource = rsMain
    rptCartellaDialiticaMaddaloni.Sections("Intestazione").Controls.Item("lblCodiceID").Caption = intCodiceID
    rptCartellaDialiticaMaddaloni.PrintReport True, rptRangeAllPages

Else

    MsgBox "MODULO OPZIONALE A RICHIESTA", vbInformation, "INFORMAZIONE"

End If

End Sub

Private Sub cmdEliminaEpo_Click()
    cboEPO.ListIndex = -1
    txtUI = 0
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "ATTENZIONE!!!"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "La scheda deve essere prima memorizzata", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    End If
      
    Set rsAnamnesiDialitica = New Recordset
    rsAnamnesiDialitica.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsAnamnesiDialitica("COGNOME") & " " & rsAnamnesiDialitica("NOME")
    structIntestazione.sDataPaziente = rsAnamnesiDialitica("DATA_NASCITA")
    Set rsAnamnesiDialitica = Nothing

    Call StampaQuartaParte(False, intPazientiKey)
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim storico As Integer
    storico = 0
    
    If Completo Then
        If CompletoStorico Then
            Call Completa
            If cboAccesso.Text <> "" Then
                Call GestisciNuovo("ACCESSI_VASCOLARI", cboAccesso)
            End If
            If cboTipoDialisi.Text <> "" Then
                Call GestisciNuovo("TIPI_DIALISI", cboTipoDialisi)
            End If
            If cboCartuccia.Text <> "" Then
                Call GestisciNuovo("CARTUCCE", cboCartuccia)
            End If
            If cboSolDialitica.Text <> "" Then
                Call GestisciNuovo("SOL_DIALITICHE", cboSolDialitica)
            End If
            If cboSolInfusionale.Text <> "" Then
                Call GestisciNuovo("SOL_INFUSIONALI", cboSolInfusionale)
            End If
            ' carica i vettori
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DIURESI", "QUANTITA", "PESO_SECCO", "AUMENTO_POND", _
                     "SEDE_ACCESSO", "SODIO", "POTASSIO", "BICARBONATO", "CALCIO", "DOSE1", _
                     "DOSE2", "DOSE3", "DOSE4", "ORE", "MINUTI", "DATA_PESO", _
                     "DATA_FILTRO", "DATA_LINEE", "FLUSSO", "FLUSSO_SANGUE", "SOL_INF_CC", "EPO", _
                     "UI", "NOTE", "CARTUCCIA", "TIPO_FILTRO", "ACCESSO_VASCOLARE", "AGO1", _
                     "AGO2", "TIPO_DIALISI", "ANTICOAGULANTE1", "ANTICOAGULANTE2", "TIPO_LINEE", "SOL_DIALITICA", _
                     "SOL_INFUSIONALE", "DOSI_UNITA_MISURA", "UNITA_VAL_SOL_INF", "GLUCOSIO", "CODICE_PRESTAZIONE", "RITMO_DIALITICO")
            
            If Not modifica Then
                keyId = GetNumero("ANAMNESI_DIALITICHE")
            End If
            
            v_Val = Array(keyId, intPazientiKey, IIf(chkDiuresiResidua.Value = Checked, True, False), _
                    txtQuantita, txtPesoSecco, txtAumentoPond, txtSedeAccesso, txtSodio.GetDecimal, txtPotassio, _
                    txtBicarbonato, txtCalcio, txtDose(0), txtDose(1), txtDose(2), txtDose(3), txtOre, _
                    txtMinuti, oData(0).data, oData(1).data, oData(2).data, txtFlusso, txtFlussoSangue, _
                    txtSolInfCc, cboEPO.ListIndex, txtUI, IIf(txtNote = "", " ", txtNote), _
                    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, cboDosiUnitaMisura.ListIndex, cboUnit‡ValoreInfusionale.ListIndex, txtGlucosio, cboCodicePrestaione.ItemData(cboCodicePrestaione.ListIndex), txtRitmoDialitico)
                    
            If cboCartuccia.ListIndex <> -1 Then v_Val(26) = cboCartuccia.ItemData(cboCartuccia.ListIndex)
            If cboTipoFiltro.ListIndex <> -1 Then v_Val(27) = cboTipoFiltro.ItemData(cboTipoFiltro.ListIndex)
            If cboAccesso.ListIndex <> -1 Then v_Val(28) = cboAccesso.ItemData(cboAccesso.ListIndex)
            If cboTipoAgo(0).ListIndex <> -1 Then v_Val(29) = cboTipoAgo(0).ItemData(cboTipoAgo(0).ListIndex)
            If cboTipoAgo(1).ListIndex <> -1 Then v_Val(30) = cboTipoAgo(1).ItemData(cboTipoAgo(1).ListIndex)
            If cboTipoDialisi.ListIndex <> -1 Then v_Val(31) = cboTipoDialisi.ItemData(cboTipoDialisi.ListIndex)
            If cboAnticoagulante(0).ListIndex <> -1 Then v_Val(32) = cboAnticoagulante(0).ItemData(cboAnticoagulante(0).ListIndex)
            If cboAnticoagulante(1).ListIndex <> -1 Then v_Val(33) = cboAnticoagulante(1).ItemData(cboAnticoagulante(1).ListIndex)
            If cboTipoLinee.ListIndex <> -1 Then v_Val(34) = cboTipoLinee.ItemData(cboTipoLinee.ListIndex)
            If cboSolDialitica.ListIndex <> -1 Then v_Val(35) = cboSolDialitica.ItemData(cboSolDialitica.ListIndex)
            If cboSolInfusionale.ListIndex <> -1 Then v_Val(36) = cboSolInfusionale.ItemData(cboSolInfusionale.ListIndex)
                                    
            cnPrinc.BeginTrans
            Set rsAnamnesiDialitica = New Recordset
            If modifica Then
                rsAnamnesiDialitica.Open "SELECT * FROM ANAMNESI_DIALITICHE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                rsAnamnesiDialitica.Update v_Nomi, v_Val
                If TRACCIATO Then
                    Call Confronta
                End If
                storico = GestisciStorico(keyId)
                If storico = -1 Then Exit Sub
            Else
                rsAnamnesiDialitica.Open "ANAMNESI_DIALITICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                rsAnamnesiDialitica.AddNew v_Nomi, v_Val
                rsAnamnesiDialitica.Update
                Call Upd_rsDisco
            End If
            Set rsAnamnesiDialitica = Nothing
            cnPrinc.CommitTrans
            
            MsgBox "Salvataggio effettuato" & vbCrLf & IIf(storico, "(CON STORICIZZAZIONE)", "(SENZA STORICIZZAZIONE)"), vbInformation, "Salvataggio"
            blnModificato = False
            modifica = True
        End If
    End If
End Sub

Private Sub cboTipoDialisi_KeyPress(KeyAscii As Integer)
    If Len(cboTipoDialisi.Text) >= 25 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub cboAccesso_KeyPress(KeyAscii As Integer)
    If Len(cboAccesso.Text) >= 25 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub cboCartuccia_KeyPress(KeyAscii As Integer)
    If Len(cboCartuccia.Text) > 39 And Not KeyAscii = vbKeyBack Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub cboSolDialitica_KeyPress(KeyAscii As Integer)
    If Len(cboSolDialitica.Text) > 39 And Not KeyAscii = vbKeyBack Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub cboEPO_Click()
    If cboEPO.ListIndex = 2 Or cboEPO.ListIndex = 3 Then
        lblUnitaMisura = "mcg"
    Else
        lblUnitaMisura = "UI"
    End If
    If cboEPO.ListIndex = epoAppo Then
        txtUI = uiAppo
    Else
        txtUI = ""
    End If
    blnModificato = True
End Sub

Private Sub cboSolInfusionale_KeyPress(KeyAscii As Integer)
    If Len(cboSolInfusionale.Text) > 39 And Not KeyAscii = vbKeyBack Then
        Beep
        KeyAscii = 0
    End If
End Sub

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
    Set rsAnamnesiDialitica = New Recordset
    rsAnamnesiDialitica.Open "SELECT * FROM ANAMNESI_DIALITICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsAnamnesiDialitica.BOF And rsAnamnesiDialitica.EOF Then
        ' il paziente non ha una scheda clinica
        modifica = False
        ' inserimento di un nuovo paziente di default inserisce 39.95.4. per evitare il crash
        cboCodicePrestaione.ListIndex = GetIndex(cboCodicePrestaione, "39.95.4")
    Else
        keyId = rsAnamnesiDialitica("KEY")
        modifica = True
        ' carica i dati della scheda dialitica
        If IsNull(rsAnamnesiDialitica("RITMO_DIALITICO")) Then
            txtRitmoDialitico = 0
        Else
            txtRitmoDialitico = rsAnamnesiDialitica("RITMO_DIALITICO")
        End If
        chkDiuresiResidua.Value = IIf(CBool(rsAnamnesiDialitica("DIURESI")) = True, Checked, Unchecked)
        txtAumentoPond = VirgolaOrPunto(rsAnamnesiDialitica("AUMENTO_POND"), ",")
        txtQuantita = VirgolaOrPunto(rsAnamnesiDialitica("QUANTITA"), ",")
        txtPesoSecco = VirgolaOrPunto(rsAnamnesiDialitica("PESO_SECCO"), ",")
        txtFlusso = VirgolaOrPunto(rsAnamnesiDialitica("FLUSSO"), ",")
        txtFlussoSangue = VirgolaOrPunto(rsAnamnesiDialitica("FLUSSO_SANGUE"), ",")
        txtSolInfCc = VirgolaOrPunto(rsAnamnesiDialitica("SOL_INF_CC"), ",")
        cboTipoFiltro.ListIndex = GetCboListIndex(rsAnamnesiDialitica("TIPO_FILTRO"), cboTipoFiltro)
        cboTipoLinee.ListIndex = GetCboListIndex(rsAnamnesiDialitica("TIPO_LINEE"), cboTipoLinee)
        cboAccesso.ListIndex = GetCboListIndex(rsAnamnesiDialitica("ACCESSO_VASCOLARE"), cboAccesso)
        cboTipoAgo(0).ListIndex = GetCboListIndex(rsAnamnesiDialitica("AGO1"), cboTipoAgo(0))
        cboTipoAgo(1).ListIndex = GetCboListIndex(rsAnamnesiDialitica("AGO2"), cboTipoAgo(1))
        cboTipoDialisi.ListIndex = GetCboListIndex(rsAnamnesiDialitica("TIPO_DIALISI"), cboTipoDialisi)
        If IsNull(rsAnamnesiDialitica("CODICE_PRESTAZIONE")) Then
            ' se non c'Ë il codice prestazione di default inserisce 39.95.4.
            cboCodicePrestaione.ListIndex = GetIndex(cboCodicePrestaione, "39.95.4")
        Else
            cboCodicePrestaione.ListIndex = GetCboListIndex(rsAnamnesiDialitica("CODICE_PRESTAZIONE"), cboCodicePrestaione)
        End If
        cboCartuccia.ListIndex = GetCboListIndex(rsAnamnesiDialitica("CARTUCCIA"), cboCartuccia)
        cboSolDialitica.ListIndex = GetCboListIndex(rsAnamnesiDialitica("SOL_DIALITICA"), cboSolDialitica)
        cboSolInfusionale.ListIndex = GetCboListIndex(rsAnamnesiDialitica("SOL_INFUSIONALE"), cboSolInfusionale)
        cboUnit‡ValoreInfusionale.ListIndex = rsAnamnesiDialitica("UNITA_VAL_SOL_INF")
        If cboUnit‡ValoreInfusionale.ListIndex = -1 Then cboUnit‡ValoreInfusionale.ListIndex = 0
        For i = 0 To 1
            cboAnticoagulante(i).ListIndex = GetCboListIndex(rsAnamnesiDialitica("ANTICOAGULANTE" & i + 1), cboAnticoagulante(i))
        Next i
        cboEPO.ListIndex = rsAnamnesiDialitica("EPO")
        epoAppo = cboEPO.ListIndex
        txtUI = rsAnamnesiDialitica("UI")
        uiAppo = txtUI
        txtNote = rsAnamnesiDialitica("NOTE")
        txtSedeAccesso = rsAnamnesiDialitica("SEDE_ACCESSO")
        txtSodio.Text = rsAnamnesiDialitica("SODIO")
        txtPotassio = VirgolaOrPunto(rsAnamnesiDialitica("POTASSIO"), ",")
        txtBicarbonato = VirgolaOrPunto(rsAnamnesiDialitica("BICARBONATO"), ",")
        txtCalcio = VirgolaOrPunto(rsAnamnesiDialitica("CALCIO"), ",")
        txtGlucosio = VirgolaOrPunto(rsAnamnesiDialitica("GLUCOSIO"), ",")
        txtDose(0) = VirgolaOrPunto(rsAnamnesiDialitica("DOSE1"), ",")
        txtDose(1) = VirgolaOrPunto(rsAnamnesiDialitica("DOSE2"), ",")
        txtDose(2) = VirgolaOrPunto(rsAnamnesiDialitica("DOSE3"), ",")
        txtDose(3) = VirgolaOrPunto(rsAnamnesiDialitica("DOSE4"), ",")
        cboDosiUnitaMisura.ListIndex = rsAnamnesiDialitica("DOSI_UNITA_MISURA")
        If cboDosiUnitaMisura.ListIndex = 0 Then Call PredisponiDosi
        If cboDosiUnitaMisura.ListIndex = -1 Then cboDosiUnitaMisura.ListIndex = 0
        txtOre = rsAnamnesiDialitica("ORE")
        txtMinuti = rsAnamnesiDialitica("MINUTI")
        oData(0).data = rsAnamnesiDialitica("DATA_PESO")
        oData(1).data = rsAnamnesiDialitica("DATA_FILTRO")
        oData(2).data = rsAnamnesiDialitica("DATA_LINEE")
        dataFiltro = oData(1).data
        dataPeso = oData(0).data
        dataLinee = oData(2).data
        vecchioFiltro = GetCboListIndex(rsAnamnesiDialitica("TIPO_FILTRO"), cboTipoFiltro)
        vecchioPeso = VirgolaOrPunto(txtPesoSecco, ".")
        vecchioLinee = GetCboListIndex(rsAnamnesiDialitica("TIPO_LINEE"), cboTipoLinee)
        
        Call Upd_rsDisco
    End If
    Set rsAnamnesiDialitica = Nothing
    blnModificato = False
End Sub

Private Sub lblPesoIniziale_Click()

End Sub

Private Sub oData_OnDataChange(Index As Integer)
    blnModificato = True
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

Private Sub picCaricaPesoSecco_Click()
    If MsgBox("Sicuro di voler caricare il peso secco dall'ultima seduta dialitica ?", vbQuestion + vbYesNo, "Carica peso secco") = vbYes Then
        Dim rsDataset As New Recordset
        rsDataset.Open "Select Top 1 PESO_FINE, DATA From SCHEDE_DIALISI Where CODICE_PAZIENTE=" & intPazientiKey & " AND ERRATA=FALSE ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            txtPesoSecco = VirgolaOrPunto(rsDataset("PESO_FINE"), ",")
            oData(0).data = rsDataset("DATA")
        Else
            MsgBox "Il paziente " & lblCognome & " " & lblNome & " non ha sedute dialitiche memorizzare in archivio", vbInformation, "Carica Peso Secco"
        End If
        rsDataset.Close
        Set rsDataset = Nothing
    End If
End Sub

Private Sub picCaricaPesoSecco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCaricaPesoSecco.Picture = LoadResPicture("PESO2", 0)
End Sub

Private Sub picCaricaPesoSecco_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCaricaPesoSecco.Picture = LoadResPicture("PESO1", 0)
End Sub

Private Sub picElenca_Click(Index As Integer)
    ' apre il form dello storico
    Select Case Index
        Case 0
            tStorico.Tipo = tpsPESO
        Case 1
            tStorico.Tipo = tpsFILTRO
        Case 2
            tStorico.Tipo = tpsLINEE
    End Select
    tStorico.condizione = "WHERE CODICE_SCHEDA=" & keyId
    frmStorico.Show 1
End Sub

Private Sub picElenca_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picElenca(Index).Picture = LoadResPicture("storico2", 0)
End Sub

Private Sub picElenca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picElenca(Index).Picture = LoadResPicture("storico1", 0)
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
            'Call RitmoDialisiSettimanale
        End If
    End If
End Sub

'Private Sub RitmoDialisiSettimanale()
'    Dim rsDataset As Recordset
'    Dim strSql As String
'    Dim i As Integer
'    Dim NumeroDialisiSettimanale As Integer
    
'    Set rsDataset = New Recordset
    
'    strSql = "  SELECT      TOP 1 * " & _
'             "   FROM        ((SCHEDE_DIALISI " & _
'             "               INNER JOIN TURNI ON SCHEDE_DIALISI.CODICE_PAZIENTE=TURNI.CODICE_PAZIENTE) " & _
'             "               INNER JOIN APPARATI ON TURNI.CODICE_RENE=APPARATI.KEY) " & _
'             "   WHERE       SCHEDE_DIALISI.CODICE_PAZIENTE=" & intPazientiKey & " AND ERRATA=FALSE " & _
'             "   ORDER BY    DATA DESC"
'    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
'        If Not (rsDataset.EOF And rsDataset.BOF) Then
'            For i = 1 To 7
'                If rsDataset("AM_INIZIO" & i) <> "" Or rsDataset("PM_INIZIO" & i) <> "" Or rsDataset("SR_INIZIO" & i) <> "" Then
'                    NumeroDialisiSettimanale = NumeroDialisiSettimanale + 1
'                End If
'            Next
'            lblRitmoDialisiSettimanale.Caption = NumeroDialisiSettimanale & " sedute"
'        Else
'            lblRitmoDialisiSettimanale.Caption = "- -"
'        End If
        
'    rsDataset.Close
'    Set rsDataset = Nothing
    
    ' Azzero la variabile per evitare di caricare lo stesso valore
'    NumeroDialisiSettimanale = 0
'End Sub

Private Sub tabSchede_Click(PreviousTab As Integer)
    cboTipoDialisi.SelLength = 0
    cboAccesso.SelLength = 0
    cboSolDialitica.SelLength = 0
    cboSolInfusionale.SelLength = 0
    cboCartuccia.SelLength = 0
End Sub

Private Sub txtAumentoPond_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtAumentoPond, lettera)
    blnModificato = True
End Sub

Private Sub txtAumentoPond_GotFocus()
    txtAumentoPond.BackColor = colArancione
End Sub

Private Sub txtAumentoPond_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtAumentoPond_LostFocus()
    txtAumentoPond.BackColor = vbWhite
End Sub

Private Sub txtAumentoPond_Validate(Cancel As Boolean)
    If txtAumentoPond = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtAumentoPond.Text)
    End If
End Sub

Private Sub txtDose_Change(Index As Integer)
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtDose(Index), lettera)
    blnModificato = True
End Sub

Private Sub txtDose_GotFocus(Index As Integer)
    txtDose(Index).BackColor = colArancione
End Sub

Private Sub txtDose_KeyPress(Index As Integer, KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtDose_LostFocus(Index As Integer)
    txtDose(Index).BackColor = vbWhite
End Sub

Private Sub txtDose_Validate(Index As Integer, Cancel As Boolean)
    If txtDose(Index).Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtDose(Index).Text)
    End If
End Sub

Private Sub txtFlusso_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtFlusso, lettera)
    blnModificato = True
End Sub

Private Sub txtFlusso_GotFocus()
    txtFlusso.BackColor = colArancione
End Sub

Private Sub txtFlusso_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtFlusso_LostFocus()
    txtFlusso.BackColor = vbWhite
End Sub

Private Sub txtFlusso_Validate(Cancel As Boolean)
    If txtFlusso.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtFlusso.Text)
    End If
End Sub

Private Sub txtFlussoSangue_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtFlussoSangue, lettera)
    blnModificato = True
End Sub

Private Sub txtFlussoSangue_GotFocus()
    txtFlussoSangue.BackColor = colArancione
End Sub

Private Sub txtFlussoSangue_LostFocus()
    txtFlussoSangue.BackColor = vbWhite
End Sub

Private Sub txtGlucosio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtGlucosio, lettera)
    blnModificato = True
End Sub

Private Sub txtGlucosio_GotFocus()
    txtGlucosio.BackColor = colArancione
End Sub

Private Sub txtGlucosio_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtGlucosio_LostFocus()
    txtGlucosio.BackColor = vbWhite
End Sub

Private Sub txtGlucosio_Validate(Cancel As Boolean)
    If txtGlucosio = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtGlucosio.Text)
    End If
End Sub

Private Sub txtMinuti_GotFocus()
    txtMinuti.BackColor = colArancione
End Sub

Private Sub txtMinuti_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMinuti_LostFocus()
    txtMinuti.BackColor = vbWhite
End Sub

Private Sub txtNote_GotFocus()
    txtNote.BackColor = colArancione
End Sub

Private Sub txtNote_LostFocus()
    txtNote.BackColor = vbWhite
End Sub

Private Sub txtOre_GotFocus()
    txtOre.BackColor = colArancione
End Sub

Private Sub txtOre_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtOre_LostFocus()
    txtOre.BackColor = vbWhite
End Sub

Private Sub txtPesoSecco_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPesoSecco, lettera)
    blnModificato = True
End Sub

Private Sub txtPesoSecco_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPesoSecco_Validate(Cancel As Boolean)
    If txtPesoSecco = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPesoSecco.Text)
    End If
End Sub

Private Sub txtBicarbonato_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtBicarbonato, lettera)
    blnModificato = True
End Sub

Private Sub txtPotassio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPotassio, lettera)
    blnModificato = True
End Sub

Private Sub txtBicarbonato_GotFocus()
    txtBicarbonato.BackColor = colArancione
End Sub

Private Sub txtPotassio_GotFocus()
    txtPotassio.BackColor = colArancione
End Sub

Private Sub txtBicarbonato_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPotassio_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtBicarbonato_LostFocus()
    txtBicarbonato.BackColor = vbWhite
End Sub

Private Sub txtPotassio_LostFocus()
    txtPotassio.BackColor = vbWhite
End Sub

Private Sub txtBicarbonato_Validate(Cancel As Boolean)
    If txtBicarbonato = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtBicarbonato.Text)
    End If
End Sub

Private Sub txtPotassio_Validate(Cancel As Boolean)
    If txtPotassio = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPotassio.Text)
    End If
End Sub

Private Sub txtQuantita_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtQuantita, lettera)
    blnModificato = True
End Sub

Private Sub txtQuantita_GotFocus()
    txtQuantita.BackColor = colArancione
End Sub

Private Sub txtQuantita_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtQuantita_LostFocus()
    txtQuantita.BackColor = vbWhite
End Sub

Private Sub txtQuantita_Validate(Cancel As Boolean)
    If txtQuantita = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtQuantita.Text)
    End If
End Sub

Private Sub txtRitmoDialitico_Change()
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtRitmoDialitico, lettera)
End Sub

Private Sub txtRitmoDialitico_GotFocus()
    txtRitmoDialitico.BackColor = colArancione
End Sub

Private Sub txtRitmoDialitico_KeyPress(KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtRitmoDialitico_LostFocus()
    txtRitmoDialitico.BackColor = vbWhite
End Sub

Private Sub txtSedeAccesso_GotFocus()
    txtSedeAccesso.BackColor = colArancione
End Sub

Private Sub txtSedeAccesso_LostFocus()
    txtSedeAccesso.BackColor = vbWhite
End Sub

Private Sub txtCalcio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtCalcio, lettera)
    blnModificato = True
End Sub

Private Sub txtCalcio_GotFocus()
    txtCalcio.BackColor = colArancione
End Sub

Private Sub txtCalcio_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtCalcio_LostFocus()
    txtCalcio.BackColor = vbWhite
End Sub

Private Sub txtCalcio_Validate(Cancel As Boolean)
    If txtCalcio = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtCalcio.Text)
    End If
End Sub

Private Sub txtSodio_OnChange()
    blnModificato = True
End Sub

Private Sub txtSolInfCc_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtSodio, lettera)
    blnModificato = True
End Sub

Private Sub txtSolInfCc_GotFocus()
    txtSolInfCc.BackColor = colArancione
End Sub

Private Sub txtSolInfCc_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtSolInfCc_LostFocus()
    txtSolInfCc.BackColor = vbWhite
End Sub

Private Sub txtSolInfCc_Validate(Cancel As Boolean)
    If txtSolInfCc = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtSolInfCc.Text)
    End If
End Sub

Private Sub txtUI_GotFocus()
    txtUI.BackColor = colArancione
End Sub

Private Sub txtUI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtUI_LostFocus()
    txtUI.BackColor = vbWhite
End Sub

'******** Gestione Modificato

Private Sub txtSedeAccesso_Change()
    blnModificato = True
End Sub

Private Sub txtMinuti_Change()
    blnModificato = True
End Sub

Private Sub txtOre_Change()
    blnModificato = True
End Sub

Private Sub txtNote_Change()
    blnModificato = True
End Sub

Private Sub txtUI_Change()
    blnModificato = True
End Sub

Private Sub cboTipoFiltro_Click()
    blnModificato = True
End Sub

Private Sub cboTipoLinee_Click()
    blnModificato = True
End Sub

Private Sub cboAccesso_Click()
    blnModificato = True
End Sub

Private Sub cboAnticoagulante_Click(Index As Integer)
    blnModificato = True
End Sub

Private Sub cboCartuccia_Change()
    blnModificato = True
End Sub

Private Sub cboCartuccia_Click()
    blnModificato = True
End Sub

Private Sub cboDosiUnitaMisura_Click()
    blnModificato = True
End Sub

Private Sub cboSolDialitica_Change()
    blnModificato = True
End Sub

Private Sub cboSolDialitica_Click()
    blnModificato = True
End Sub

Private Sub cboSolInfusionale_Change()
    blnModificato = True
End Sub

Private Sub cboSolInfusionale_Click()
    blnModificato = True
End Sub

Private Sub cboTipoAgo_Click(Index As Integer)
    blnModificato = True
End Sub

Private Sub cboTipoDialisi_Change()
    blnModificato = True
End Sub

Private Sub cboTipoDialisi_Click()
    blnModificato = True
End Sub

Private Sub chkDiuresiResidua_Click()
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
        For i = 0 To rsAnamnesiDialitica.Fields.count - 1
            rsDisco.Fields(i) = rsAnamnesiDialitica.Fields(i)
        Next i
        rsDisco.Update
End Sub
