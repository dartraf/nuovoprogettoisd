VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSchedaDialitica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seduta Dialitica Giornaliera - Compilazione"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSchede 
      Height          =   4305
      Left            =   120
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7594
      _Version        =   393216
      TabHeight       =   520
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
      TabCaption(0)   =   "Scheda dialitica 1"
      TabPicture(0)   =   "frmSchedaDialitica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(27)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(25)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(30)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(31)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(33)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(34)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(26)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(38)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(35)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(28)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(29)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(42)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(43)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(23)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(32)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(52)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(51)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(50)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(49)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblTipoLinee"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblAgo1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblAgo2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblFiltro"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblTipoDialisi"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblAccessoVascolare"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblAnticoagulante(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblAnticoagulante(1)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPesoSecco"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblUltimoPeso"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblDataUltimoPeso"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblOreDialisi"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblDoseIniziale"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblDoseAltroAnticoagulante"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblSodio"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblPotassio"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lblBicarbonato"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lblCalcio"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblGlucosio"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(48)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lblDoseIntermedia"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label1(54)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lblDoseFinale"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label1(55)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lblDoseUnitaMisura"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Scheda dialitica 2"
      TabPicture(1)   =   "frmSchedaDialitica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(1)=   "Label1(7)"
      Tab(1).Control(2)=   "Label1(18)"
      Tab(1).Control(3)=   "Label1(19)"
      Tab(1).Control(4)=   "Label1(20)"
      Tab(1).Control(5)=   "Label1(21)"
      Tab(1).Control(6)=   "lblFlusso"
      Tab(1).Control(7)=   "lblFlussoSangue"
      Tab(1).Control(8)=   "lblSolDialitica"
      Tab(1).Control(9)=   "lblSolInfusionale"
      Tab(1).Control(10)=   "lblCartuccia"
      Tab(1).Control(11)=   "lblSolInfCc"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Terapia"
      TabPicture(2)   =   "frmSchedaDialitica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxGriglia(0)"
      Tab(2).Control(1)=   "flxGriglia(1)"
      Tab(2).Control(2)=   "Label1(37)"
      Tab(2).Control(3)=   "Label1(36)"
      Tab(2).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Index           =   0
         Left            =   -74880
         TabIndex        =   50
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaDialitica.frx":0054
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Index           =   1
         Left            =   -69000
         TabIndex        =   51
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaDialitica.frx":00EC
      End
      Begin VB.Label lblDoseUnitaMisura 
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
         Left            =   6360
         TabIndex        =   137
         Top             =   2880
         Width           =   375
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
         Index           =   55
         Left            =   8880
         TabIndex        =   136
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblDoseFinale 
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
         Left            =   10320
         TabIndex        =   135
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose interm."
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
         Index           =   54
         Left            =   6840
         TabIndex        =   134
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Label lblDoseIntermedia 
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
         TabIndex        =   133
         Top             =   2880
         Width           =   615
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
         Index           =   48
         Left            =   8760
         TabIndex        =   132
         Top             =   3855
         Width           =   480
      End
      Begin VB.Label lblGlucosio 
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
         Left            =   9480
         TabIndex        =   131
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblSolInfCc 
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
         Left            =   -64560
         TabIndex        =   117
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblCartuccia 
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
         Left            =   -72000
         TabIndex        =   116
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label lblSolInfusionale 
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
         Left            =   -72000
         TabIndex        =   115
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label lblSolDialitica 
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
         Left            =   -72000
         TabIndex        =   114
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label lblFlussoSangue 
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
         Left            =   -67680
         TabIndex        =   113
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblFlusso 
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
         Left            =   -72000
         TabIndex        =   112
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblCalcio 
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
         Left            =   7800
         TabIndex        =   111
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblBicarbonato 
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
         Left            =   6360
         TabIndex        =   110
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblPotassio 
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
         Left            =   4440
         TabIndex        =   109
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblSodio 
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
         Left            =   3000
         TabIndex        =   108
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblDoseAltroAnticoagulante 
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
         Left            =   10320
         TabIndex        =   107
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblDoseIniziale 
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
         Left            =   5760
         TabIndex        =   106
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblOreDialisi 
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
         Left            =   10320
         TabIndex        =   105
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblDataUltimoPeso 
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
         TabIndex        =   104
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblUltimoPeso 
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
         Left            =   5160
         TabIndex        =   103
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblPesoSecco 
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
         Left            =   2520
         TabIndex        =   102
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblAnticoagulante 
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
         Left            =   2520
         TabIndex        =   101
         Top             =   3360
         Width           =   5535
      End
      Begin VB.Label lblAnticoagulante 
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
         Left            =   2520
         TabIndex        =   100
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblAccessoVascolare 
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
         Left            =   2520
         TabIndex        =   99
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label lblTipoDialisi 
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
         Left            =   2520
         TabIndex        =   98
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label lblFiltro 
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
         Left            =   2520
         TabIndex        =   97
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblAgo2 
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
         TabIndex        =   96
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label lblAgo1 
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
         TabIndex        =   95
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblTipoLinee 
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
         Left            =   2520
         TabIndex        =   94
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo di Linee"
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
         Index           =   49
         Left            =   240
         TabIndex        =   82
         Top             =   480
         Width           =   1380
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
         Index           =   50
         Left            =   240
         TabIndex        =   81
         Top             =   1920
         Width           =   2040
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
         Index           =   51
         Left            =   7275
         TabIndex        =   80
         Top             =   480
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
         Index           =   52
         Left            =   7275
         TabIndex        =   79
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bagno dialisi"
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
         Left            =   240
         TabIndex        =   70
         Top             =   3840
         Width           =   1380
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
         Index           =   23
         Left            =   7320
         TabIndex        =   69
         Top             =   3840
         Width           =   420
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
         Index           =   43
         Left            =   5520
         TabIndex        =   68
         Top             =   3840
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "del"
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
         Index           =   42
         Left            =   6240
         TabIndex        =   60
         Top             =   2400
         Width           =   345
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
         Index           =   21
         Left            =   -74760
         TabIndex        =   59
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "valore (cc)"
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
         Left            =   -65880
         TabIndex        =   58
         Top             =   1485
         Width           =   1125
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
         Index           =   19
         Left            =   -74760
         TabIndex        =   57
         Top             =   1440
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   56
         Top             =   960
         Width           =   1950
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
         Index           =   7
         Left            =   -70560
         TabIndex        =   55
         Top             =   480
         Width           =   2805
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
         Index           =   6
         Left            =   -74760
         TabIndex        =   54
         Top             =   480
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Terapia Postdialitica"
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
         Left            =   -66960
         TabIndex        =   53
         Top             =   480
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Terapia Intradialitica"
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
         Left            =   -73080
         TabIndex        =   52
         Top             =   480
         Width           =   2175
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
         Index           =   29
         Left            =   240
         TabIndex        =   41
         Top             =   2880
         Width           =   1560
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
         Index           =   28
         Left            =   240
         TabIndex        =   40
         Top             =   3360
         Width           =   2100
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
         Index           =   35
         Left            =   4320
         TabIndex        =   39
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dosi"
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
         Left            =   9600
         TabIndex        =   38
         Top             =   3360
         Width           =   495
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
         Index           =   26
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ore di Dialisi"
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
         Index           =   34
         Left            =   8760
         TabIndex        =   36
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo Peso"
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
         Left            =   3720
         TabIndex        =   35
         Top             =   2400
         Width           =   1275
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
         Index           =   31
         Left            =   2520
         TabIndex        =   34
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
         Index           =   30
         Left            =   4080
         TabIndex        =   33
         Top             =   3840
         Width           =   270
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
         Index           =   25
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peso Secco"
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
         Left            =   240
         TabIndex        =   31
         Top             =   2400
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   71
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmSchedaDialitica.frx":0184
         Style           =   1  'Graphical
         TabIndex        =   75
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         Left            =   10560
         TabIndex        =   72
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   120
      TabIndex        =   42
      Top             =   720
      Width           =   11895
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   405
         Index           =   1
         Left            =   380
         Picture         =   "frmSchedaDialitica.frx":05DD
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   740
         Width           =   405
      End
      Begin VB.CommandButton cmdCercaOra 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaOra 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
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
         Left            =   10800
         TabIndex        =   130
         Top             =   855
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo: "
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
         Left            =   10200
         TabIndex        =   129
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblNumeroRene 
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   4200
         TabIndex        =   128
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° rene"
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
         Index           =   47
         Left            =   3360
         TabIndex        =   127
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postazione"
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
         Index           =   46
         Left            =   1080
         TabIndex        =   126
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monitor"
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
         Left            =   5640
         TabIndex        =   125
         Top             =   840
         Width           =   780
      End
      Begin VB.Label lblPostazione 
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   2280
         TabIndex        =   124
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTipoRene 
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   6480
         TabIndex        =   123
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
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
         Left            =   7560
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ora Fine"
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
         Left            =   5160
         TabIndex        =   48
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ora Inizio"
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
         Left            =   2760
         TabIndex        =   47
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblOra 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblOra 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTurno 
         AutoSize        =   -1  'True
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
         Left            =   8280
         TabIndex        =   44
         Top             =   380
         Width           =   75
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
         Index           =   2
         Left            =   360
         TabIndex        =   43
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraScheda 
      Height          =   4455
      Left            =   120
      TabIndex        =   61
      Top             =   1920
      Width           =   11895
      Begin VB.TextBox txtPressioneMax 
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
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMax 
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
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMax 
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
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMax 
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
         Index           =   4
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMax 
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
         Left            =   8280
         MaxLength       =   3
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMin 
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
         Index           =   4
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMin 
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
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMin 
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
         Left            =   8280
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMin 
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
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtPressioneMin 
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
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtFC 
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
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtFC 
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
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtFC 
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
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtFC 
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
         Left            =   8280
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtFC 
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
         Index           =   4
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtPesoIniziale 
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
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtIncremento 
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
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtPesoFinale 
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
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   4
         Top             =   840
         Width           =   605
      End
      Begin VB.CheckBox chkErrata 
         Caption         =   "Scheda Annullata"
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
         Left            =   9360
         TabIndex        =   23
         Top             =   3390
         Width           =   2325
      End
      Begin VB.TextBox txtConplicanze 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1920
         Width           =   9855
      End
      Begin VB.CheckBox chkConferma 
         Caption         =   "Conferma Avvenuta Somministrazione"
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
         Left            =   4920
         TabIndex        =   22
         Top             =   3390
         Width           =   4335
      End
      Begin VB.Label lblUI 
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
         Left            =   3720
         TabIndex        =   121
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblEpo 
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
         TabIndex        =   120
         Top             =   3360
         Width           =   1095
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
         Left            =   8160
         TabIndex        =   119
         Top             =   3960
         Width           =   3375
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
         Left            =   3120
         TabIndex        =   118
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P.A. Max"
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
         Index           =   53
         Left            =   5280
         TabIndex        =   93
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Finale"
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
         Left            =   9960
         TabIndex        =   92
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3ª ora"
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
         Left            =   9120
         TabIndex        =   91
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1ª ora"
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
         Left            =   7440
         TabIndex        =   90
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P.A. Min"
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
         Left            =   5280
         TabIndex        =   89
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F.C."
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
         Index           =   44
         Left            =   5760
         TabIndex        =   88
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Iniziale"
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
         Left            =   6600
         TabIndex        =   87
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2ª ora"
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
         Index           =   45
         Left            =   8280
         TabIndex        =   86
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peso Finale"
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
         Left            =   240
         TabIndex        =   85
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peso Iniziale"
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
         TabIndex        =   84
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Incremento Pond. (Kg)"
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
         Left            =   240
         TabIndex        =   83
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Diario infermieristico"
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
         Index           =   24
         Left            =   240
         TabIndex        =   67
         Top             =   1920
         Width           =   1470
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Scheda compilata da:"
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
         TabIndex        =   66
         Top             =   3720
         Width           =   1575
         WordWrap        =   -1  'True
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
         Left            =   1920
         TabIndex        =   65
         Top             =   3960
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
         Index           =   41
         Left            =   7200
         TabIndex        =   64
         Top             =   3960
         Width           =   630
      End
      Begin VB.Label lblUnitaMisura 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UI"
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
         Left            =   3120
         TabIndex        =   63
         Top             =   3390
         Width           =   480
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
         Index           =   22
         Left            =   240
         TabIndex        =   62
         Top             =   3390
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   6240
      Width           =   11895
      Begin VB.CommandButton cmdTerapia 
         Caption         =   "&Terapia"
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
         Left            =   2520
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdKtv 
         Caption         =   "Calcola &Kt/V"
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
         Left            =   7080
         TabIndex        =   26
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
         Height          =   615
         Left            =   10440
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMostraInfo 
         Caption         =   "&Scheda Dialitica"
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
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
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
         Height          =   615
         Left            =   8760
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSchedaDialitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' questo form è utilizzato solo in compilazione

Dim rsDialisi As Recordset
Dim modifica As Boolean
Dim keyId As Long
Dim lettera As String
Dim periodo As Integer
Dim codice_rene  As Integer
Dim codice_storico_dialisi As Long
Dim intPazientiKey As Integer

Const icsCAS As String = "   X "

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        frmPannelloPeriodo.LetSenzaData = False
        frmPannelloPeriodo.Show 1
        periodo = frmPannelloPeriodo.GetPeriodo
        laData = frmPannelloPeriodo.getData
        Unload frmPannelloPeriodo
        If periodo = -1 Then
            Unload Me
            Exit Sub
        End If
        cmdTrova_Click (0)
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim k As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    lblData.BackColor = vbWhite
    For i = 0 To 1
        lblOra(i).BackColor = vbWhite
    Next i
    modifica = False
    lblData = date
    flxGriglia(0).Rows = 2
    flxGriglia(1).Rows = 2
    
    For i = 0 To 1
        With flxGriglia(i)
            .Row = 0
            For k = 0 To 3
                .Col = k
                .ColAlignment(k) = vbLeftJustify
                .CellFontBold = True
            Next k
        End With
    Next i
    tabSchede.Tab = 0
    lblCognomeMedico = tAccesso.cognome
    lblNomeMedico = tAccesso.nome
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo gestione
    Dim data As Date
    
    If tTrova.keyReturn = -1 Then   'per evitare che si apri due volte il pannello delle dialisi da fare
        Cancel = False
        Exit Sub
    End If
    
    intPazientiKey = 0
    If periodo = -1 Then
        ' nel caso abbia premuto annulla dal pannello periodo
        Cancel = False
    Else
        Unload frmDialisiDaFare
        Load frmDialisiDaFare
        frmDialisiDaFare.LetTurno = periodo
        frmDialisiDaFare.Show 1
        If tTrova.keyReturn = -1 Then
            Cancel = False
        Else
            data = laData
            Call PulisciTutto
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
            laData = data
            Call CaricaLaData
            lblCognomeMedico = tAccesso.cognome
            lblNomeMedico = tAccesso.nome
            Cancel = True
        End If
    End If
    Exit Sub
    
gestione:
    If Err.Number = 440 Then Exit Sub
End Sub

Private Sub CaricaTurno()
    Dim campo As String
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        ' cerca il turno del giorno ricercato
        ' cerca se l'orario è di mattina o pomer
        If rsDataset("AM_INIZIO" & Weekday(laData, vbMonday)) <> vbNullString Then
            campo = "AM"
            lblTurno = "MAT."
        ElseIf rsDataset("PM_INIZIO" & Weekday(laData, vbMonday)) <> vbNullString Then
            campo = "PM"
            lblTurno = "POM."
        ElseIf rsDataset("SR_INIZIO" & Weekday(laData, vbMonday)) <> vbNullString Then
            campo = "SR"
            lblTurno = "SERA"
        Else
            Exit Sub
        End If
        lblOra(0) = rsDataset(campo & "_INIZIO" & Weekday(laData, vbMonday))
        lblOra(1) = rsDataset(campo & "_FINE" & Weekday(laData, vbMonday))
    End If
    Set rsDataset = Nothing
End Sub

Private Sub CaricaScheda()
    Dim data As Date
    Dim i As Integer
    ' la data americana
    data = DateValue(Month(lblData) & "/" & Day(lblData) & "/" & Year(lblData))
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI WHERE SPECIALE=FALSE AND CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDialisi.EOF And rsDialisi.BOF) Then
        keyId = rsDialisi("KEY")
        modifica = True
        lblOra(0) = rsDialisi("ORA_INIZIO")
        lblOra(1) = rsDialisi("ORA_FINE")
        txtPesoIniziale = VirgolaOrPunto(rsDialisi("PESO_INIZIO"), ",")
        txtPesoFinale = VirgolaOrPunto(rsDialisi("PESO_FINE"), ",")
        txtIncremento = VirgolaOrPunto(rsDialisi("INCREMENTO"), ",")
        For i = 0 To 4
            txtPressioneMax(i) = rsDialisi("PA_MAX" & i + 1)
            txtPressioneMin(i) = rsDialisi("PA_MIN" & i + 1)
            txtFC(i) = rsDialisi("FC" & i + 1)
        Next i
        txtConplicanze = rsDialisi("COMPLICANZE")
        chkConferma.Value = IIf(CBool(rsDialisi("CONFERMA_SOMM")), Checked, Unchecked)
        chkErrata.Value = IIf(CBool(rsDialisi("ERRATA")), Checked, Unchecked)
        codice_storico_dialisi = rsDialisi("CODICE_STORICO_DIALISI")
    Else
        modifica = False
        'Call Pulisci
    End If
    Set rsDialisi = Nothing
End Sub

Private Function Completo() As Boolean
    Dim i As Integer
    Dim k As Integer
    Dim lista As String
    
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If lblData = "" Then
        MsgBox "La data inserita non è corretta", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPesoFinale = "" Then
        MsgBox "Inserire il peso finale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPesoIniziale = "" Then
        MsgBox "Inserire il peso iniziale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtIncremento = "" Then
        MsgBox "Inserire l'incremento ponderale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPressioneMin(0) = "" Or txtPressioneMax(0) = "" Then
        MsgBox "Inserire la pressione iniziale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtFC(0) = "" Then
        MsgBox "Inserire la frequenza cardiaca iniziale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtFC(4) = "" Then
        MsgBox "Inserire la frequenza cardiaca finale", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPressioneMin(4) = "" Or txtPressioneMax(4) = "" Then
        MsgBox "Inserire la pressione finale", vbCritical, "Attenzione"
        Exit Function
    End If
    If lblOra(0) = "" Or lblOra(1) = "" Then
        MsgBox "Inserire l'ora della seduta", vbCritical, "Attenzione"
        Exit Function
    Else
        If CDate(lblOra(0)) > CDate(lblOra(1)) Then
            MsgBox "Inserimento ora errata", vbCritical, "Attenzione"
            Exit Function
        End If
    End If
    If CLng(lblUI) > 0 And chkConferma.Value = Unchecked Then
        If MsgBox("La somministrazione di EPO non è stata confermata" & vbCrLf & _
                  "Sei sicuro di voler memorizzare la scheda dialitica?", vbQuestion + vbYesNo, "Conferma Somministrazione") = vbNo Then
            Exit Function
        End If
    End If
    For k = 0 To 1
        For i = 1 To flxGriglia(k).Rows - 1
            If flxGriglia(k).TextMatrix(i, 2) = "" Then
                lista = lista & "- " & flxGriglia(k).TextMatrix(i, 0) & vbCrLf
            End If
        Next i
    Next k
    If lista <> "" Then
        If MsgBox("I farmaci " & vbCrLf & lista & "non sono stati confermati" & vbCrLf & _
                  "Sei sicuro di voler memorizzare la scheda dialitica?", vbQuestion + vbYesNo, "Conferma Somministrazione") = vbNo Then
            Exit Function
        End If
    End If
    If CSng(VirgolaOrPunto(txtPesoIniziale, ".")) - CSng(VirgolaOrPunto(txtPesoFinale, ".")) > 6 Then
        MsgBox "Differenza peso iniziale-finale maggiore di 6 Kg", vbCritical, "Attenzione"
        Exit Function
    End If
    If Abs(CSng(VirgolaOrPunto(txtPesoIniziale, ".")) - PesoSeccoDialitico) > 5 Then
        If Not MsgBox("Il peso iniziale è troppo diverso dal peso secco prescritto." & vbCrLf & "Sei sicuro di memorizzarlo?", vbCritical + vbYesNo + vbDefaultButton2, "Attenzione") = vbYes Then
            Exit Function
        End If
    End If
    Completo = True
End Function

Private Sub PulisciLabel()
    lblTipoLinee = ""
    lblAccessoVascolare = ""
    lblAgo1 = ""
    lblAgo2 = ""
    lblFiltro = ""
    lblTipoDialisi = ""
    lblAnticoagulante(0) = ""
    lblAnticoagulante(1) = ""
    lblPesoSecco = ""
    lblUltimoPeso = ""
    lblDataUltimoPeso = ""
    lblOreDialisi = ""
    lblDoseAltroAnticoagulante = ""
    lblDoseFinale = ""
    lblDoseIniziale = ""
    lblDoseIntermedia = ""
    lblDoseUnitaMisura = ""
    lblSodio = ""
    lblPotassio = ""
    lblBicarbonato = ""
    lblCalcio = ""
    lblGlucosio = ""
    lblFlusso = ""
    lblFlussoSangue = ""
    lblSolDialitica = ""
    lblSolInfusionale = ""
    lblSolInfCc = ""
    lblCartuccia = ""
    lblEpo = ""
    lblUI = ""
    lblCognomeMedico = ""
    lblNomeMedico = ""
    lblPostazione = ""
    lblNumeroRene = ""
    lblTipo = ""
    lblTipoRene = ""
End Sub

Private Sub PulisciTutto()
    Dim i As Integer
    modifica = False
    codice_storico_dialisi = 0
    codice_rene = 0
    keyId = -1
    intPazientiKey = 0
    lblData = ""
    chkConferma.Value = False
    chkErrata.Value = False
    For i = 0 To 1
        lblOra(i) = ""
    Next i
    lblTurno = ""
    Call PulisciForm(Me)
    Call PulisciLabel
    flxGriglia(0).Rows = 1
    flxGriglia(1).Rows = 1
    flxGriglia(0).TextMatrix(0, 3) = "Note                                                                 "
    flxGriglia(1).TextMatrix(0, 3) = "Note                                                                 "
    cmdTrova(0).SetFocus
    cmdMemorizza.Enabled = False
End Sub

Private Function PesoSeccoDialitico() As Single
    ' calcola il peso secco dall'anamnesi dialitica
    On Error GoTo gestione
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT PESO_SECCO FROM ANAMNESI_DIALITICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        PesoSeccoDialitico = 0
    Else
        PesoSeccoDialitico = rsDataset("PESO_SECCO")
    End If
    Set rsDataset = Nothing
    Exit Function
gestione:
    PesoSeccoDialitico = 0
End Function

Private Function UltimoPeso(ByRef data As Date) As Single
    ' calcola l'ultimo peso secco dalla seduta precedente
    ' e da in uscita la data della seduta precedente
    On Error GoTo gestione
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT PESO_FINE, DATA FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & intPazientiKey & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        UltimoPeso = 0
        data = 0
    Else
        If rsDataset("DATA") = date And Not rsDataset.EOF Then
            rsDataset.MoveNext
        End If
        UltimoPeso = rsDataset("PESO_FINE")
        data = rsDataset("DATA")
    End If
    Set rsDataset = Nothing
    Exit Function
gestione:
    UltimoPeso = 0
    data = 0
End Function

Private Function Incremento() As String
    ' calcola l'incremento dall'ultima seduta
    Dim rsDataset As Recordset
    Dim valore As Single
    
    If CDate(lblData) < date Then
        Incremento = ""
    Else
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT PESO_FINE FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & intPazientiKey & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
        If rsDataset.EOF And rsDataset.BOF Then
            Incremento = ""
        Else
            If txtPesoIniziale <> "" Then
                valore = VirgolaOrPunto(txtPesoIniziale, ".") - CSng(rsDataset("PESO_FINE"))
                Incremento = valore
                If valore > 10 Or valore < 0 Then
                    Incremento = ""
                End If
            End If
        End If
        Set rsDataset = Nothing
    End If
End Function

Private Sub CaricaLaData()
    ' carica la data e richiama le altre sub
    lblData = laData
    Call CaricaTurno
    Call CaricaScheda
End Sub

Private Function getDurataDecimale(valore As String) As Single
    Dim valori() As String
    valori = Split(valore, ":")
    getDurataDecimale = CInt(valori(0)) + CSng(valori(1) / 60)
End Function

Private Sub cmdKtv_Click()
    ' carica la scheda del ktv precompilata
    ' passandogli i parametri da qui
    Dim diff As Single
    Dim peso As Single
    Dim durata As Single
    Dim inizio As Single
    Dim fine As Single
    
    If txtPesoIniziale = "" Then
        MsgBox "Inserire il peso iniziale", vbCritical, "Impossibile calcolare il Kt/V"
        Exit Sub
    End If
    
    If txtPesoFinale = "" Then
        MsgBox "Inserire il peso finale", vbCritical, "Impossibile calcolare il Kt/V"
        Exit Sub
    End If
    If lblOra(0) = "" Or lblOra(1) = "" Then
        MsgBox "Inserire l'ora della seduta", vbCritical, "Attenzione"
        Exit Sub
    Else
        If CDate(lblOra(0)) > CDate(lblOra(1)) Then
            MsgBox "Inserimento ora errata", vbCritical, "Attenzione"
            Exit Sub
        End If
    End If
    peso = CSng(VirgolaOrPunto(txtPesoFinale, "."))
    diff = CSng(VirgolaOrPunto(txtPesoIniziale, ".") - VirgolaOrPunto(txtPesoFinale, "."))
    inizio = getDurataDecimale(lblOra(0))
    fine = getDurataDecimale(lblOra(1))
    durata = Round(fine - inizio)
    
    Unload frmKtv
    Load frmKtv
    frmKtv.LetDiff_Peso = diff
    frmKtv.LetDurata = durata
    frmKtv.LetPeso_Post = peso
    frmKtv.LetAttiva = True
    frmKtv.LetCod_paz = intPazientiKey
    frmKtv.LetData = laData
End Sub

Private Sub chkErrata_Click()
    If chkErrata.Value = Checked Then
        chkErrata.ForeColor = vbRed
    Else
        chkErrata.ForeColor = vbBlack
    End If
End Sub

Private Sub cmdCercaOra_Click(Index As Integer)
    ' gli do piena liberta di scegliere l'orario indipendentemente dal turno
    tOrario = tpNULL
    frmOrario.Show 1
    If laOra <> "" Then lblOra(Index) = laOra
End Sub

Private Sub cmdCercaOra_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 And Shift Then
        MsgBox strConnectionStringCentro & " " & strConnectionStringTracciatura
    End If
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    If Index = 0 Then
        ' pulisce per evitare problemi
        Call PulisciTutto
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = CreaCondizione
        tTrova.condStato = "(-1) OR TRUE"
        frmTrova.Show 1
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
        Call CaricaLaData
        lblCognomeMedico = tAccesso.cognome
        lblNomeMedico = tAccesso.nome
    Else
        frmVisualizzaReni.Show 1
        If tReni.postazione <> Str(-1) Then
            codice_rene = tReni.key
            lblPostazione = tReni.postazione
            lblNumeroRene = tReni.numero_rene
            lblTipoRene = tReni.monitor
            lblTipo = tReni.Tipo
        End If
    End If
End Sub

Private Sub cmdTerapia_Click()
    tabSchede.Tab = 2
    tabSchede.Visible = Not tabSchede.Visible
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub SalvaBackup(v_Val() As Variant)
    Dim rsDataset As Recordset
    Dim v_campi() As Variant
    Dim v_valori() As Variant
    v_campi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_DIALISI", "CODICE_PAZIENTE", "ORA_INIZIO", "ORA_FINE", "PESO_INIZIO", "INCREMENTO", "PESO_FINE", _
                  "PA_MAX1", "PA_MAX2", "PA_MAX3", "PA_MAX4", "PA_MAX5", "PA_MIN1", "PA_MIN2", "PA_MIN3", "PA_MIN4", "PA_MIN5", "FC1", "FC2", "FC3", "FC4", "FC5", "COMPLICANZE", "SPECIALE", "CODICE_STORICO_DIALISI", "CONFERMA_SOMM", "ERRATA")
    v_valori = Array(tAccesso.key, date, Time, v_Val(1), v_Val(2), v_Val(4), v_Val(5), v_Val(6), v_Val(7), v_Val(8), _
                    v_Val(9), v_Val(10), v_Val(11), v_Val(12), v_Val(13), v_Val(14), v_Val(15), v_Val(16), v_Val(17), v_Val(18), v_Val(19), v_Val(20), v_Val(21), v_Val(22), v_Val(23), v_Val(26), v_Val(29), v_Val(27), v_Val(28), v_Val(30))
    Set rsDataset = New Recordset
    rsDataset.Open "BACKUP_SCHEDE_DIALISI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_campi, v_valori
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

Private Function getTempo(Tipo As Byte) As Byte
    Dim i As Integer
    i = IIf(Tipo = 1, 1, InStr(1, lblOreDialisi, "-") + 2)
    Do
        If Mid(lblOreDialisi, i, 1) <> " " Then
            getTempo = getTempo & Int(Mid(lblOreDialisi, i, 1))
        Else
            Exit Do
        End If
        i = i + 1
    Loop Until i = Len(lblOreDialisi)
End Function

Private Function GetEpo() As Integer
    Select Case lblEpo
        Case Is = ""
            GetEpo = -1
        Case Is = "ALFA"
            GetEpo = 0
        Case Is = "BETA"
            GetEpo = 1
        Case Is = "DARBO"
            GetEpo = 2
        Case Is = "MIRCERA"
            GetEpo = 3
        Case Else
            GetEpo = 4
    End Select
End Function

Private Function SalvaDatiDialisi(numKey As Integer) As Boolean
    On Error GoTo gestione
    
    Dim rsDataset As New Recordset
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
            
    ' aggiunge solo se e in inserimento perche questi dati (scheda dialitica e terapie) non vengono modificati
    If Not modifica Then
 '       strNomeTabella = "Ricette"
        v_Nomi() = Array("KEY", "CODICE_RENE", "TIPO_FILTRO", "TIPO_DIALISI", "PESO_SECCO", _
                        "ULTIMO_PESO", "DATA_PESO", "SODIO", "POTASSIO", "BICARBONATO", "CALCIO", "GLUCOSIO", "ORE_DIALISI", "MIN_DIALISI", "ANTICOAGULANTE1", "DOSE1", "DOSE_INTERMEDIA", "DOSE_FINALE", "DOSI_UNITA_MISURA", _
                        "ANTICOAGULANTE2", "DOSE2", "FLUSSO", "FLUSSO_SANGUE", "SOLUZIONE_DIALITICA", "SOLUZIONE_INFUSIONALE", _
                        "VALORE_CC", "CARTUCCIA", "EPO", "UI", "TIPO_LINEE", "ACCESSO_VASCOLARE", "TIPO_AGO1", "TIPO_AGO2")
        v_Val() = Array(GetNumero("STORICO_DIALISI_GIORNALIERA"), _
                      codice_rene, lblFiltro, lblTipoDialisi, lblPesoSecco, lblUltimoPeso, IIf(lblDataUltimoPeso = "", Null, lblDataUltimoPeso), _
                      lblSodio, lblPotassio, lblBicarbonato, lblCalcio, lblGlucosio, getTempo(1), getTempo(2), lblAnticoagulante(0), lblDoseIniziale, lblDoseIntermedia, lblDoseFinale, IIf(lblDoseUnitaMisura = "ui", 0, 1), lblAnticoagulante(1), _
                      lblDoseAltroAnticoagulante, lblFlusso, lblFlussoSangue, lblSolDialitica, lblSolInfusionale, lblSolInfCc, lblCartuccia, _
                      GetEpo, lblUI, lblTipoLinee, lblAccessoVascolare, lblAgo1, lblAgo2)
        rsDataset.Open "STORICO_DIALISI_GIORNALIERA", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        rsDataset.Close
        
'        Dim laData2 As String
'        laData2 = Format(Day(Now), "00") & "/" & Format(intValore, "00") & "/" & Format(Year(Now), "0000")
'        If date > CDate(laData2) Then
'            Dim cmCommand As New Command
'            cmCommand.CommandType = adCmdText
'            cmCommand.ActiveConnection = cnPrinc
'            cmCommand.CommandText = "Delete From " & strNomeTabella & " Where Mese=" & Month(CDate(lblData)) - 1
'            cmCommand.Execute
'        End If
        SalvaDatiDialisi = True
    Else
        ' modifica solo il rene perche potrebbe essere cambiato
        rsDataset.Open "SELECT * FROM STORICO_DIALISI_GIORNALIERA WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            rsDataset.Update "CODICE_RENE", codice_rene
        End If
        rsDataset.Close
        SalvaDatiDialisi = True
    End If
    Set rsDataset = Nothing
    Exit Function
    
gestione:
    MsgBox "Descrizione: Valore non valido", vbCritical, "Errore n°: " & Err.Number
    cnPrinc.RollbackTrans
    SalvaDatiDialisi = False
End Function

Private Function SalvaDatiTerapia() As Boolean
    On Error GoTo gestione
    
    Dim cmCommand As New Command
    Dim i As Integer
    Dim k As Integer
    Dim rsDataset As New Recordset
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim conferma As Boolean
    Dim num As Long
    
    ' cancella le precedenti se esistono
    If modifica Then
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
        cmCommand.CommandText = "DELETE * FROM STORICO_TERAPIA_DIALISI WHERE CODICE_DIALISI=" & keyId
        cmCommand.Execute
    End If
    
    ' TIPO  false=0=>intradialitica    true=1=>postdialitica
    v_Nomi = Array("KEY", "MEDICINALE", "POSOLOGIA", "CONFERMA_SOMMINISTRAZIONE", "NOTE", "TIPO", "CODICE_DIALISI")
    rsDataset.Open "STORICO_TERAPIA_DIALISI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    For i = 0 To 1
        With flxGriglia(i)
            For k = 1 To .Rows - 1
                conferma = IIf(.TextMatrix(k, 2) = icsCAS, True, False)
                num = GetNumero("STORICO_TERAPIA_DIALISI")
                v_Val = Array(num, .TextMatrix(k, 0), .TextMatrix(k, 1), conferma, .TextMatrix(k, 3) & "", i, keyId)
                'rsDataset.AddNew v_nomi, v_val
                Dim j As Integer
                rsDataset.AddNew
                For j = 0 To UBound(v_Nomi)
                    rsDataset(v_Nomi(j)) = v_Val(j)
                Next
                rsDataset.Update
            Next k
        End With
    Next i
    Set rsDataset = Nothing
    SalvaDatiTerapia = True
    Exit Function
    
gestione:
    MsgBox "Descrizione: Valore non valido", vbCritical, "Errore n°: " & Err.Number
    cnPrinc.RollbackTrans
    SalvaDatiTerapia = False
End Function

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    Dim data As Date
    Dim giorno As Integer
    Dim strSql As String
    
    If intPazientiKey = 0 Then
        Exit Sub
    Else
        cmdMemorizza.Enabled = True
    End If
    
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
    
    ' carica le info sul paziente
    ' scheda dialitica e terapia
    strSql = "SELECT    ANAMNESI_DIALITICHE.*, FILTRI.NOME AS FILTRINOME, LINEE.NOME AS LINEENOME, ACCESSI_VASCOLARI.NOME AS ACCESSI_VASCOLARINOME, " & _
            "           AGO1.NOME AS AGO1NOME, AGO2.NOME AS AGO2NOME, TIPI_DIALISI.NOME AS TIPI_DIALISINOME, ANTICOAGULANTI1.NOME AS ANTICOAGULANTI1NOME, " & _
            "           ANTICOAGULANTI2.NOME AS ANTICOAGULANTI2NOME, SOL_DIALITICHE.NOME AS SOL_DIALITICHENOME, SOL_INFUSIONALI.NOME AS SOL_INFUSIONALINOME, " & _
            "           CARTUCCE.NOME AS CARTUCCENOME " & _
            " FROM      (((((((((((ANAMNESI_DIALITICHE " & _
            "           LEFT OUTER JOIN FILTRI ON FILTRI.KEY=ANAMNESI_DIALITICHE.TIPO_FILTRO) " & _
            "           LEFT OUTER JOIN LINEE ON LINEE.KEY=ANAMNESI_DIALITICHE.TIPO_LINEE) " & _
            "           LEFT OUTER JOIN ACCESSI_VASCOLARI ON ACCESSI_VASCOLARI.KEY=ANAMNESI_DIALITICHE.ACCESSO_VASCOLARE) " & _
            "           LEFT OUTER JOIN AGO AGO1 ON AGO1.KEY=ANAMNESI_DIALITICHE.AGO1) " & _
            "           LEFT OUTER JOIN AGO AGO2 ON AGO2.KEY=ANAMNESI_DIALITICHE.AGO2) " & _
            "           LEFT OUTER JOIN TIPI_DIALISI ON TIPI_DIALISI.KEY=ANAMNESI_DIALITICHE.TIPO_DIALISI) " & _
            "           LEFT OUTER JOIN ANTICOAGULANTI ANTICOAGULANTI1 ON ANTICOAGULANTI1.KEY=ANAMNESI_DIALITICHE.ANTICOAGULANTE1) " & _
            "           LEFT OUTER JOIN ANTICOAGULANTI ANTICOAGULANTI2 ON ANTICOAGULANTI2.KEY=ANAMNESI_DIALITICHE.ANTICOAGULANTE2) " & _
            "           LEFT OUTER JOIN SOL_DIALITICHE ON SOL_DIALITICHE.KEY=ANAMNESI_DIALITICHE.SOL_DIALITICA) " & _
            "           LEFT OUTER JOIN SOL_INFUSIONALI ON SOL_INFUSIONALI.KEY=ANAMNESI_DIALITICHE.SOL_INFUSIONALE) " & _
            "           LEFT OUTER JOIN CARTUCCE ON CARTUCCE.KEY=ANAMNESI_DIALITICHE.CARTUCCIA) " & _
            "WHERE      CODICE_PAZIENTE=" & intPazientiKey
             
    Set rsDataset = New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.BOF And rsDataset.EOF) Then
        lblAnticoagulante(0) = rsDataset("ANTICOAGULANTI1NOME") & ""
        lblAnticoagulante(1) = rsDataset("ANTICOAGULANTI2NOME") & ""
        lblDoseIniziale = rsDataset("DOSE1")
        lblDoseIntermedia = rsDataset("DOSE2")
        lblDoseFinale = rsDataset("DOSE3")
        lblDoseUnitaMisura = IIf(rsDataset("DOSI_UNITA_MISURA") = 0, "ui", "cc")
        lblDoseAltroAnticoagulante = rsDataset("DOSE4")
        lblPotassio = VirgolaOrPunto(rsDataset("POTASSIO"), ",")
        lblBicarbonato = VirgolaOrPunto(rsDataset("BICARBONATO"), ",")
        lblCalcio = VirgolaOrPunto(rsDataset("CALCIO"), ",")
        lblSodio = VirgolaOrPunto(rsDataset("SODIO"), ",")
        lblGlucosio = VirgolaOrPunto(rsDataset("GLUCOSIO"), ",")
        lblPesoSecco = VirgolaOrPunto(rsDataset("PESO_SECCO"), ",")
        lblUltimoPeso = VirgolaOrPunto(UltimoPeso(data), ",")
        lblDataUltimoPeso = IIf(data = CDate("0.00.00"), "", data)
        lblFiltro = rsDataset("FILTRINOME") & ""
        lblTipoDialisi = rsDataset("TIPI_DIALISINOME") & ""
        lblOreDialisi = IIf(rsDataset("ORE") = "", 0, rsDataset("ORE")) & " h - " & IIf(rsDataset("MINUTI") = "", 0, rsDataset("MINUTI")) & " m"
        lblFlusso = VirgolaOrPunto(rsDataset("FLUSSO"), ",")
        lblFlussoSangue = VirgolaOrPunto(rsDataset("FLUSSO_SANGUE"), ",")
        lblSolDialitica = rsDataset("SOL_DIALITICHENOME") & ""
        lblSolInfusionale = rsDataset("SOL_INFUSIONALINOME") & ""
        lblSolInfCc = VirgolaOrPunto(rsDataset("SOL_INF_CC"), ",")
        lblCartuccia = rsDataset("CARTUCCENOME") & ""
        If rsDataset("EPO") = -1 Then
            lblEpo = ""
            chkConferma.Enabled = False
        Else
            lblEpo = Choose(rsDataset("EPO") + 1, "ALFA", "BETA", "DARBO", "MIRCERA", "ZETA")
            chkConferma.Enabled = True
        End If
        If rsDataset("EPO") = 2 Or rsDataset("EPO") = 3 Then
            lblUnitaMisura = "mcg"
        Else
            lblUnitaMisura = "UI"
        End If
        lblUI = rsDataset("UI")
        lblTipoLinee = rsDataset("LINEENOME") & ""
        lblAccessoVascolare = rsDataset("ACCESSI_VASCOLARINOME") & ""
        lblAgo1 = rsDataset("AGO1NOME") & ""
        lblAgo2 = rsDataset("AGO2NOME") & ""
    Else
        MsgBox "Impossibile caricare il paziente" & vbCrLf & "Dati dialitici mancanti", vbCritical, "Scheda Dialitica Giornaliera"
        Call PulisciTutto
        Exit Sub
    End If
    rsDataset.Close
    
    giorno = Weekday(laData) - 1
    giorno = IIf(giorno = 0, 7, giorno)
    ' se è una scheda gia inserita (stesso giorno di oggi) allora deve caricare le terapie storicizzate relative alla scheda e non al paziente
    data = DateValue(Month(laData) & "/" & Day(laData) & "/" & Year(laData))
    rsDataset.Open "SELECT * FROM SCHEDE_DIALISI WHERE SPECIALE=FALSE AND CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        strSql = "SELECT * FROM STORICO_TERAPIA_DIALISI WHERE CODICE_DIALISI=" & rsDataset("KEY")
    Else
        strSql = "SELECT M.NOME AS MEDICINALE, (SOMMINISTRAZIONE-1) AS TIPO, POSOLOGIA, CONFERMA_SOMMINISTRAZIONE, NOTE FROM (TERAPIE_DIALITICHE T INNER JOIN MEDICINALI M ON M.KEY=T.CODICE_MEDICINALE) WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND SOSPESA=FALSE AND (TUTTI_GIORNI=TRUE OR GIORNO" & giorno & "=TRUE) ORDER BY DATA DESC"
    End If
    rsDataset.Close
    
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    flxGriglia(0).Rows = 1
    flxGriglia(1).Rows = 1
    Do While Not rsDataset.EOF
        ' se non è stata specificata la posologia nn carica il record
        If rsDataset("TIPO") <> -1 Then
            With flxGriglia(rsDataset("TIPO"))
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsDataset("MEDICINALE")
                .TextMatrix(.Rows - 1, 1) = rsDataset("POSOLOGIA") & ""
                .TextMatrix(.Rows - 1, 2) = IIf(CBool(rsDataset("CONFERMA_SOMMINISTRAZIONE")), icsCAS, "")
                .TextMatrix(.Rows - 1, 3) = rsDataset("NOTE") & ""
            End With
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM SCHEDE_DIALISI WHERE SPECIALE=FALSE AND CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        strSql = "SELECT    RENI.* " & _
                 "FROM      (STORICO_DIALISI_GIORNALIERA " & _
                 "          INNER JOIN RENI ON STORICO_DIALISI_GIORNALIERA.CODICE_RENE=RENI.KEY) " & _
                 "WHERE     STORICO_DIALISI_GIORNALIERA.KEY=" & rsDataset("CODICE_STORICO_DIALISI")
    Else
        strSql = "SELECT    RENI.* " & _
                 "FROM      (TURNI " & _
                 "          INNER JOIN RENI ON RENI.KEY=TURNI.CODICE_RENE) " & _
                 "WHERE     CODICE_PAZIENTE=" & intPazientiKey
    End If
    rsDataset.Close
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        codice_rene = rsDataset("KEY")
        lblPostazione = rsDataset("POSTAZIONE")
        lblNumeroRene = rsDataset("NUMERO_RENE") & ""
        lblTipoRene = rsDataset("TIPO_RENE")
        lblTipo = Choose(rsDataset("TIPO") + 1, "NEG", "HCV POS", "HBV POS")
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

Private Function CreaCondizione() As String
    ' crea la condizione del form Trova
    ' fa caricare solo i pazienti che hanno il turno dialitico
    ' oggi (o pom o mat)
    Dim rsAppo As New Recordset
    Dim giorno As Integer       ' 1 lun 2 mart 3 merc ..
    Dim rsPazientiTurni As Recordset
    Dim strTurno As String
    
    Select Case periodo
        Case 1
            strTurno = "AM"
        Case 2
            strTurno = "PM"
        Case 3
            strTurno = "SR"
    End Select
        
    giorno = Weekday(laData, vbMonday)
    Set rsPazientiTurni = New Recordset
    rsPazientiTurni.Open "SELECT P.KEY, T." & strTurno & "_INIZIO" & giorno & " " & _
                         "FROM ((PAZIENTI P INNER JOIN TURNI T ON P.KEY = T.CODICE_PAZIENTE) INNER JOIN RENI R ON R.KEY=T.CODICE_RENE) " & _
                         " WHERE  NOT ( T." & strTurno & "_INIZIO" & giorno & "="""")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsAppo.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Do While Not rsPazientiTurni.EOF
        ' es: KEY IN (1,2,4)
        ' effettua il controllo sulla data fine nnn in query perche il campo nn è obbligatorio
        rsAppo.Filter = ("CODICE_PAZIENTE=" & rsPazientiTurni("KEY"))
        ' se nn esiste nn  puo effettuare la dialisi (cazzi suoi)
        If Not (rsAppo.BOF And rsAppo.EOF) Then
            If rsAppo("DATA_INIZIO") <> "" Then
                If CDate(rsAppo("DATA_INIZIO")) <= date Then
                    If rsAppo("DATA_FINE") <> "" Then
                        If CDate(rsAppo("DATA_FINE")) >= date Then
                            CreaCondizione = CreaCondizione & rsPazientiTurni("KEY") & ","
                        End If
                    Else
                        CreaCondizione = CreaCondizione & rsPazientiTurni("KEY") & ","
                    End If
                End If
            End If
        End If
        rsPazientiTurni.MoveNext
    Loop
    If CreaCondizione <> "" Then
        ' elimina la , finale e aggiunge le parentesi
        CreaCondizione = Left(CreaCondizione, Len(CreaCondizione) - 1)
        CreaCondizione = " KEY IN (" & CreaCondizione & ")"
    Else
        ' non deve trovare nessun paziente (key=-1 piezzo)
        CreaCondizione = " KEY IN (-1)"
    End If
    ' solo quelli in dialisi
    'CreaCondizione = CreaCondizione & " AND (STATO=0 OR STATO=4)"
    Set rsAppo = Nothing
    Set rsPazientiTurni = Nothing
End Function

Private Sub cmdMemorizza_Click()
    Dim v_Val(1 To 30) As Variant
    Dim v_Nomi(1 To 30) As Variant
    Dim numKey As Long
    Dim i As Integer
    
    If Completo Then
        txtConplicanze = UCase(txtConplicanze)
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "CODICE_PAZIENTE"
        v_Nomi(3) = "DATA"
        v_Nomi(4) = "ORA_INIZIO"
        v_Nomi(5) = "ORA_FINE"
        v_Nomi(6) = "PESO_INIZIO"
        v_Nomi(7) = "INCREMENTO"
        v_Nomi(8) = "PESO_FINE"
        For i = 1 To 5
            v_Nomi(8 + i) = "PA_MAX" & i
            v_Nomi(8 + 5 + i) = "PA_MIN" & i
            v_Nomi(8 + 5 + 5 + i) = "FC" & i
        Next i
        v_Nomi(24) = "CODICE_DOTTORE"
        v_Nomi(25) = "TIPO_DOTTORE"
        v_Nomi(26) = "COMPLICANZE"
        v_Nomi(27) = "CODICE_STORICO_DIALISI"
        v_Nomi(28) = "CONFERMA_SOMM"
        v_Nomi(29) = "SPECIALE"
        v_Nomi(30) = "ERRATA"
        keyId = IIf(modifica, keyId, GetNumero("SCHEDE_DIALISI"))
        v_Val(1) = keyId
        v_Val(2) = intPazientiKey
        v_Val(3) = lblData
        v_Val(4) = lblOra(0)
        v_Val(5) = lblOra(1)
        v_Val(6) = IIf(txtPesoIniziale = "", 0, txtPesoIniziale)
        v_Val(7) = IIf(txtIncremento = "", 0, txtIncremento)
        v_Val(8) = IIf(txtPesoFinale = "", 0, txtPesoFinale)
        For i = 1 To 5
            v_Val(8 + i) = IIf(txtPressioneMax(i - 1) = "", 0, txtPressioneMax(i - 1))
            v_Val(8 + 5 + i) = IIf(txtPressioneMin(i - 1) = "", 0, txtPressioneMin(i - 1))
            v_Val(8 + 5 + 5 + i) = IIf(txtFC(i - 1) = "", 0, txtFC(i - 1))
        Next i
        v_Val(24) = tAccesso.key
        v_Val(25) = tAccesso.Tipo
        v_Val(26) = txtConplicanze
        If modifica Then
            numKey = codice_storico_dialisi
        Else
            numKey = GetNumero("STORICO_DIALISI_GIORNALIERA")
        End If
        v_Val(27) = numKey
        v_Val(28) = IIf(chkConferma.Value = Checked, True, False)
        v_Val(29) = False
        v_Val(30) = IIf(chkErrata.Value = Checked, True, False)
        
        cnPrinc.BeginTrans
        Set rsDialisi = New Recordset
        If modifica Then
            rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDialisi.Update v_Nomi, v_Val
        Else
            rsDialisi.Open "SCHEDE_DIALISI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDialisi.AddNew
            i = 1
            Do While i <> UBound(v_Nomi)
                rsDialisi(v_Nomi(i)) = v_Val(i)
                i = i + 1
            Loop
            'rsDialisi.AddNew v_nomi, v_val
            rsDialisi.Update
        End If
        Set rsDialisi = Nothing

        If Not SalvaDatiDialisi(numKey) Then Exit Sub
        ' invece le terapie puo modificare il CAS
        If Not SalvaDatiTerapia Then Exit Sub
        cnPrinc.CommitTrans
        
        If TRACCIATO Then
            ' effettua il backup della scheda di dialisi in connessioni
            Call SalvaBackup(v_Val)
        End If
        Call PulisciTutto
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        cmdMemorizza.Enabled = False
        
        Dim data As String
        Unload frmDialisiDaFare
        Load frmDialisiDaFare
        frmDialisiDaFare.LetTurno = periodo
        frmDialisiDaFare.Show 1
        If tTrova.keyReturn = -1 Then
            Unload Me
        Else
            data = laData
            Call PulisciTutto
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
            laData = data
            Call CaricaLaData
            lblCognomeMedico = tAccesso.cognome
            lblNomeMedico = tAccesso.nome
        End If
        
    End If
End Sub

Private Sub cmdMostraInfo_Click()
    tabSchede.Visible = Not tabSchede.Visible
    If tabSchede.Visible = True Then
        ' mostra la caption scheda
        cmdMostraInfo.Caption = "&Scheda Giornaliera"
    Else
        cmdMostraInfo.Caption = "&Scheda Dialitica"
    End If
End Sub

Private Sub flxGriglia_Click(Index As Integer)
    If VerificaClickFlx(flxGriglia(Index)) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1, True)
        ' annulla le row e col
        flxGriglia(Index).Row = 0
        flxGriglia(Index).Col = 0
    Else
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
    End If
End Sub

Private Sub flxGriglia_DblClick(Index As Integer)
    If VerificaClickFlx(flxGriglia(Index)) = False Then Exit Sub
    With flxGriglia(Index)
        .SetFocus
        If .Col = 2 Then
            If .TextMatrix(.Row, 2) = "" Then
                .TextMatrix(.Row, 2) = icsCAS
            Else
                .TextMatrix(.Row, 2) = ""
            End If
        End If
    End With
End Sub

Private Sub txtConplicanze_GotFocus()
    txtConplicanze.BackColor = colArancione
End Sub

Private Sub txtConplicanze_LostFocus()
    txtConplicanze.BackColor = vbWhite
End Sub

Private Sub txtIncremento_GotFocus()
    txtIncremento.BackColor = colArancione
End Sub

Private Sub txtIncremento_LostFocus()
    txtIncremento.BackColor = vbWhite
End Sub

Private Sub txtPesoFinale_GotFocus()
    txtPesoFinale.BackColor = colArancione
End Sub

Private Sub txtPesoFinale_LostFocus()
    txtPesoFinale.BackColor = vbWhite
End Sub

Private Sub txtPesoIniziale_GotFocus()
    txtPesoIniziale.BackColor = colArancione
End Sub

Private Sub txtPesoIniziale_LostFocus()
    Dim incr As String
    If keyId = -1 Then
        incr = Incremento()
        If incr <> "" Then
            txtIncremento = VirgolaOrPunto(CSng(Format(incr, "##.00")), ",")
        End If
    End If
    txtPesoIniziale.BackColor = vbWhite
End Sub

' sub per il controllo dei valori numerici

Private Sub txtIncremento_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtIncremento, lettera)
End Sub

Private Sub txtIncremento_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtIncremento_Validate(Cancel As Boolean)
    If txtIncremento = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtIncremento.Text)
    End If
    If txtIncremento = "0" Then txtIncremento = ""
End Sub

Private Sub txtPesoFinale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPesoFinale, lettera)
End Sub

Private Sub txtPesoFinale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPesoFinale_Validate(Cancel As Boolean)
    If txtPesoFinale = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPesoFinale.Text)
        If Cancel = False Then
            If val(txtPesoFinale) > val(txtPesoIniziale) Then
                MsgBox "Il peso finale non può essere maggiore del peso iniziale" & vbCrLf & "Controllare i valori", vbCritical, "Attenzione"
                Cancel = True
            End If
        End If
    End If
    If txtPesoFinale = "0" Then txtPesoFinale = ""
End Sub

Private Sub txtPesoIniziale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPesoIniziale, lettera)
End Sub

Private Sub txtPesoIniziale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPesoIniziale_Validate(Cancel As Boolean)
    If txtPesoIniziale = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPesoIniziale.Text)
    End If
    If txtPesoIniziale = "0" Then txtPesoIniziale = ""
End Sub

Private Sub txtFC_LostFocus(Index As Integer)
    txtFC(Index).BackColor = vbWhite
End Sub

Private Sub txtFC_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtFC(Index), lettera)
End Sub

Private Sub txtFC_GotFocus(Index As Integer)
    txtFC(Index).BackColor = colArancione
End Sub

Private Sub txtFC_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPressioneMax_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtPressioneMax(Index), lettera)
End Sub

Private Sub txtPressioneMax_GotFocus(Index As Integer)
    txtPressioneMax(Index).BackColor = colArancione
End Sub

Private Sub txtPressioneMax_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPressioneMax_LostFocus(Index As Integer)
    txtPressioneMax(Index).BackColor = vbWhite
End Sub

Private Sub txtPressioneMin_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtPressioneMin(Index), lettera)
End Sub

Private Sub txtPressioneMin_GotFocus(Index As Integer)
    txtPressioneMin(Index).BackColor = colArancione
End Sub

Private Sub txtPressioneMin_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPressioneMin_LostFocus(Index As Integer)
    txtPressioneMin(Index).BackColor = vbWhite
End Sub

