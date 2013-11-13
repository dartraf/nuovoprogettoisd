VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmSchedaDialiticaPassate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seduta Dialitica Giornaliera - Consultazione"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSchede 
      Height          =   6135
      Left            =   120
      TabIndex        =   19
      Top             =   1640
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10821
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
      TabCaption(0)   =   "Scheda giornaliera"
      TabPicture(0)   =   "frmSchedaDialiticaPassate.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblErrata"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSchedaCompilataDa(39)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblConferma"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(22)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblUnitaMisura"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(24)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(13)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(45)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(15)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(44)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(12)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(16)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(17)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(53)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblUI"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblEpo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblComplicanze"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblPesoIniziale"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblPesoFinale"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblIncremento"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblPressioneMax(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblPressioneMax(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblPressioneMax(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblPressioneMax(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPressioneMax(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblPressioneMin(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblPressioneMin(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblPressioneMin(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblPressioneMin(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblPressioneMin(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblFC(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblFC(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lblFC(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lblFC(3)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblFC(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Line5"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Line2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Line1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label8"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label7"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label6"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label5"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label4"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label3"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "lblKtvRilevato"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lblTotSangueRilevato"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "lblPaExtracorporeo"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "lblPvExtracorporeo"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "Scheda dialitica"
      TabPicture(1)   =   "frmSchedaDialiticaPassate.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(56)"
      Tab(1).Control(1)=   "Label1(48)"
      Tab(1).Control(2)=   "lblDoseIniziale"
      Tab(1).Control(3)=   "lblDoseAltroAnticoagulante"
      Tab(1).Control(4)=   "lblDoseIntermedia"
      Tab(1).Control(5)=   "Label1(38)"
      Tab(1).Control(6)=   "lblDoseFinale"
      Tab(1).Control(7)=   "Label1(35)"
      Tab(1).Control(8)=   "lblDoseUnitaMisura"
      Tab(1).Control(9)=   "lblGlucosio"
      Tab(1).Control(10)=   "Label1(27)"
      Tab(1).Control(11)=   "Label1(30)"
      Tab(1).Control(12)=   "Label1(31)"
      Tab(1).Control(13)=   "Label1(43)"
      Tab(1).Control(14)=   "Label1(23)"
      Tab(1).Control(15)=   "Label1(32)"
      Tab(1).Control(16)=   "lblSodio"
      Tab(1).Control(17)=   "lblPotassio"
      Tab(1).Control(18)=   "lblBicarbonato"
      Tab(1).Control(19)=   "lblCalcio"
      Tab(1).Control(20)=   "Label1(18)"
      Tab(1).Control(21)=   "Label1(19)"
      Tab(1).Control(22)=   "Label1(20)"
      Tab(1).Control(23)=   "Label1(21)"
      Tab(1).Control(24)=   "lblSolDialitica"
      Tab(1).Control(25)=   "lblSolInfusionale"
      Tab(1).Control(26)=   "lblCartuccia"
      Tab(1).Control(27)=   "lblSolInfCc"
      Tab(1).Control(28)=   "Label1(6)"
      Tab(1).Control(29)=   "Label1(7)"
      Tab(1).Control(30)=   "lblFlusso"
      Tab(1).Control(31)=   "lblFlussoSangue"
      Tab(1).Control(32)=   "Label1(55)"
      Tab(1).Control(33)=   "Label1(54)"
      Tab(1).Control(34)=   "Label1(46)"
      Tab(1).Control(35)=   "Label1(34)"
      Tab(1).Control(36)=   "Label1(33)"
      Tab(1).Control(37)=   "Label1(28)"
      Tab(1).Control(38)=   "Label1(29)"
      Tab(1).Control(39)=   "Label1(42)"
      Tab(1).Control(40)=   "Label1(52)"
      Tab(1).Control(41)=   "Label1(51)"
      Tab(1).Control(42)=   "Label1(50)"
      Tab(1).Control(43)=   "Label1(49)"
      Tab(1).Control(44)=   "lblTipoLinee"
      Tab(1).Control(45)=   "lblAgo1"
      Tab(1).Control(46)=   "lblAgo2"
      Tab(1).Control(47)=   "lblFiltro"
      Tab(1).Control(48)=   "lblTipoDialisi"
      Tab(1).Control(49)=   "lblAccessoVascolare"
      Tab(1).Control(50)=   "lblAnticoagulante(0)"
      Tab(1).Control(51)=   "lblAnticoagulante(1)"
      Tab(1).Control(52)=   "lblPesoSecco"
      Tab(1).Control(53)=   "lblUltimoPeso"
      Tab(1).Control(54)=   "lblDataUltimoPeso"
      Tab(1).Control(55)=   "lblOreDialisi"
      Tab(1).ControlCount=   56
      TabCaption(2)   =   "Terapia"
      TabPicture(2)   =   "frmSchedaDialiticaPassate.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxGriglia(0)"
      Tab(2).Control(1)=   "flxGriglia(1)"
      Tab(2).Control(2)=   "Label1(36)"
      Tab(2).Control(3)=   "Label1(37)"
      Tab(2).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   5055
         Index           =   0
         Left            =   -74880
         TabIndex        =   20
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaDialiticaPassate.frx":0054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   5055
         Index           =   1
         Left            =   -69000
         TabIndex        =   21
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaDialiticaPassate.frx":00EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPvExtracorporeo 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   136
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblPaExtracorporeo 
         Alignment       =   1  'Right Justify
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
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTotSangueRilevato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
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
         Left            =   10320
         TabIndex        =   134
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblKtvRilevato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
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
         Left            =   10320
         TabIndex        =   133
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Valori Rilevati dal monitor:"
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
         Left            =   8520
         TabIndex        =   132
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Kt/v"
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
         Left            =   8760
         TabIndex        =   131
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Tot. Sangue Trattato (lt.)"
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
         TabIndex        =   130
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "C.E.C."
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
         Left            =   8760
         TabIndex        =   129
         Top             =   1770
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "P.V."
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
         Left            =   9765
         TabIndex        =   128
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "P.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9765
         TabIndex        =   127
         Top             =   1620
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   9480
         X2              =   9480
         Y1              =   1725
         Y2              =   2085
      End
      Begin VB.Line Line2 
         X1              =   9480
         X2              =   9665
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line Line5 
         X1              =   9480
         X2              =   9660
         Y1              =   2085
         Y2              =   2085
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
         Index           =   56
         Left            =   -65040
         TabIndex        =   125
         Top             =   3360
         Width           =   495
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
         Index           =   48
         Left            =   -70080
         TabIndex        =   124
         Top             =   2880
         Width           =   1365
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
         Left            =   -68640
         TabIndex        =   123
         Top             =   2880
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
         Left            =   -64320
         TabIndex        =   122
         Top             =   3360
         Width           =   1095
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
         Left            =   -65760
         TabIndex        =   121
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose Intermed."
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
         Left            =   -67440
         TabIndex        =   120
         Top             =   2880
         Width           =   1590
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
         Left            =   -63720
         TabIndex        =   119
         Top             =   2880
         Width           =   615
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
         Index           =   35
         Left            =   -65040
         TabIndex        =   118
         Top             =   2880
         Width           =   1275
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
         Left            =   -68040
         TabIndex        =   117
         Top             =   2880
         Width           =   375
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
         Left            =   -65280
         TabIndex        =   116
         Top             =   5640
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
         Index           =   27
         Left            =   -65880
         TabIndex        =   115
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label lblFC 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   7440
         TabIndex        =   106
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblFC 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6600
         TabIndex        =   105
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblFC 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   5760
         TabIndex        =   104
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblFC 
         Alignment       =   1  'Right Justify
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
         Left            =   4920
         TabIndex        =   103
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblFC 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   102
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblPressioneMin 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   7440
         TabIndex        =   101
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPressioneMin 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6600
         TabIndex        =   100
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPressioneMin 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   5760
         TabIndex        =   99
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPressioneMin 
         Alignment       =   1  'Right Justify
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
         Left            =   4920
         TabIndex        =   98
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPressioneMin 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   97
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblPressioneMax 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   7440
         TabIndex        =   96
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPressioneMax 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6600
         TabIndex        =   95
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPressioneMax 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   5760
         TabIndex        =   94
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPressioneMax 
         Alignment       =   1  'Right Justify
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
         Left            =   4920
         TabIndex        =   93
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPressioneMax 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   92
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblIncremento 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   91
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblPesoFinale 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   90
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblPesoIniziale 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   89
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblComplicanze 
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
         Height          =   1965
         Left            =   1800
         TabIndex        =   88
         Top             =   2400
         Width           =   9975
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
         Left            =   1800
         TabIndex        =   87
         Top             =   4800
         Width           =   1095
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
         TabIndex        =   86
         Top             =   4800
         Width           =   855
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
         Left            =   -70440
         TabIndex        =   85
         Top             =   5640
         Width           =   270
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
         Left            =   -72000
         TabIndex        =   84
         Top             =   5640
         Width           =   435
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
         Left            =   -69120
         TabIndex        =   83
         Top             =   5655
         Width           =   690
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
         Left            =   -67320
         TabIndex        =   82
         Top             =   5640
         Width           =   420
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
         Left            =   -74760
         TabIndex        =   81
         Top             =   5640
         Width           =   1380
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
         Left            =   -71520
         TabIndex        =   80
         Top             =   5640
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
         Left            =   -70080
         TabIndex        =   79
         Top             =   5640
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
         Left            =   -68280
         TabIndex        =   78
         Top             =   5640
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
         Left            =   -66840
         TabIndex        =   77
         Top             =   5640
         Width           =   615
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
         TabIndex        =   76
         Top             =   4290
         Width           =   1950
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
         TabIndex        =   75
         Top             =   4740
         Width           =   2220
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
         TabIndex        =   74
         Top             =   4845
         Width           =   1125
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
         TabIndex        =   73
         Top             =   5190
         Width           =   990
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
         TabIndex        =   72
         Top             =   4290
         Width           =   5775
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
         TabIndex        =   71
         Top             =   4740
         Width           =   5775
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
         TabIndex        =   70
         Top             =   5190
         Width           =   5775
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
         TabIndex        =   69
         Top             =   4800
         Width           =   615
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
         TabIndex        =   68
         Top             =   3840
         Width           =   2670
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
         TabIndex        =   67
         Top             =   3840
         Width           =   2805
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
         TabIndex        =   66
         Top             =   3840
         Width           =   615
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
         Left            =   -67560
         TabIndex        =   65
         Top             =   3840
         Width           =   615
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
         Index           =   55
         Left            =   -74760
         TabIndex        =   64
         Top             =   2400
         Width           =   1275
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
         Index           =   54
         Left            =   -74760
         TabIndex        =   63
         Top             =   1440
         Width           =   1470
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
         Index           =   46
         Left            =   -71280
         TabIndex        =   62
         Top             =   2400
         Width           =   1275
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
         Left            =   -66000
         TabIndex        =   61
         Top             =   2400
         Width           =   1365
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
         Index           =   33
         Left            =   -74760
         TabIndex        =   60
         Top             =   960
         Width           =   1335
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
         Left            =   -74760
         TabIndex        =   59
         Top             =   3360
         Width           =   2100
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
         Left            =   -74760
         TabIndex        =   58
         Top             =   2880
         Width           =   1560
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
         Left            =   -68760
         TabIndex        =   57
         Top             =   2400
         Width           =   345
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
         Left            =   -67725
         TabIndex        =   56
         Top             =   960
         Width           =   705
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
         Left            =   -67725
         TabIndex        =   55
         Top             =   480
         Width           =   705
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
         Left            =   -74760
         TabIndex        =   54
         Top             =   1920
         Width           =   2040
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
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   1380
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
         Left            =   -72480
         TabIndex        =   52
         Top             =   480
         Width           =   4455
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
         Left            =   -66840
         TabIndex        =   51
         Top             =   480
         Width           =   3495
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
         Left            =   -66840
         TabIndex        =   50
         Top             =   960
         Width           =   3495
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
         Left            =   -72480
         TabIndex        =   49
         Top             =   960
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
         Left            =   -72480
         TabIndex        =   48
         Top             =   1440
         Width           =   4455
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
         Left            =   -72480
         TabIndex        =   47
         Top             =   1920
         Width           =   4455
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
         Left            =   -72480
         TabIndex        =   46
         Top             =   2880
         Width           =   2295
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
         Left            =   -72480
         TabIndex        =   45
         Top             =   3360
         Width           =   5535
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
         Left            =   -72480
         TabIndex        =   44
         Top             =   2400
         Width           =   855
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
         Left            =   -69840
         TabIndex        =   43
         Top             =   2400
         Width           =   855
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
         Left            =   -68280
         TabIndex        =   42
         Top             =   2400
         Width           =   1335
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
         Left            =   -64440
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
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
         Left            =   3000
         TabIndex        =   40
         Top             =   960
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
         Left            =   7440
         TabIndex        =   39
         Top             =   600
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
         Left            =   6600
         TabIndex        =   38
         Top             =   600
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
         Left            =   4920
         TabIndex        =   37
         Top             =   600
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
         Left            =   3000
         TabIndex        =   36
         Top             =   1320
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
         Left            =   3480
         TabIndex        =   35
         Top             =   1680
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
         Left            =   4080
         TabIndex        =   34
         Top             =   600
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
         Left            =   5760
         TabIndex        =   33
         Top             =   600
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
         TabIndex        =   32
         Top             =   1200
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
         TabIndex        =   31
         Top             =   720
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
         Height          =   480
         Index           =   13
         Left            =   240
         TabIndex        =   30
         Top             =   1620
         Width           =   1275
         WordWrap        =   -1  'True
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
         Left            =   -73200
         TabIndex        =   29
         Top             =   480
         Width           =   2175
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
         Left            =   -67080
         TabIndex        =   28
         Top             =   480
         Width           =   2190
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
         TabIndex        =   27
         Top             =   2400
         Width           =   1470
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUnitaMisura 
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
         Left            =   3210
         TabIndex        =   26
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eritropoietina"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
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
         TabIndex        =   25
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblConferma 
         AutoSize        =   -1  'True
         Caption         =   "Somministrato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   24
         Top             =   4800
         Width           =   1380
      End
      Begin VB.Label lblSchedaCompilataDa 
         AutoSize        =   -1  'True
         Caption         =   "Scheda compilata da:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   39
         Left            =   120
         TabIndex        =   23
         Top             =   5520
         Width           =   8295
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblErrata 
         AutoSize        =   -1  'True
         Caption         =   "SCHEDA ANNULLATA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   7320
         TabIndex        =   22
         Top             =   4740
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmSchedaDialiticaPassate.frx":0184
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   170
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
         TabIndex        =   18
         Top             =   240
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
         TabIndex        =   17
         Top             =   240
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
         TabIndex        =   16
         Top             =   240
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
         Left            =   10560
         TabIndex        =   14
         Top             =   240
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
         TabIndex        =   13
         Top             =   240
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
         TabIndex        =   12
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12015
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Left            =   960
         TabIndex        =   126
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   10200
         TabIndex        =   114
         Top             =   720
         Width           =   555
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
         Left            =   4440
         TabIndex        =   113
         Top             =   720
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
         Left            =   3600
         TabIndex        =   112
         Top             =   720
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
         Index           =   4
         Left            =   240
         TabIndex        =   111
         Top             =   720
         Width           =   1170
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
         TabIndex        =   110
         Top             =   735
         Width           =   75
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
         Index           =   25
         Left            =   5520
         TabIndex        =   109
         Top             =   720
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
         Left            =   1560
         TabIndex        =   108
         Top             =   720
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
         TabIndex        =   107
         Top             =   720
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
         Left            =   7320
         TabIndex        =   8
         Top             =   285
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
         Left            =   5520
         TabIndex        =   7
         Top             =   285
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
         Left            =   3360
         TabIndex        =   6
         Top             =   285
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
         Left            =   4440
         TabIndex        =   1
         Top             =   285
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
         TabIndex        =   2
         Top             =   285
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
         Left            =   8040
         TabIndex        =   5
         Top             =   300
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
         Left            =   240
         TabIndex        =   4
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   2520
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   12015
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
         Left            =   9000
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
         Left            =   10560
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSchedaDialiticaPassate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' questo form e utilizzato solo in consultazione

Dim rsDialisi As Recordset
Dim keyId As Long                        ' utile per elimina
Dim codice_storico_dialisi As Long
Dim intPazientiKey As Integer
Const icsCAS As String = " X"

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
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
    
    oData.ConnectionString = strConnectionStringCentro
    For i = 0 To 1
        lblOra(i).BackColor = vbWhite
    Next i
    oData.data = date
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Sub CaricaMedico(codice_medico As Integer)
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM LOGIN WHERE KEY=" & codice_medico, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblSchedaCompilataDa(39) = "Scheda compilata da: " & rsDataset("COGNOME") & " " & rsDataset("NOME")
    Set rsDataset = Nothing
End Sub

Private Sub CaricaScheda()
    Dim data As Date
    Dim i As Integer
    Dim ora As Integer
    Dim strSql As String
    
    If oData.data = "" Then Exit Sub
    If intPazientiKey = 0 Then Exit Sub
    ' la data americana
    data = oData.DataAmericana
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDialisi.EOF And rsDialisi.BOF) Then
        keyId = rsDialisi("KEY")
        lblOra(0) = rsDialisi("ORA_INIZIO")
        lblOra(1) = rsDialisi("ORA_FINE")
        ora = Int(Left(lblOra(0), 2))
        
        If ora < 13 Then
            lblTurno = "MAT"
        ElseIf ora > 12 And ora <= 18 Then
            lblTurno = "POM"
        Else
            lblTurno = "SER"
        End If
        
        lblPesoIniziale = VirgolaOrPunto(rsDialisi("PESO_INIZIO"), ",")
        lblPesoFinale = VirgolaOrPunto(rsDialisi("PESO_FINE"), ",")
        lblIncremento = VirgolaOrPunto(rsDialisi("INCREMENTO"), ",")
        
        For i = 0 To 4
            lblPressioneMax(i) = rsDialisi("PA_MAX" & i + 1)
            lblPressioneMin(i) = rsDialisi("PA_MIN" & i + 1)
        Next i
        
        For i = 0 To 4
            lblFC(i) = rsDialisi("FC" & i + 1)
        Next i
        
        lblKtvRilevato = VirgolaOrPunto(rsDialisi("KTV_RILEVATO") & "", ",")
        lblTotSangueRilevato = VirgolaOrPunto(rsDialisi("TOT_SANGUE_RILEVATO") & "", ",")
        lblPaExtracorporeo = rsDialisi("PA_EXTRACORPOREA") & ""
        lblPvExtracorporeo = rsDialisi("PV_EXTRACORPOREA") & ""
        
        lblComplicanze = rsDialisi("COMPLICANZE")
        
        If CBool(rsDialisi("CONFERMA_SOMM")) Then
            lblConferma.ForeColor = &H8000&
            lblConferma = "Somministrata"
        Else
            lblConferma.ForeColor = vbRed
            lblConferma = "Non Somministrata"
        End If
        
        If CBool(rsDialisi("ERRATA")) Then
            lblErrata = "SCHEDA ANNULLATA"
            lblErrata.ForeColor = vbRed
        Else
            lblErrata = ""
        End If
        
        Call CaricaMedico(rsDialisi("CODICE_DOTTORE"))
        codice_storico_dialisi = rsDialisi("CODICE_STORICO_DIALISI")
        rsDialisi.Close
        
        strSql = "SELECT    * " & _
                " FROM      (STORICO_DIALISI_GIORNALIERA STORICO_DIALISI_GIORNALIERA " & _
                "           INNER JOIN APPARATI ON APPARATI.KEY=STORICO_DIALISI_GIORNALIERA.CODICE_RENE) " & _
                "WHERE      STORICO_DIALISI_GIORNALIERA.KEY=" & codice_storico_dialisi
        rsDialisi.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDialisi.EOF And rsDialisi.BOF) Then
            ' carica i dati della 2 scheda
            For i = 0 To 1
                lblAnticoagulante(i) = rsDialisi("ANTICOAGULANTE" & i + 1)
            Next i
            lblDoseIniziale = rsDialisi("DOSE1")
            lblDoseIntermedia = rsDialisi("DOSE_INTERMEDIA")
            lblDoseFinale = rsDialisi("DOSE_FINALE")
            lblDoseUnitaMisura = IIf(rsDialisi("DOSI_UNITA_MISURA") = 0, "UI", "cc")
            lblDoseAltroAnticoagulante = rsDialisi("DOSE2")
            lblPotassio = VirgolaOrPunto(rsDialisi("POTASSIO"), ",")
            lblBicarbonato = VirgolaOrPunto(rsDialisi("BICARBONATO"), ",")
            lblCalcio = VirgolaOrPunto(rsDialisi("CALCIO"), ",")
            lblSodio = VirgolaOrPunto(rsDialisi("SODIO"), ",")
            lblGlucosio = VirgolaOrPunto(rsDialisi("GLUCOSIO"), ",")
            lblPesoSecco = VirgolaOrPunto(rsDialisi("PESO_SECCO"), ",")
            lblUltimoPeso = VirgolaOrPunto(rsDialisi("ULTIMO_PESO"), ",")
            If Not IsNull(rsDialisi("DATA_PESO")) Then
                lblDataUltimoPeso = rsDialisi("DATA_PESO")
            End If
            lblTipoLinee = rsDialisi("TIPO_LINEE")
            lblAgo1 = rsDialisi("TIPO_AGO1")
            lblFiltro = rsDialisi("TIPO_FILTRO")
            lblAgo2 = rsDialisi("TIPO_AGO2")
            lblTipoDialisi = rsDialisi("TIPO_DIALISI")
            lblAccessoVascolare = rsDialisi("ACCESSO_VASCOLARE")
            lblOreDialisi = IIf(rsDialisi("ORE_DIALISI") = "", 0, rsDialisi("ORE_DIALISI")) & " h - " & IIf(rsDialisi("MIN_DIALISI") = "", 0, rsDialisi("MIN_DIALISI")) & " min"
            lblFlusso = VirgolaOrPunto(rsDialisi("FLUSSO"), ",")
            lblFlussoSangue = VirgolaOrPunto(rsDialisi("FLUSSO_SANGUE"), ",")
            lblSolDialitica = rsDialisi("SOLUZIONE_DIALITICA")
            lblSolInfusionale = rsDialisi("SOLUZIONE_INFUSIONALE")
            lblSolInfCc = VirgolaOrPunto(rsDialisi("VALORE_CC"), ",")
            lblCartuccia = rsDialisi("CARTUCCIA")
            If rsDialisi("EPO") = -1 Then
                lblEpo = ""
                lblConferma = ""
            Else
                lblEpo = Choose(rsDialisi("EPO") + 1, "ALFA", "BETA", "DARBO", "MIRCERA", "ZETA")
            End If
            If rsDialisi("EPO") = 2 Or rsDialisi("EPO") = 3 Then
                lblUnitaMisura = "mcg"
            Else
                lblUnitaMisura = "UI"
            End If
            lblUI = rsDialisi("UI")
            lblPostazione = rsDialisi("POSTAZIONE")
            lblNumeroRene = rsDialisi("NUMERO_APPARATO") & ""
            lblTipoRene = rsDialisi("MODELLO")
            lblTipo = Choose(rsDialisi("TP_RENE") + 1, "NEG", "HCV POS", "HBV POS")

        End If
        rsDialisi.Close
        
        rsDialisi.Open "SELECT * FROM STORICO_TERAPIA_DIALISI WHERE CODICE_DIALISI=" & keyId, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        flxGriglia(0).Rows = 1
        flxGriglia(1).Rows = 1
        Do While Not rsDialisi.EOF
            With flxGriglia(rsDialisi("TIPO"))
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsDialisi("MEDICINALE")
                .TextMatrix(.Rows - 1, 1) = rsDialisi("POSOLOGIA")
                .TextMatrix(.Rows - 1, 2) = IIf(CBool(rsDialisi("CONFERMA_SOMMINISTRAZIONE")), icsCAS, "")
                .TextMatrix(.Rows - 1, 3) = rsDialisi("NOTE")
            End With
            rsDialisi.MoveNext
        Loop
        rsDialisi.Close
    Else
        MsgBox "Il paziente: " & UCase(lblCognome) & " " & UCase(lblNome) & " non ha dialisi" & vbCrLf & _
               "dialitiche giornaliere in data: " & oData.data, vbInformation, Me.Caption
    End If
    Set rsDialisi = Nothing
End Sub

Private Sub PulisciTutto(conData As Boolean)
    Dim i As Integer
    codice_storico_dialisi = -1
    keyId = -1
    If conData Then
        oData.Pulisci
        lblCognome = ""
        lblNome = ""
        lblEta = ""
        intPazientiKey = 0
    End If
    lblConferma = ""
    lblErrata = ""
    For i = 0 To 1
        lblOra(i) = ""
    Next i
    lblTurno = ""
    Call PulisciLabel
    flxGriglia(0).TextMatrix(0, 3) = "Note                                                                 "
    flxGriglia(1).TextMatrix(0, 3) = "Note                                                                 "
    flxGriglia(0).Rows = 1
    flxGriglia(1).Rows = 1
End Sub

Private Sub PulisciLabel()
    Dim i As Integer
    
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
    lblSchedaCompilataDa(39) = "Scheda compilata da: "
    lblComplicanze = ""
    lblPesoIniziale = ""
    lblPesoFinale = ""
    lblIncremento = ""
    lblPostazione = ""
    lblNumeroRene = ""
    lblTipoRene = ""
    lblTipo = ""
    For i = 0 To 4
        lblPressioneMax(i) = ""
        lblPressioneMin(i) = ""
        lblFC(i) = ""
    Next i
End Sub

Private Sub Pulisci()
    Dim codicePaz As Integer
    codicePaz = intPazientiKey
    Call PulisciTutto(False)
    intPazientiKey = codicePaz
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    Dim i As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If
    If oData.data = "" Then
        MsgBox "Selezionare la data", vbCritical, "Attenzione"
        Exit Sub
    End If
      
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDialisi("COGNOME") & " " & rsDialisi("NOME")
    structIntestazione.sDataPaziente = rsDialisi("DATA_NASCITA")
    codiceId = rsDialisi("CODICE_ID")
    Set rsDialisi = Nothing

    strSql = "SHAPE APPEND    NEW adVarChar (20) as TURNO, " & _
                    "       NEW adVarChar (5) as ORA_ATTACCO, " & _
                    "       NEW adVarChar (5) as ORA_STACCO, " & _
                    "       NEW adVarChar (20) as DURATA, " & _
                    "       NEW adVarChar (5) as POSTAZIONE_RENE, " & _
                    "       NEW adSingle as RENE, " & _
                    "       NEW adVarChar (30) as MODELLO, " & _
                    "       NEW adVarChar (15) as TIPO, " & _
                    "       NEW adSingle as PESO_SECCO, " & _
                    "       NEW adSingle as ULTIMO_PESO, " & _
                    "       NEW adVarChar (50) AS FILTRO, " & _
                    "       NEW adSingle as SODIO, " & _
                    "       NEW adSingle as POTASSIO, " & _
                    "       NEW adSingle as BICARBONATO, " & _
                    "       NEW adSingle as CALCIO, " & _
                    "       NEW adSingle as GLUCOSIO, " & _
                    "       NEW adSingle as FLUSSO_QB, " & _
                    "       NEW adSingle as FLUSSO_QD, " & _
                    "       NEW adVarChar (50) as SOL_DIALITICA, " & _
                    "       NEW adVarChar (50) as SOL_INFUSIONALE, " & _
                    "       NEW adVarChar (20) as EPO, " & _
                    "       NEW adLongVarChar as DIARIO_INFERMIERISTICO, " & _
                    "       NEW adSingle as PESO_INIZIALE, " & _
                    "       NEW adSingle as PESO_FINALE, " & _
                    "       NEW adSingle as INCREMENTO_PONDERALE, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as PA_MAX_INIZIALE, " & _
                    "       NEW adVarChar (10) as PA_MAX_INIZIALE_1, " & _
                    "       NEW adVarChar (10) as PA_MAX_INIZIALE_2, " & _
                    "       NEW adVarChar (10) as PA_MAX_INIZIALE_3, " & _
                    "       NEW adVarChar (10) as PA_MAX_FINALE, " & _
                    "       NEW adVarChar (10) as PA_MIN_INIZIALE, " & _
                    "       NEW adVarChar (10) as PA_MIN_INIZIALE_1, " & _
                    "       NEW adVarChar (10) as PA_MIN_INIZIALE_2, " & _
                    "       NEW adVarChar (10) as PA_MIN_INIZIALE_3, " & _
                    "       NEW adVarChar (10) as PA_MIN_FINALE, " & _
                    "       NEW adVarChar (10) as FC_INIZIALE, " & _
                    "       NEW adVarChar (10) as FC_INIZIALE_1, " & _
                    "       NEW adVarChar (10) as FC_INIZIALE_2, " & _
                    "       NEW adVarChar (10) as FC_INIZIALE_3, " & _
                    "       NEW adVarChar (10) as FC_FINALE, " & _
                    "       NEW adVarChar (10) as ULTIMO_PESO_DEL, " & _
                    "       NEW adVarChar (50) as TIPO_LINEA, " & _
                    "       NEW adVarChar (50) as TIPO_DIALISI, " & _
                    "       NEW adVarChar (50) as ACCESSO_VASCOLARE, " & _
                    "       NEW adVarChar (50) as AGO1, " & _
                    "       NEW adVarChar (50) as AGO2, "
    strSql = strSql & _
                    "       NEW adVarChar (50) as ANTICOAGULANTE, " & _
                    "       NEW adVarChar (5) as DOSE_UNITA_MISURA, " & _
                    "       NEW adSingle as DOSE_INIZIALE, " & _
                    "       NEW adSingle as DOSE_INTERMEDIA, " & _
                    "       NEW adSingle as DOSE_FINALE, " & _
                    "       NEW adVarChar (20) as DOSE_ANTICOAGULANTE, " & _
                    "       NEW adVarChar (50) as ALTRO_ANTICOAGULANTE, " & _
                    "       NEW adVarChar (20) as DOSE_ALTRO_ANTICOAGULANTE, " & _
                    "       NEW adVarChar (20) as VALORE_SOL_INFUSIONALE, " & _
                    "       NEW adVarChar (50) as CARTUCCIA, " & _
                    "       NEW adLongVarChar as FARMACO_TERAPIA_INTRA, " & _
                    "       NEW adLongVarChar as POS_TERAPIA_INTRA, " & _
                    "       NEW adLongVarChar as NOTE_TERAPIA_INTRA, " & _
                    "       NEW adLongVarChar as SOMM_TERAPIA_INTRA, " & _
                    "       NEW adLongVarChar as FARMACO_TERAPIA_POST, " & _
                    "       NEW adLongVarChar as SOMM_TERAPIA_POST, " & _
                    "       NEW adLongVarChar as POS_TERAPIA_POST, " & _
                    "       NEW adVarChar (4) as KTV_RILEVATO, " & _
                    "       NEW adVarChar (4) as TOT_SANGUE_RILEVATO, " & _
                    "       NEW adVarChar (4) as PA_EXTRACORPOREO, " & _
                    "       NEW adVarChar (4) as PV_EXTRACORPOREO, " & _
                    "       NEW adLongVarChar as NOTE_TERAPIA_POST "


         
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Set rsDialisi = New Recordset
        
    With rsMain
        .AddNew
        
        .Fields("TURNO") = oData.data & " " & lblTurno
        .Fields("ORA_ATTACCO") = lblOra(0)
        .Fields("ORA_STACCO") = lblOra(1)
        
        .Fields("POSTAZIONE_RENE") = lblPostazione
        .Fields("RENE") = lblNumeroRene
        .Fields("MODELLO") = lblTipoRene
        .Fields("TIPO") = lblTipo
        
        .Fields("PESO_INIZIALE") = lblPesoIniziale
        .Fields("PESO_FINALE") = lblPesoFinale
        .Fields("INCREMENTO_PONDERALE") = lblIncremento
        
        .Fields("PA_MAX_INIZIALE") = lblPressioneMax(0)
        .Fields("PA_MAX_INIZIALE_1") = lblPressioneMax(1)
        .Fields("PA_MAX_INIZIALE_2") = lblPressioneMax(2)
        .Fields("PA_MAX_INIZIALE_3") = lblPressioneMax(3)
        .Fields("PA_MAX_FINALE") = lblPressioneMax(4)
        
        .Fields("PA_MIN_INIZIALE") = lblPressioneMin(0)
        .Fields("PA_MIN_INIZIALE_1") = lblPressioneMin(1)
        .Fields("PA_MIN_INIZIALE_2") = lblPressioneMin(2)
        .Fields("PA_MIN_INIZIALE_3") = lblPressioneMin(3)
        .Fields("PA_MIN_FINALE") = lblPressioneMin(4)
        
        .Fields("FC_INIZIALE") = lblFC(0)
        .Fields("FC_INIZIALE_1") = lblFC(1)
        .Fields("FC_INIZIALE_2") = lblFC(2)
        .Fields("FC_INIZIALE_3") = lblFC(3)
        .Fields("FC_FINALE") = lblFC(4)
        
        .Fields("KTV_RILEVATO") = lblKtvRilevato.Caption
        .Fields("TOT_SANGUE_RILEVATO") = lblTotSangueRilevato.Caption
        .Fields("PA_EXTRACORPOREO") = lblPaExtracorporeo.Caption
        .Fields("PV_EXTRACORPOREO") = lblPvExtracorporeo.Caption
        
        .Fields("DIARIO_INFERMIERISTICO") = lblComplicanze
        
        .Fields("PESO_SECCO") = lblPesoSecco
        .Fields("ULTIMO_PESO") = lblUltimoPeso
        .Fields("ULTIMO_PESO_DEL") = lblDataUltimoPeso
        .Fields("DURATA") = lblOreDialisi
        
        .Fields("TIPO_LINEA") = lblTipoLinee
        .Fields("FILTRO") = lblFiltro
        
        .Fields("TIPO_DIALISI") = lblTipoDialisi
        .Fields("SODIO") = lblSodio
        .Fields("POTASSIO") = lblPotassio
        .Fields("BICARBONATO") = lblBicarbonato
        .Fields("CALCIO") = lblCalcio
        .Fields("GLUCOSIO") = lblGlucosio
        
        .Fields("ACCESSO_VASCOLARE") = lblAccessoVascolare
        .Fields("AGO1") = lblAgo1
        .Fields("AGO2") = lblAgo2
        
        .Fields("ANTICOAGULANTE") = lblAnticoagulante(0)
        .Fields("DOSE_INIZIALE") = lblDoseIniziale
        .Fields("DOSE_UNITA_MISURA") = lblDoseUnitaMisura
        .Fields("DOSE_INTERMEDIA") = lblDoseIntermedia
        .Fields("DOSE_FINALE") = lblDoseFinale
                
        .Fields("ALTRO_ANTICOAGULANTE") = lblAnticoagulante(1)
        .Fields("DOSE_ALTRO_ANTICOAGULANTE") = lblDoseAltroAnticoagulante
        
        .Fields("FLUSSO_QB") = lblFlusso
        .Fields("FLUSSO_QD") = lblFlussoSangue
                
        .Fields("SOL_DIALITICA") = lblSolDialitica
        
        .Fields("SOL_INFUSIONALE") = lblSolInfusionale
        .Fields("VALORE_SOL_INFUSIONALE") = lblSolInfCc
        
        .Fields("CARTUCCIA") = lblCartuccia
        
        For i = 1 To flxGriglia(0).Rows - 1
            .Fields("FARMACO_TERAPIA_INTRA") = .Fields("FARMACO_TERAPIA_INTRA") & vbCrLf & flxGriglia(0).TextMatrix(i, 0)
            .Fields("POS_TERAPIA_INTRA") = .Fields("POS_TERAPIA_INTRA") & vbCrLf & flxGriglia(0).TextMatrix(i, 1)
            .Fields("NOTE_TERAPIA_INTRA") = .Fields("NOTE_TERAPIA_INTRA") & vbCrLf & flxGriglia(0).TextMatrix(i, 3)
            .Fields("SOMM_TERAPIA_INTRA") = .Fields("SOMM_TERAPIA_INTRA") & vbCrLf & IIf(flxGriglia(0).TextMatrix(i, 2) = icsCAS, "Somministrato", "Non somministrato")
        Next i
        
        For i = 1 To flxGriglia(1).Rows - 1
            .Fields("FARMACO_TERAPIA_POST") = .Fields("FARMACO_TERAPIA_POST") & vbCrLf & flxGriglia(1).TextMatrix(i, 0)
            .Fields("POS_TERAPIA_POST") = .Fields("POS_TERAPIA_POST") & vbCrLf & flxGriglia(1).TextMatrix(i, 1)
            .Fields("NOTE_TERAPIA_POST") = .Fields("NOTE_TERAPIA_POST") & vbCrLf & flxGriglia(1).TextMatrix(i, 3)
            .Fields("SOMM_TERAPIA_POST") = .Fields("SOMM_TERAPIA_POST") & vbCrLf & IIf(flxGriglia(1).TextMatrix(i, 2) = icsCAS, "Somministrato", "Non somministrato")
        Next i

    End With

    Set rptSchedaDialiticaGiornaliera.DataSource = rsMain
    rptSchedaDialiticaGiornaliera.TopMargin = 0
    rptSchedaDialiticaGiornaliera.BottomMargin = 0
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblGiorno").Caption = oData.data
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblSchedaCompilataDa").Caption = lblSchedaCompilataDa(39).Caption
    rptSchedaDialiticaGiornaliera.PrintReport True, rptRangeAllPages
    
    Set rsDialisi = Nothing
End Sub

Private Sub Timer1_Timer()
    If lblErrata.ForeColor = vbRed Then
        lblErrata.ForeColor = vbBlack
    Else
        lblErrata.ForeColor = vbRed
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

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        Exit Sub
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
    ' carica la scheda dopo aver scelto la data
End Sub

Private Sub oData_OnDataChange()
    Call Pulisci
    Call CaricaScheda
End Sub

Private Sub oData_OnDataClick()
    oData.Pulisci
End Sub

Private Sub oData_OnElencaClick()
    ' setta le variabili che saranno viste dal frmElencaDate
    tElenca.Tipo = tpSCHEDEDIALITICHE
    tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND SPECIALE=FALSE"
    frmElencaDate.Show 1
    If laData <> "" Then oData.data = laData
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto(True)
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    Call CaricaPaziente
End Sub

