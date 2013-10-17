VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmSchedaStraordinaria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seduta supplementare"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSchede 
      Height          =   6075
      Left            =   120
      TabIndex        =   64
      Top             =   1725
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10716
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
      TabPicture(0)   =   "frmSchedaStraordinaria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(39)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(22)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUnitaMisura"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(24)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(13)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(45)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(15)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(44)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(12)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(53)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblCognomeMedico"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblNomeMedico"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label7"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label5"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtConplicanze"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtUI"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkConferma"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cboEPO"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "chkErrata"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtPesoFinale"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtIncremento"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtPesoIniziale"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtFC(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtFC(3)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtFC(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtFC(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtFC(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtPressioneMin(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtPressioneMin(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtPressioneMin(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtPressioneMin(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtPressioneMin(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtPressioneMax(2)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtPressioneMax(4)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtPressioneMax(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtPressioneMax(3)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtPressioneMax(0)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Scheda dialitica"
      TabPicture(1)   =   "frmSchedaStraordinaria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboTipoFiltro"
      Tab(1).Control(1)=   "cboDosiUnitaMisura"
      Tab(1).Control(2)=   "txtDoseFinale"
      Tab(1).Control(3)=   "txtDoseIntermedia"
      Tab(1).Control(4)=   "txtGlucosio"
      Tab(1).Control(5)=   "cboTipoLinee"
      Tab(1).Control(6)=   "cboAccesso"
      Tab(1).Control(7)=   "cboTipoAgo(0)"
      Tab(1).Control(8)=   "cboTipoAgo(1)"
      Tab(1).Control(9)=   "txtPotassio"
      Tab(1).Control(10)=   "txtSodio"
      Tab(1).Control(11)=   "txtBicarbonato"
      Tab(1).Control(12)=   "txtCalcio"
      Tab(1).Control(13)=   "txtFlusso"
      Tab(1).Control(14)=   "txtFlussoSangue"
      Tab(1).Control(15)=   "cboCartuccia"
      Tab(1).Control(16)=   "cboSolInf"
      Tab(1).Control(17)=   "cboSolDialitica"
      Tab(1).Control(18)=   "cboAnticoagulante(1)"
      Tab(1).Control(19)=   "cboAnticoagulante(0)"
      Tab(1).Control(20)=   "txtMinuti"
      Tab(1).Control(21)=   "cboTipoDialisi"
      Tab(1).Control(22)=   "txtSolInfCc"
      Tab(1).Control(23)=   "txtDoseIniziale"
      Tab(1).Control(24)=   "txtDoseAltroAnticoagulante"
      Tab(1).Control(25)=   "txtOre"
      Tab(1).Control(26)=   "txtUltimoPeso"
      Tab(1).Control(27)=   "txtPesoSecco"
      Tab(1).Control(28)=   "oData(1)"
      Tab(1).Control(29)=   "Label1(55)"
      Tab(1).Control(30)=   "Label1(54)"
      Tab(1).Control(31)=   "Label1(52)"
      Tab(1).Control(32)=   "Label1(49)"
      Tab(1).Control(33)=   "Label1(48)"
      Tab(1).Control(34)=   "Label1(47)"
      Tab(1).Control(35)=   "Label1(46)"
      Tab(1).Control(36)=   "Label1(32)"
      Tab(1).Control(37)=   "Label1(30)"
      Tab(1).Control(38)=   "Label1(31)"
      Tab(1).Control(39)=   "Label1(43)"
      Tab(1).Control(40)=   "Label1(23)"
      Tab(1).Control(41)=   "Label1(6)"
      Tab(1).Control(42)=   "Label1(7)"
      Tab(1).Control(43)=   "Label1(5)"
      Tab(1).Control(44)=   "Label1(28)"
      Tab(1).Control(45)=   "Label1(18)"
      Tab(1).Control(46)=   "Label1(19)"
      Tab(1).Control(47)=   "Label1(20)"
      Tab(1).Control(48)=   "Label1(21)"
      Tab(1).Control(49)=   "Label1(29)"
      Tab(1).Control(50)=   "Label1(35)"
      Tab(1).Control(51)=   "Label1(38)"
      Tab(1).Control(52)=   "Label1(26)"
      Tab(1).Control(53)=   "Label1(25)"
      Tab(1).Control(54)=   "Label1(42)"
      Tab(1).Control(55)=   "Label1(34)"
      Tab(1).Control(56)=   "Label1(33)"
      Tab(1).Control(57)=   "Label1(27)"
      Tab(1).ControlCount=   58
      TabCaption(2)   =   "Terapia"
      TabPicture(2)   =   "frmSchedaStraordinaria.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboMedicinali"
      Tab(2).Control(1)=   "txtAppo"
      Tab(2).Control(2)=   "cmdInserisci(1)"
      Tab(2).Control(3)=   "cmdInserisci(0)"
      Tab(2).Control(4)=   "flxGriglia(0)"
      Tab(2).Control(5)=   "flxGriglia(1)"
      Tab(2).Control(6)=   "Label1(36)"
      Tab(2).Control(7)=   "Label1(37)"
      Tab(2).ControlCount=   8
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   137
         Text            =   "cboTipoAgo"
         Top             =   960
         Width           =   4215
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
         ItemData        =   "frmSchedaStraordinaria.frx":0054
         Left            =   -68160
         List            =   "frmSchedaStraordinaria.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtDoseFinale 
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
         Left            =   -63840
         MaxLength       =   6
         TabIndex        =   134
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtDoseIntermedia 
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
         Left            =   -66000
         MaxLength       =   6
         TabIndex        =   132
         Top             =   2880
         Width           =   615
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
         Left            =   -65760
         MaxLength       =   5
         TabIndex        =   130
         Top             =   5680
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
         Index           =   0
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   3
         Top             =   960
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
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   12
         Top             =   960
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   6
         Top             =   960
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
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   15
         Top             =   960
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   9
         Top             =   960
         Width           =   615
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   23
         Text            =   "cboTipoLinee"
         Top             =   480
         Width           =   4215
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   1920
         Width           =   5655
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
         Left            =   -66360
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "cboTipoAgo"
         Top             =   480
         Width           =   3135
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
         Left            =   -66360
         Sorted          =   -1  'True
         TabIndex        =   25
         Text            =   "cboTipoAgo"
         Top             =   960
         Width           =   3135
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
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1320
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
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1320
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1320
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1320
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
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1320
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
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1680
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1680
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
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1680
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1680
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
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1680
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   0
         Top             =   720
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1680
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1200
         Width           =   605
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
         Left            =   -70440
         TabIndex        =   43
         Top             =   5685
         Width           =   615
      End
      Begin VB.TextBox txtSodio 
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
         Left            =   -71880
         TabIndex        =   42
         Top             =   5685
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
         Left            =   -68640
         TabIndex        =   44
         Top             =   5685
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
         Left            =   -67320
         TabIndex        =   45
         Top             =   5685
         Width           =   615
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
         Left            =   -72000
         MaxLength       =   5
         TabIndex        =   36
         Top             =   3840
         Width           =   615
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
         Left            =   -67320
         MaxLength       =   5
         TabIndex        =   37
         Top             =   3840
         Width           =   615
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
         Left            =   9240
         TabIndex        =   22
         Top             =   4680
         Width           =   2295
      End
      Begin VB.ComboBox cboMedicinali 
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
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
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
         Left            =   -66240
         MaxLength       =   3
         TabIndex        =   86
         Top             =   1320
         Visible         =   0   'False
         Width           =   600
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
         Height          =   495
         Index           =   1
         Left            =   -69000
         TabIndex        =   56
         Top             =   5470
         Width           =   1335
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
         Height          =   495
         Index           =   0
         Left            =   -74880
         TabIndex        =   55
         Top             =   5470
         Width           =   1335
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   5205
         Width           =   5655
      End
      Begin VB.ComboBox cboSolInf 
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   4725
         Width           =   5655
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   4245
         Width           =   5655
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   34
         Text            =   "cboAnticoagulante"
         Top             =   3360
         Width           =   5655
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   32
         Text            =   "cboAnticoagulante"
         Top             =   2880
         Width           =   2055
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
         TabIndex        =   31
         Top             =   2400
         Width           =   735
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
         Left            =   -72360
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1440
         Width           =   5655
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
         ItemData        =   "frmSchedaStraordinaria.frx":006A
         Left            =   1800
         List            =   "frmSchedaStraordinaria.frx":007D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CheckBox chkConferma 
         Caption         =   "Conferma Avvenuta Somministrazione"
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
         Left            =   4920
         TabIndex        =   21
         Top             =   4680
         Width           =   4335
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
         Left            =   -65160
         MaxLength       =   5
         TabIndex        =   40
         Top             =   4725
         Width           =   615
      End
      Begin VB.TextBox txtDoseIniziale 
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
         Left            =   -68760
         MaxLength       =   6
         TabIndex        =   33
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtDoseAltroAnticoagulante 
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
         Left            =   -65640
         MaxLength       =   6
         TabIndex        =   35
         Top             =   3960
         Width           =   735
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
         Left            =   3840
         MaxLength       =   5
         TabIndex        =   20
         Top             =   4680
         Width           =   795
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
         Height          =   1965
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2400
         Width           =   9855
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
         Left            =   -65760
         MaxLength       =   2
         TabIndex        =   30
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtUltimoPeso 
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
         Left            =   -69960
         MaxLength       =   5
         TabIndex        =   29
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtPesoSecco 
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
         Left            =   -72360
         MaxLength       =   5
         TabIndex        =   28
         Top             =   2400
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   4695
         Index           =   0
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaStraordinaria.frx":00A3
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   4695
         Index           =   1
         Left            =   -69000
         TabIndex        =   47
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   $"frmSchedaStraordinaria.frx":013B
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   1
         Left            =   -68520
         TabIndex        =   139
         Top             =   2350
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Valori Rilevati dal monitor"
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
         Left            =   8400
         TabIndex        =   146
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
         Left            =   8520
         TabIndex        =   145
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
         Left            =   8520
         TabIndex        =   144
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
         Left            =   8520
         TabIndex        =   143
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
         Left            =   9525
         TabIndex        =   142
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
         Left            =   9525
         TabIndex        =   141
         Top             =   1620
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   9240
         X2              =   9240
         Y1              =   1725
         Y2              =   2085
      End
      Begin VB.Line Line2 
         X1              =   9240
         X2              =   9425
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line Line5 
         X1              =   9240
         X2              =   9420
         Y1              =   2085
         Y2              =   2085
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
         Left            =   -65160
         TabIndex        =   135
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dose Interm."
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
         Left            =   -67350
         TabIndex        =   133
         Top             =   2880
         Width           =   1320
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
         Index           =   52
         Left            =   -66360
         TabIndex        =   131
         Top             =   5685
         Width           =   480
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
         Left            =   6120
         TabIndex        =   119
         Top             =   5400
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
         Left            =   2640
         TabIndex        =   118
         Top             =   5400
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
         Left            =   2880
         TabIndex        =   116
         Top             =   960
         Width           =   915
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
         TabIndex        =   115
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
         Index           =   48
         Left            =   -74760
         TabIndex        =   114
         Top             =   1920
         Width           =   2040
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
         Index           =   47
         Left            =   -67200
         TabIndex        =   113
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
         Index           =   46
         Left            =   -67200
         TabIndex        =   112
         Top             =   480
         Width           =   705
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
         Left            =   7320
         TabIndex        =   111
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
         Left            =   6480
         TabIndex        =   110
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
         Left            =   4800
         TabIndex        =   109
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
         Left            =   2880
         TabIndex        =   108
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
         Left            =   3360
         TabIndex        =   107
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
         Left            =   3960
         TabIndex        =   106
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
         Left            =   5640
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
         Top             =   1560
         Width           =   1335
         WordWrap        =   -1  'True
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
         TabIndex        =   94
         Top             =   5680
         Width           =   1380
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
         Left            =   -70800
         TabIndex        =   93
         Top             =   5685
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
         Left            =   -72360
         TabIndex        =   92
         Top             =   5680
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
         Left            =   -69480
         TabIndex        =   91
         Top             =   5685
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
         Left            =   -67680
         TabIndex        =   90
         Top             =   5680
         Width           =   300
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
         TabIndex        =   89
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
         Left            =   -70200
         TabIndex        =   88
         Top             =   3840
         Width           =   2805
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
         Index           =   5
         Left            =   -64680
         TabIndex        =   85
         Top             =   2400
         Width           =   615
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
         TabIndex        =   84
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
         Left            =   -66840
         TabIndex        =   83
         Top             =   480
         Width           =   2190
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
         TabIndex        =   82
         Top             =   3360
         Width           =   2100
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
         TabIndex        =   81
         Top             =   4260
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
         TabIndex        =   80
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
         Left            =   -66360
         TabIndex        =   79
         Top             =   4740
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
         TabIndex        =   78
         Top             =   5215
         Width           =   990
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
         TabIndex        =   77
         Top             =   2895
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
         Index           =   35
         Left            =   -70200
         TabIndex        =   76
         Top             =   2880
         Width           =   1365
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
         Index           =   38
         Left            =   -66360
         TabIndex        =   75
         Top             =   4005
         Width           =   570
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
         Left            =   -74760
         TabIndex        =   74
         Top             =   960
         Width           =   1335
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
         Left            =   -74760
         TabIndex        =   73
         Top             =   1440
         Width           =   1470
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
         TabIndex        =   72
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
         Left            =   3360
         TabIndex        =   71
         Top             =   4680
         Width           =   240
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
         TabIndex        =   70
         Top             =   4680
         Width           =   1410
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
         Left            =   -69000
         TabIndex        =   68
         Top             =   2400
         Width           =   345
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
         Index           =   34
         Left            =   -66360
         TabIndex        =   67
         Top             =   2400
         Width           =   390
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
         Left            =   -71400
         TabIndex        =   66
         Top             =   2400
         Width           =   1275
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
         Left            =   -74760
         TabIndex        =   65
         Top             =   2400
         Width           =   1275
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
         Height          =   240
         Index           =   39
         Left            =   240
         TabIndex        =   69
         Top             =   5400
         Width           =   2295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   95
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   240
         Picture         =   "frmSchedaStraordinaria.frx":01D3
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   170
         Width           =   450
      End
      Begin VB.CheckBox chkFiltra 
         Height          =   270
         Left            =   960
         Picture         =   "frmSchedaStraordinaria.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Filtra pazienti che hanno dialisi straordinarie"
         Top             =   240
         Width           =   375
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
         Left            =   2520
         TabIndex        =   101
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
         Left            =   7080
         TabIndex        =   100
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
         Left            =   11040
         TabIndex        =   99
         Top             =   240
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
         Left            =   1440
         TabIndex        =   98
         Top             =   240
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
         Left            =   6240
         TabIndex        =   97
         Top             =   240
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
         TabIndex        =   96
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   120
      TabIndex        =   57
      Top             =   600
      Width           =   11895
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   405
         Index           =   1
         Left            =   260
         Picture         =   "frmSchedaStraordinaria.frx":0776
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   690
         Width           =   405
      End
      Begin VB.CommandButton cmdCercaOra 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   54
         Top             =   320
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaOra 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   53
         Top             =   320
         Width           =   375
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   138
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   -1  'True
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
         Top             =   795
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
         Left            =   5160
         TabIndex        =   128
         Top             =   795
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
         Index           =   4
         Left            =   4320
         TabIndex        =   127
         Top             =   795
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
         Index           =   50
         Left            =   960
         TabIndex        =   126
         Top             =   795
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
         TabIndex        =   125
         Top             =   810
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
         Index           =   51
         Left            =   6360
         TabIndex        =   124
         Top             =   795
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
         Left            =   2160
         TabIndex        =   123
         Top             =   795
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
         Left            =   7200
         TabIndex        =   122
         Top             =   795
         Width           =   2895
      End
      Begin VB.Label lblTurno 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
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
         Height          =   180
         Left            =   8880
         TabIndex        =   120
         Top             =   360
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   62
         Top             =   315
         Width           =   510
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
         Left            =   7560
         TabIndex        =   61
         Top             =   315
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
         Index           =   0
         Left            =   5160
         TabIndex        =   60
         Top             =   315
         Width           =   615
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
         Left            =   3720
         TabIndex        =   59
         Top             =   315
         Width           =   990
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
         Left            =   6240
         TabIndex        =   58
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   63
      Top             =   7680
      Width           =   11895
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
         Left            =   5640
         TabIndex        =   140
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
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
         Left            =   7440
         TabIndex        =   117
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCarica 
         Caption         =   "Carica &Eritropoietina"
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
         TabIndex        =   48
         Top             =   240
         Width           =   2175
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
         TabIndex        =   50
         Top             =   240
         Width           =   1215
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
         Left            =   8880
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSchedaStraordinaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDialisi As Recordset
Dim modifica As Boolean
Dim keyId As Integer
Dim codice_storico_dialisi As Integer
Dim codice_rene As Integer
Dim tprene As Byte
Dim lettera As String * 1
Dim numFlx As Byte
Dim intPazientiKey As Integer

Const icsCAS As String = " X"

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
    
    laData = oData(0).data
    
    Unload frmKtv
    Load frmKtv
    frmKtv.LetDiff_Peso = diff
    frmKtv.LetDurata = durata
    frmKtv.LetPeso_Post = peso
    frmKtv.LetAttiva = True
    frmKtv.LetCod_paz = intPazientiKey
    frmKtv.LetData = laData
End Sub

Private Function getDurataDecimale(valore As String) As Single
    Dim valori() As String
    valori = Split(valore, ":")
    getDurataDecimale = CInt(valori(0)) + CSng(valori(1) / 60)
End Function

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Call RicaricaComboBox("FILTRI", "NOME", cboTipoFiltro)
    Call RicaricaComboBox("LINEE", "NOME", cboTipoLinee)
    Call RicaricaComboBox("ACCESSI_VASCOLARI", "NOME", cboAccesso)
    Call RicaricaComboBox("AGO", "NOME", cboTipoAgo(0))
    Call RicaricaComboBox("AGO", "NOME", cboTipoAgo(1))
    Call RicaricaComboBox("TIPI_DIALISI", "NOME", cboTipoDialisi)
    Call RicaricaComboBox("ANTICOAGULANTI", "NOME", cboAnticoagulante(0))
    Call RicaricaComboBox("ANTICOAGULANTI", "NOME", cboAnticoagulante(1))
    Call RicaricaComboBox("SOL_DIALITICHE", "NOME", cboSolDialitica)
    Call RicaricaComboBox("SOL_INFUSIONALI", "NOME", cboSolInf)
    Call RicaricaComboBox("CARTUCCE", "NOME", cboCartuccia)
    Call RicaricaComboBox("MEDICINALI", "NOME", cboMedicinali)
    
    If intPazientiKey = 0 Then
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
    
    For i = 0 To 1
        lblOra(i).BackColor = vbWhite
        oData(i).ConnectionString = strConnectionStringCentro
    Next i
    modifica = False
    flxGriglia(0).Rows = 2
    flxGriglia(1).Rows = 2
    For i = 0 To 1
        With flxGriglia(i)
            .Rows = 1
            .Row = 0
            For k = 0 To 3
                .Col = k
                .ColAlignment(k) = vbLeftJustify
                .CellFontBold = True
            Next k
        End With
    Next i
    tabSchede.Tab = 0
    cboDosiUnitaMisura.ListIndex = 0
    lblCognomeMedico = tAccesso.cognome
    lblNomeMedico = tAccesso.nome
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Sub CaricaMedico(codice_medico As Integer)
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM LOGIN WHERE KEY=" & codice_medico, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognomeMedico = rsDataset("COGNOME")
    lblNomeMedico = rsDataset("NOME")
    Set rsDataset = Nothing
End Sub

Private Sub CaricaScheda()
    Dim data As Date
    Dim i As Integer
    Dim ora As Integer
    Dim strSql As String
    
    If oData(0).data = "" Then Exit Sub
    If intPazientiKey = 0 Then Exit Sub
    
    Call Pulisci
    ' la data americana
    data = oData(0).DataAmericana
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND DATA=#" & data & "# AND SPECIALE=TRUE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDialisi.EOF And rsDialisi.BOF) Then
        keyId = rsDialisi("KEY")
        modifica = True
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
        Call CaricaMedico(rsDialisi("CODICE_DOTTORE"))
        codice_storico_dialisi = rsDialisi("CODICE_STORICO_DIALISI")
        rsDialisi.Close
        
        strSql = "SELECT    STORICO_DIALISI_GIORNALIERA.*, " & _
                 "          APPARATI.KEY AS RENIKEY, APPARATI.* " & _
                 "FROM      (STORICO_DIALISI_GIORNALIERA " & _
                 "          INNER JOIN APPARATI ON APPARATI.KEY=STORICO_DIALISI_GIORNALIERA.CODICE_RENE) " & _
                 "WHERE     STORICO_DIALISI_GIORNALIERA.KEY=" & codice_storico_dialisi
        rsDialisi.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDialisi.EOF And rsDialisi.BOF) Then
            ' carica i dati della 2 scheda
            txtDoseIniziale = VirgolaOrPunto(rsDialisi("DOSE1"), ",")
            txtDoseIntermedia = VirgolaOrPunto(rsDialisi("DOSE_INTERMEDIA"), ",")
            txtDoseFinale = VirgolaOrPunto(rsDialisi("DOSE_FINALE"), ",")
            cboDosiUnitaMisura.ListIndex = rsDialisi("DOSI_UNITA_MISURA")
            txtDoseAltroAnticoagulante = VirgolaOrPunto(rsDialisi("DOSE2"), ",")
            txtPotassio = VirgolaOrPunto(rsDialisi("POTASSIO"), ",")
            txtBicarbonato = VirgolaOrPunto(rsDialisi("BICARBONATO"), ",")
            txtCalcio = VirgolaOrPunto(rsDialisi("CALCIO"), ",")
            txtSodio = VirgolaOrPunto(rsDialisi("SODIO"), ",")
            txtGlucosio = VirgolaOrPunto(rsDialisi("GLUCOSIO"), ",")
            txtPesoSecco = VirgolaOrPunto(rsDialisi("PESO_SECCO"), ",")
            txtUltimoPeso = VirgolaOrPunto(rsDialisi("ULTIMO_PESO"), ",")
            If Not IsNull(rsDialisi("DATA_PESO")) Then
                oData(1).data = rsDialisi("DATA_PESO")
            End If
            
            cboTipoFiltro.ListIndex = GetIndex(cboTipoFiltro, rsDialisi("TIPO_FILTRO"))
            If cboTipoFiltro.ListIndex = -1 Then cboTipoFiltro.Text = rsDialisi("TIPO_FILTRO") & ""
            cboTipoDialisi.ListIndex = GetIndex(cboTipoDialisi, rsDialisi("TIPO_DIALISI"))
            If cboTipoDialisi.ListIndex = -1 Then cboTipoDialisi.Text = rsDialisi("TIPO_DIALISI") & ""
            cboTipoLinee.ListIndex = GetIndex(cboTipoLinee, rsDialisi("TIPO_LINEE"))
            If cboTipoLinee.ListIndex = -1 Then cboTipoLinee.Text = rsDialisi("TIPO_LINEE") & ""
            cboAccesso.ListIndex = GetIndex(cboAccesso, rsDialisi("ACCESSO_VASCOLARE"))
            If cboAccesso.ListIndex = -1 Then cboAccesso.Text = rsDialisi("ACCESSO_VASCOLARE") & ""
            cboTipoAgo(0).ListIndex = GetIndex(cboTipoAgo(0), rsDialisi("TIPO_AGO1"))
            If cboTipoAgo(0).ListIndex = -1 Then cboTipoAgo(0).Text = rsDialisi("TIPO_AGO1") & ""
            cboTipoAgo(1).ListIndex = GetIndex(cboTipoAgo(1), rsDialisi("TIPO_AGO2"))
            If cboTipoAgo(1).ListIndex = -1 Then cboTipoAgo(1).Text = rsDialisi("TIPO_AGO2") & ""
            cboSolDialitica.ListIndex = GetIndex(cboSolDialitica, rsDialisi("SOLUZIONE_DIALITICA"))
            If cboSolDialitica.ListIndex = -1 Then cboSolDialitica.Text = rsDialisi("SOLUZIONE_DIALITICA") & ""
            cboSolInf.ListIndex = GetIndex(cboSolInf, rsDialisi("SOLUZIONE_INFUSIONALE"))
            If cboSolInf.ListIndex = -1 Then cboSolInf.Text = rsDialisi("SOLUZIONE_INFUSIONALE") & ""
            cboCartuccia.ListIndex = GetIndex(cboCartuccia, rsDialisi("CARTUCCIA"))
            If cboCartuccia.ListIndex = -1 Then cboCartuccia.Text = rsDialisi("CARTUCCIA") & ""
            For i = 0 To 1
                cboAnticoagulante(i).ListIndex = GetIndex(cboAnticoagulante(i), rsDialisi("ANTICOAGULANTE" & i + 1))
                If cboAnticoagulante(i).ListIndex = -1 Then cboAnticoagulante(i).Text = rsDialisi("ANTICOAGULANTE" & i + 1) & ""
            Next i
            
            txtOre = IIf(rsDialisi("ORE_DIALISI") = "", 0, rsDialisi("ORE_DIALISI"))
            txtMinuti = IIf(rsDialisi("MIN_DIALISI") = "", 0, rsDialisi("MIN_DIALISI"))
            txtFlusso = VirgolaOrPunto(rsDialisi("FLUSSO"), ",")
            txtFlussoSangue = VirgolaOrPunto(rsDialisi("FLUSSO_SANGUE"), ",")
            txtSolInfCc = VirgolaOrPunto(rsDialisi("VALORE_CC"), ",")
            cboEPO.ListIndex = rsDialisi("EPO")
            txtUI = rsDialisi("UI")
            codice_rene = rsDialisi("RENIKEY")
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
        cmdStampa.Enabled = True
    Else
        modifica = False
    End If
    Set rsDialisi = Nothing
End Sub

Private Sub Pulisci()
    Dim codicePaz As Integer
    codicePaz = intPazientiKey
    Call PulisciTutto(False)
    intPazientiKey = codicePaz
End Sub

Private Sub PulisciTutto(conData As Boolean)
    Dim i As Integer
    codice_storico_dialisi = -1
    codice_rene = 0
    keyId = -1
    If conData Then
        laData = ""
        oData(0).Pulisci
        lblCognome = ""
        lblNome = ""
        lblEta = ""
        intPazientiKey = 0
    End If
    For i = 0 To 1
        lblOra(i) = ""
    Next i
    oData(1).Pulisci
    chkConferma.Value = Unchecked
    chkErrata.Value = Unchecked
    chkFiltra.Value = Unchecked
    Call PulisciForm(Me, conData)
    flxGriglia(0).TextMatrix(0, 3) = "Note                                                                 "
    flxGriglia(1).TextMatrix(0, 3) = "Note                                                                 "
    flxGriglia(0).Rows = 1
    flxGriglia(1).Rows = 1
    lblCognomeMedico = tAccesso.cognome
    lblNomeMedico = tAccesso.nome
    lblPostazione = ""
    lblNumeroRene = ""
    lblTipo = ""
    lblTipoRene = ""
    cmdStampa.Enabled = False
    lblTurno = ""
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

Private Function Completo() As Boolean
    Dim k As Integer
    Dim i As Integer
    Dim lista As String
    
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData(0).data = "" Then
        MsgBox "La data non è stata specificata", vbCritical, "Attenzione"
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
    If txtPressioneMin(4) = "" Or txtPressioneMax(4) = "" Then
        MsgBox "Inserire la pressione finale", vbCritical, "Attenzione"
        Exit Function
    End If
    If lblOra(0) = "" Or lblOra(1) = "" Then
        MsgBox "Inserire l'orario di inizio/fine seduta", vbCritical, "Attenzione"
        Exit Function
    Else
        If CDate(lblOra(0)) > CDate(lblOra(1)) Then
            MsgBox "Inserimento ora errata", vbCritical, "Attenzione"
            Exit Function
        End If
    End If
    If CLng(txtUI) > 0 And chkConferma.Value = Unchecked Then
        If MsgBox("La somministrazione di EPO non è stata confermata" & vbCrLf & _
                  "Sei sicuro di voler memorizzare la scheda dialitica?", vbQuestion + vbYesNo, "Conferma Somministrazione") = vbNo Then
            Exit Function
        End If
    End If
    If codice_rene = 0 Then
        MsgBox "Inserire la postazione rene", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboTipoFiltro.Text = "" Then
        MsgBox "Selezionare il tipo di filtro", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboTipoDialisi.Text = "" Then
        MsgBox "Selezionare il tipo di dialisi", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtUltimoPeso = "" Then
        MsgBox "Inserire l'ultimo peso", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData(1).data = "" Then
        MsgBox "Inserire la data relativa all'ultimo peso", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtSodio = "" Then
        MsgBox "Inserire il valore del campo SODIO", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPotassio = "" Then
        MsgBox "Inserire il valore del campo POTASSIO", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtBicarbonato = "" Then
        MsgBox "Inserire il valore del campo BICARBONATO", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtCalcio = "" Then
        MsgBox "Inserire il valore del campo CALCIO", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboAnticoagulante(0).Text = "" Then
        MsgBox "Selezionare il tipo di anticoagulante", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboAccesso.Text = "" Then
        MsgBox "Selezionare il tipo di accesso vascolare", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtDoseIniziale = "" Then
        MsgBox "Inserire il valore della dose iniziale relativa all'anticoagulante", vbCritical, "Attenzione"
        Exit Function
    End If
    If txtPesoSecco = "" Then
        MsgBox "Inserire il peso secco", vbCritical, "Attenzione"
        Exit Function
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
    If CSng(VirgolaOrPunto(txtPesoIniziale, ".")) - CSng(VirgolaOrPunto(txtPesoFinale, ".")) > 7 Then
        MsgBox "Differenza peso iniziale-finale maggiore di 7 Kg", vbCritical, "Attenzione"
        Exit Function
    End If
    If Abs(CSng(VirgolaOrPunto(txtPesoIniziale, ".")) - PesoSeccoDialitico) > 5 Then
        If Not MsgBox("Il peso iniziale è troppo diverso dal peso secco prescritto." & vbCrLf & "Sei sicuro di memorizzarlo?", vbCritical + vbYesNo + vbDefaultButton2, "Attenzione") = vbYes Then
            Exit Function
        End If
    End If
    Completo = True
End Function

Private Function Incremento() As String
    ' calcola l'incremento dall'ultima seduta
    Dim rsDataset As Recordset
    Dim valore As Single
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
End Function

Private Function SalvaDatiDialisi() As Boolean
    On Error GoTo gestione
    
    Dim rsDataset As New Recordset
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim numKey As Integer
    If modifica Then
        numKey = codice_storico_dialisi
    Else
        numKey = GetNumero("STORICO_DIALISI_GIORNALIERA")
    End If
    
    ' punta il TIPO di rene (NEG-HCV+/HBV+)
    rsDataset.Open "SELECT TIPO FROM APPARATI WHERE KEY=" & codice_rene, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        tprene = rsDataset("TIPO")
    rsDataset.Close
    
    v_Nomi() = Array("KEY", "CODICE_RENE", "TP_RENE", "TIPO_FILTRO", "TIPO_DIALISI", "PESO_SECCO", _
                    "ULTIMO_PESO", "DATA_PESO", "SODIO", "POTASSIO", "BICARBONATO", "CALCIO", "GLUCOSIO", "ORE_DIALISI", "MIN_DIALISI", "ANTICOAGULANTE1", "DOSE1", "DOSE_INTERMEDIA", "DOSE_FINALE", "DOSI_UNITA_MISURA", _
                    "ANTICOAGULANTE2", "DOSE2", "FLUSSO", "FLUSSO_SANGUE", "SOLUZIONE_DIALITICA", "SOLUZIONE_INFUSIONALE", _
                    "VALORE_CC", "CARTUCCIA", "EPO", "UI", "TIPO_LINEE", "ACCESSO_VASCOLARE", "TIPO_AGO1", "TIPO_AGO2")
    v_Val() = Array(numKey, codice_rene, tprene, cboTipoFiltro.Text, cboTipoDialisi.Text, _
                    txtPesoSecco, txtUltimoPeso, IIf(oData(1).data = "", Null, oData(1).data), txtSodio, txtPotassio, txtBicarbonato, txtCalcio, txtGlucosio, txtOre, txtMinuti, cboAnticoagulante(0).Text, txtDoseIniziale, txtDoseIntermedia, txtDoseFinale, cboDosiUnitaMisura.ListIndex, _
                    cboAnticoagulante(1).Text, txtDoseAltroAnticoagulante, txtFlusso, txtFlussoSangue, cboSolDialitica.Text, cboSolInf.Text, _
                    txtSolInfCc, cboCartuccia.Text, cboEPO.ListIndex, txtUI, cboTipoLinee.Text, cboAccesso.Text, cboTipoAgo(0).Text, cboTipoAgo(1).Text)
    If modifica Then
        rsDataset.Open "SELECT * FROM STORICO_DIALISI_GIORNALIERA WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsDataset.Update v_Nomi, v_Val
    Else
        rsDataset.Open "STORICO_DIALISI_GIORNALIERA", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
    End If
    Set rsDataset = Nothing
    SalvaDatiDialisi = True
    Exit Function
    
gestione:
    MsgBox "Descrizione: Valore non valido", vbCritical, "Errore n°: " & Err.Number
    cnPrinc.RollbackTrans
    SalvaDatiDialisi = False
End Function

Private Sub Completa()
    lettera = "0"
    If txtUI = "" Then txtUI = 0
    'If txtUltimoPeso = "" Then txtUltimoPeso = "0"
    'If txtPesoSecco = "" Then txtPesoSecco = "0"
    'If txtSodio = "" Then txtSodio = "0"
    If txtFlusso = "" Then txtFlusso = "0"
    If txtFlussoSangue = "" Then txtFlussoSangue = "0"
    If txtSolInfCc = "" Then txtSolInfCc = "0"
    'If txtPotassio = "" Then txtPotassio = "0"
    If txtGlucosio = "" Then txtGlucosio = "0"
    If txtDoseAltroAnticoagulante = "" Then txtDoseAltroAnticoagulante = "0"
    If txtDoseFinale = "" Then txtDoseFinale = "0"
    If txtDoseIniziale = "" Then txtDoseIniziale = "0"
    If txtDoseIntermedia = "" Then txtDoseIntermedia = "0"
    If txtOre = "" Then txtOre = "0"
    If txtMinuti = "" Then txtMinuti = "0"
    If cboDosiUnitaMisura.ListIndex = -1 Then cboDosiUnitaMisura.ListIndex = 0
End Sub

Private Function SalvaDatiTerapia() As Boolean
    On Error GoTo gestione
    
    Dim i As Integer
    Dim k As Integer
    Dim conferma As Boolean
    Dim rsDataset As New Recordset
    Dim cmCommand As New Command
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    
    ' elimina prima tutte le terapie collegate a questa scheda
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    cmCommand.CommandText = "DELETE * FROM STORICO_TERAPIA_DIALISI WHERE CODICE_DIALISI=" & keyId
    cmCommand.Execute

    ' poi le ricrea
    ' TIPO  false=0=>intradialitica    true=1=>postdialitica
    v_Nomi = Array("KEY", "MEDICINALE", "POSOLOGIA", "CONFERMA_SOMMINISTRAZIONE", "NOTE", "TIPO", "CODICE_DIALISI")
    rsDataset.Open "STORICO_TERAPIA_DIALISI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    For i = 0 To 1
        With flxGriglia(i)
            For k = 1 To .Rows - 1
                conferma = IIf(.TextMatrix(k, 2) = icsCAS, True, False)
                v_Val = Array(GetNumero("STORICO_TERAPIA_DIALISI"), .TextMatrix(k, 0), .TextMatrix(k, 1), conferma, .TextMatrix(k, 3), i, keyId)
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

Private Sub flxGriglia_Click(Index As Integer)
    Dim vCol As Integer
    flxGriglia(Index).SetFocus
    If VerificaClickFlx(flxGriglia(Index)) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1, True)
        ' annulla le row e col
        flxGriglia(Index).Row = 0
        flxGriglia(Index).Col = 0
    Else
        vCol = flxGriglia(Index).Col
        Call ColoraFlx(flxGriglia(Index), flxGriglia(Index).Cols - 1)
        flxGriglia(Index).Col = vCol
    End If
End Sub

Private Sub flxGriglia_DblClick(Index As Integer)
    If VerificaClickFlx(flxGriglia(Index)) = False Then Exit Sub
    With flxGriglia(Index)
        .SetFocus
        numFlx = Index
        Select Case flxGriglia(Index).Col
            Case 0  ' medicinali
                cboMedicinali.Left = .colPos(.Col) + .Left + 45
                cboMedicinali.Top = .rowPos(.Row) + .Top + 45
                cboMedicinali.ListIndex = GetIndex(cboMedicinali, .TextMatrix(.Row, .Col))
                cboMedicinali.Visible = True
                cboMedicinali.SetFocus
            Case 1, 3 ' posologia, note
                If .Col = 1 Then
                    txtAppo.MaxLength = 3
                Else
                    txtAppo.MaxLength = 0
                End If
                txtAppo.Left = .colPos(.Col) + .Left + 45
                txtAppo.Top = .rowPos(.Row) + .Top + 45
                txtAppo.Width = .ColWidth(.Col)
                txtAppo.Text = .TextMatrix(.Row, .Col)
                txtAppo.Visible = True
                txtAppo.SetFocus
            Case 2  'cas
                If .TextMatrix(.Row, .Col) = "" Then
                    .TextMatrix(.Row, .Col) = icsCAS
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
        End Select
    End With
End Sub

Private Sub flxGriglia_Scroll(Index As Integer)
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia(Index).rowPos(flxGriglia(Index).Row) + flxGriglia(Index).Top + 45
    End If
    If cboMedicinali.Visible Then
        cboMedicinali.Top = flxGriglia(Index).rowPos(flxGriglia(Index).Row) + flxGriglia(Index).Top + 45
    End If
End Sub

Private Sub cboMedicinali_Click()
    flxGriglia(numFlx).TextMatrix(flxGriglia(numFlx).Row, flxGriglia(numFlx).Col) = cboMedicinali.Text
    cboMedicinali.Visible = False
End Sub

Private Sub cboMedicinali_LostFocus()
    numFlx = 0
    cboMedicinali.Visible = False
End Sub

Private Sub cboEPO_Click()
    If cboEPO.ListIndex = -1 Then
        chkConferma.Enabled = False
    Else
        chkConferma.Enabled = True
    End If
    If cboEPO.ListIndex = 2 Or cboEPO.ListIndex = 3 Then
        lblUnitaMisura = "mcg"
    Else
        lblUnitaMisura = "UI"
    End If
End Sub

Private Sub chkErrata_Click()
    If chkErrata.Value = Checked Then
        chkErrata.ForeColor = vbRed
    Else
        chkErrata.ForeColor = vbBlack
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdInserisci_Click(Index As Integer)
    If intPazientiKey = 0 Then Exit Sub
    Unload frmInput
    tInput.Tipo = tpITERAPIESTRAORDINARIE
    frmInput.Show 1
    If Not (tInput.v_valori(1) = -1) Then
        With flxGriglia(Index)
            ' aggiorna solo la flx
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = cboMedicinali.List(tInput.v_valori(1))
            .TextMatrix(.Rows - 1, 1) = tInput.v_valori(2)
            .TextMatrix(.Rows - 1, 2) = IIf(tInput.v_valori(3), icsCAS, "")
            .TextMatrix(.Rows - 1, 3) = tInput.v_valori(4)
            ' si posiziona sul record (ultimo) e lo seleziona
            .Row = .Rows - 1
            Call ColoraFlx(flxGriglia(Index), .Cols - 1)
        End With
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    Dim i As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
        
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
                    "       NEW adLongVarChar as NOTE_TERAPIA_POST "


         
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Set rsDialisi = New Recordset
        
    With rsMain
        .AddNew
        
        .Fields("TURNO") = oData(0).data & " " & lblTurno
        .Fields("ORA_ATTACCO") = lblOra(0)
        .Fields("ORA_STACCO") = lblOra(1)
        
        .Fields("POSTAZIONE_RENE") = lblPostazione
        .Fields("RENE") = lblNumeroRene
        .Fields("MODELLO") = lblTipoRene
        .Fields("TIPO") = lblTipo
                             
        .Fields("PESO_INIZIALE") = txtPesoIniziale
        .Fields("PESO_FINALE") = txtPesoFinale
        .Fields("INCREMENTO_PONDERALE") = txtIncremento
        
        .Fields("PA_MAX_INIZIALE") = txtPressioneMax(0)
        .Fields("PA_MAX_INIZIALE_1") = txtPressioneMax(1)
        .Fields("PA_MAX_INIZIALE_2") = txtPressioneMax(2)
        .Fields("PA_MAX_INIZIALE_3") = txtPressioneMax(3)
        .Fields("PA_MAX_FINALE") = txtPressioneMax(4)
        
        .Fields("PA_MIN_INIZIALE") = txtPressioneMin(0)
        .Fields("PA_MIN_INIZIALE_1") = txtPressioneMin(1)
        .Fields("PA_MIN_INIZIALE_2") = txtPressioneMin(2)
        .Fields("PA_MIN_INIZIALE_3") = txtPressioneMin(3)
        .Fields("PA_MIN_FINALE") = txtPressioneMin(4)
        
        .Fields("FC_INIZIALE") = txtFC(0)
        .Fields("FC_INIZIALE_1") = txtFC(1)
        .Fields("FC_INIZIALE_2") = txtFC(2)
        .Fields("FC_INIZIALE_3") = txtFC(3)
        .Fields("FC_FINALE") = txtFC(4)
        
        .Fields("DIARIO_INFERMIERISTICO") = txtConplicanze
        
        .Fields("PESO_SECCO") = txtPesoSecco
        .Fields("ULTIMO_PESO") = txtUltimoPeso
        .Fields("ULTIMO_PESO_DEL") = oData(1).data
        .Fields("DURATA") = txtOre & " h - " & txtMinuti & " min"
        
        .Fields("TIPO_LINEA") = cboTipoLinee.Text
        .Fields("FILTRO") = cboTipoFiltro.Text
                
        .Fields("TIPO_DIALISI") = cboTipoDialisi.Text
        .Fields("SODIO") = txtSodio
        .Fields("POTASSIO") = txtPotassio
        .Fields("BICARBONATO") = txtBicarbonato
        .Fields("CALCIO") = txtCalcio
        .Fields("GLUCOSIO") = txtGlucosio
        
        .Fields("ACCESSO_VASCOLARE") = cboAccesso.Text
        .Fields("AGO1") = cboTipoAgo(0).Text
        .Fields("AGO2") = cboTipoAgo(1).Text
        
        .Fields("ANTICOAGULANTE") = cboAnticoagulante(0)
        .Fields("DOSE_INIZIALE") = txtDoseIniziale
        .Fields("DOSE_UNITA_MISURA") = cboDosiUnitaMisura
        .Fields("DOSE_INTERMEDIA") = txtDoseIntermedia
        .Fields("DOSE_FINALE") = txtDoseFinale
                
        .Fields("ALTRO_ANTICOAGULANTE") = cboAnticoagulante(1)
        .Fields("DOSE_ALTRO_ANTICOAGULANTE") = txtDoseAltroAnticoagulante
        
        .Fields("FLUSSO_QB") = txtFlusso
        .Fields("FLUSSO_QD") = txtFlussoSangue
                
        .Fields("SOL_DIALITICA") = cboSolDialitica.Text
        
        .Fields("SOL_INFUSIONALE") = cboSolInf.Text
        .Fields("VALORE_SOL_INFUSIONALE") = txtSolInfCc
        
        .Fields("CARTUCCIA") = cboCartuccia.Text
        
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
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblGiorno").Caption = oData(0).data
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptSchedaDialiticaGiornaliera.Sections("Intestazione").Controls.Item("lblSchedaCompilataDa").Caption = lblCognomeMedico & " " & lblNomeMedico
    rptSchedaDialiticaGiornaliera.PrintReport True, rptRangeAllPages
    
    Set rsDialisi = Nothing
End Sub

Private Sub cmdCarica_Click()
    Dim i As Integer
    Dim data As Date
    Dim giorno As Integer
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    Dim strSql As String
    
    Select Case tabSchede.Tab
        Case 0
            rsDataset.Open "SELECT * FROM ANAMNESI_DIALITICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.BOF And rsDataset.EOF) Then
                cboEPO.ListIndex = rsDataset("EPO")
                txtUI = rsDataset("UI")
            End If
            rsDataset.Close
        Case 1
            rsDataset.Open "SELECT * FROM ANAMNESI_DIALITICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.BOF And rsDataset.EOF) Then
                For i = 0 To 1
                    cboAnticoagulante(i).ListIndex = GetCboListIndex(rsDataset("ANTICOAGULANTE" & i + 1), cboAnticoagulante(i))
                Next i
                txtDoseIniziale = VirgolaOrPunto(rsDataset("DOSE1"), ",")
                txtDoseIntermedia = VirgolaOrPunto(rsDataset("DOSE2"), ",")
                txtDoseFinale = VirgolaOrPunto(rsDataset("DOSE3"), ",")
                txtDoseAltroAnticoagulante = VirgolaOrPunto(rsDataset("DOSE4"), ",")
                cboDosiUnitaMisura.ListIndex = rsDataset("DOSI_UNITA_MISURA")
                txtPotassio = VirgolaOrPunto(rsDataset("POTASSIO"), ",")
                txtBicarbonato = VirgolaOrPunto(rsDataset("BICARBONATO"), ",")
                txtCalcio = VirgolaOrPunto(rsDataset("CALCIO"), ",")
                txtSodio = VirgolaOrPunto(rsDataset("SODIO"), ",")
                txtGlucosio = VirgolaOrPunto(rsDataset("GLUCOSIO"), ",")
                txtPesoSecco = VirgolaOrPunto(rsDataset("PESO_SECCO"), ",")
                txtUltimoPeso = VirgolaOrPunto(UltimoPeso(data), ",")
                txtFlusso = VirgolaOrPunto(rsDataset("FLUSSO"), ",")
                txtFlussoSangue = VirgolaOrPunto(rsDataset("FLUSSO_SANGUE"), ",")
                oData(1).data = IIf(data = CDate("0.00.00"), "", data)
                txtOre = IIf(rsDataset("ORE") = "", 0, rsDataset("ORE"))
                txtMinuti = IIf(rsDataset("MINUTI") = "", 0, rsDataset("MINUTI"))
                cboTipoFiltro.ListIndex = GetCboListIndex(rsDataset("TIPO_FILTRO"), cboTipoFiltro)
                cboTipoDialisi.ListIndex = GetCboListIndex(rsDataset("TIPO_DIALISI"), cboTipoDialisi)
                cboCartuccia.ListIndex = GetCboListIndex(rsDataset("CARTUCCIA"), cboCartuccia)
                cboSolDialitica.ListIndex = GetCboListIndex(rsDataset("SOL_DIALITICA"), cboSolDialitica)
                cboSolInf.ListIndex = GetCboListIndex(rsDataset("SOL_INFUSIONALE"), cboSolInf)
                For i = 0 To 1
                    cboAnticoagulante(i).ListIndex = GetCboListIndex(rsDataset("ANTICOAGULANTE" & i + 1), cboAnticoagulante(i))
                    cboTipoAgo(i).ListIndex = GetCboListIndex(rsDataset("AGO" & i + 1), cboTipoAgo(i))
                Next i
                cboTipoLinee.ListIndex = GetCboListIndex(rsDataset("TIPO_LINEE"), cboTipoLinee)
                cboAccesso.ListIndex = GetCboListIndex(rsDataset("ACCESSO_VASCOLARE"), cboAccesso)
            End If
            rsDataset.Close
        Case 2
            If oData(0).data = "" Then
                MsgBox "Inserire la data", vbInformation, "Carica scheda terapie"
                Exit Sub
            End If
            giorno = Weekday(oData(0).data, vbMonday)
            
            strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
                     "FROM      (TERAPIE_DIALITICHE " & _
                     "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
                     "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                     "          SOSPESA=FALSE AND " & _
                     "          (TUTTI_GIORNI=TRUE OR GIORNO" & giorno & "=TRUE) " & _
                     "ORDER BY  DATA DESC"
            rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            flxGriglia(0).Rows = 1
            flxGriglia(1).Rows = 1
            Do While Not rsDataset.EOF
                ' se non è stata specificata la posologia nn carica il record
                If rsDataset("SOMMINISTRAZIONE") <> 0 Then
                    With flxGriglia(rsDataset("SOMMINISTRAZIONE") - 1)
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = rsDataset("MEDICINALINOME")
                        .TextMatrix(.Rows - 1, 1) = rsDataset("POSOLOGIA")
                        .TextMatrix(.Rows - 1, 2) = IIf(CBool(rsDataset("CONFERMA_SOMMINISTRAZIONE")), icsCAS, "")
                        .TextMatrix(.Rows - 1, 3) = rsDataset("NOTE")
                    End With
                End If
                rsDataset.MoveNext
            Loop
            rsDataset.Close
    End Select
    Set rsDataset = Nothing
End Sub

Private Sub cmdCercaOra_Click(Index As Integer)
    ' gli do piena liberta di scegliere l'orario indipendentemente dal turno
    tOrario = tpNULL
    frmOrario.Show 1
    If laOra <> "" Then lblOra(Index) = laOra
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val(1 To 30) As Variant
    Dim v_Nomi(1 To 30) As Variant
    Dim numKey As Integer
    Dim i As Integer
    Call Completa
    If Completo Then
        If cboTipoDialisi.Text <> "" Then Call GestisciNuovo("TIPI_DIALISI", cboTipoDialisi)
        If cboCartuccia.Text <> "" Then Call GestisciNuovo("CARTUCCE", cboCartuccia)
        If cboSolDialitica.Text <> "" Then Call GestisciNuovo("SOL_DIALITICHE", cboSolDialitica)
        If cboSolInf.Text <> "" Then Call GestisciNuovo("SOL_INFUSIONALI", cboSolInf)
        If cboAccesso.Text <> "" Then Call GestisciNuovo("ACCESSI_VASCOLARI", cboAccesso)
        If cboAnticoagulante(0).Text <> "" Then Call GestisciNuovo("ANTICOAGULANTI", cboAnticoagulante(0))
        If cboAnticoagulante(1).Text <> "" Then Call GestisciNuovo("ANTICOAGULANTI", cboAnticoagulante(1))
        If cboTipoAgo(0).Text <> "" Then Call GestisciNuovo("AGO", cboTipoAgo(0))
        If cboTipoAgo(1).Text <> "" Then Call GestisciNuovo("AGO", cboTipoAgo(1))
        If cboTipoFiltro.Text <> "" Then Call GestisciNuovo("FILTRI", cboTipoFiltro)
        If cboTipoLinee.Text <> "" Then Call GestisciNuovo("LINEE", cboTipoLinee)
        
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
        v_Val(3) = oData(0).data
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
        v_Val(29) = True
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
        
        ' salva sempre
        If Not SalvaDatiDialisi Then Exit Sub
        If Not SalvaDatiTerapia Then Exit Sub
        cnPrinc.CommitTrans
        
        If TRACCIATO Then
            ' effettua il backup della scheda di dialisi in connessioni
            Call SalvaBackup(v_Val)
        End If
        Call PulisciTutto(True)
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        cmdMemorizza.Enabled = False
    End If
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
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
End Sub

Private Sub lblOra_Click(Index As Integer)
    lblOra(Index) = ""
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    If Index = 0 Then
        Call CaricaScheda
    End If
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    If oData(1).data >= "" Then
        oData(Index).Pulisci
        Exit Sub
    End If
    
    oData(Index).Pulisci
    Call Pulisci
End Sub

Private Sub oData_OnElencaClick(Index As Integer)
    ' setta le variabili che saranno viste dal frmElencaDate
    tElenca.Tipo = tpSCHEDEDIALITICHE
    tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND SPECIALE=TRUE"
    frmElencaDate.Show 1
    If laData <> "" Then oData(0).data = laData
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    Dim strFiltro As String
    Dim Filtra As Boolean
    
    If Index = 0 Then
        Filtra = (chkFiltra.Value = Checked)
        ' pulisce per evitare problemi
        Call PulisciTutto(True)
        chkFiltra.Value = IIf(Filtra, Checked, Unchecked)
        tTrova.Tipo = tpPAZIENTE
        If Filtra Then
            strFiltro = " (KEY IN (SELECT DISTINCT P.KEY FROM (PAZIENTI P INNER JOIN SCHEDE_DIALISI S ON S.CODICE_PAZIENTE=P.KEY) WHERE SPECIALE=TRUE)) "
        End If
        tTrova.condizione = strFiltro
        tTrova.condStato = "(-1,0,1,2,3,4)"
        frmTrova.Show 1
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
        lblCognomeMedico = tAccesso.cognome
        lblNomeMedico = tAccesso.nome
    Else
        frmVisualizzaReni.Show 1
        If tReni.postazione <> Str(-1) Then
            codice_rene = tReni.key
            lblPostazione = tReni.postazione
            lblNumeroRene = tReni.numero_apparato
            lblTipoRene = tReni.monitor
            lblTipo = tReni.Tipo
        End If
    End If
End Sub

Private Sub tabSchede_Click(PreviousTab As Integer)
    Select Case tabSchede.Tab
        Case 0
            cmdCarica.Caption = "&Carica Eritropoietina"
        Case 1
            cmdCarica.Caption = "&Carica scheda dialitica"
        Case 2
            cmdCarica.Caption = "&Carica scheda terapia"
    End Select
    
    cboTipoLinee.SelLength = 0
    cboTipoFiltro.SelLength = 0
    cboTipoAgo(0).SelLength = 0
    cboTipoAgo(1).SelLength = 0
    cboTipoDialisi.SelLength = 0
    cboAccesso.SelLength = 0
    cboAnticoagulante(0).SelLength = 0
    cboAnticoagulante(1).SelLength = 0
    cboSolDialitica.SelLength = 0
    cboSolInf.SelLength = 0
    cboCartuccia.SelLength = 0
    cboTipoFiltro.SelLength = 0
    cboAnticoagulante(0).SelLength = 0
    cboAnticoagulante(1).SelLength = 0
    cboAccesso.SelLength = 0
    cboTipoLinee.SelLength = 0
    cboTipoAgo(0).SelLength = 0
    cboTipoAgo(1).SelLength = 0
End Sub

Private Sub txtAppo_LostFocus()
    If Not (flxGriglia(numFlx).Col = 1 And txtAppo = "") Then
        If UCase(flxGriglia(numFlx).TextMatrix(flxGriglia(numFlx).Row, flxGriglia(numFlx).Col)) <> UCase(txtAppo) Then
            flxGriglia(numFlx).TextMatrix(flxGriglia(numFlx).Row, flxGriglia(numFlx).Col) = txtAppo.Text
        End If
    End If
    numFlx = 0
    txtAppo.Visible = False
End Sub

Private Sub txtAppo_GotFocus()
    If flxGriglia(numFlx).Col = 1 Then
        txtAppo.Alignment = 1 'destra per i numeri
    Else
        txtAppo.Alignment = 0 'sinistra
    End If
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia(numFlx).SetFocus
    End If
    If flxGriglia(numFlx).Col = 1 Then
        Select Case KeyAscii
            Case Asc("0") To Asc("9"), vbKeyBack
            Case Asc(" "), vbKeyBack
            Case Else
                Beep
                KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txtConplicanze_GotFocus()
    txtConplicanze.BackColor = colArancione
End Sub

Private Sub txtConplicanze_LostFocus()
    txtConplicanze.BackColor = vbWhite
End Sub

' sub per il controllo dei valori numerici

Private Sub txtDoseAltroAnticoagulante_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtDoseAltroAnticoagulante, lettera)
End Sub

Private Sub txtDoseAltroAnticoagulante_GotFocus()
    txtDoseAltroAnticoagulante.BackColor = colArancione
End Sub

Private Sub txtDoseAltroAnticoagulante_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtDoseAltroAnticoagulante_LostFocus()
    txtDoseAltroAnticoagulante.BackColor = vbWhite
End Sub

Private Sub txtDoseAltroAnticoagulante_Validate(Cancel As Boolean)
    If txtDoseAltroAnticoagulante.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtDoseAltroAnticoagulante.Text)
    End If
End Sub

Private Sub txtDoseFinale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtDoseFinale, lettera)
End Sub

Private Sub txtDoseFinale_GotFocus()
    txtDoseFinale.BackColor = colArancione
End Sub

Private Sub txtDoseFinale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtDoseFinale_LostFocus()
    txtDoseFinale.BackColor = vbWhite
End Sub

Private Sub txtDoseFinale_Validate(Cancel As Boolean)
    If txtDoseFinale.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtDoseFinale.Text)
    End If
End Sub

Private Sub txtDoseIniziale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtDoseIniziale, lettera)
End Sub

Private Sub txtDoseIniziale_GotFocus()
    txtDoseIniziale.BackColor = colArancione
End Sub

Private Sub txtDoseIniziale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtDoseIniziale_LostFocus()
    txtDoseIniziale.BackColor = vbWhite
End Sub

Private Sub txtDoseIniziale_Validate(Cancel As Boolean)
    If txtDoseIniziale.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtDoseIniziale.Text)
    End If
End Sub

Private Sub txtDoseIntermedia_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtDoseIntermedia, lettera)
End Sub

Private Sub txtDoseIntermedia_GotFocus()
    txtDoseIntermedia.BackColor = colArancione
End Sub

Private Sub txtDoseIntermedia_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtDoseIntermedia_LostFocus()
    txtDoseIntermedia.BackColor = vbWhite
End Sub

Private Sub txtDoseIntermedia_Validate(Cancel As Boolean)
    If txtDoseIntermedia.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtDoseIntermedia.Text)
    End If
End Sub

Private Sub txtFlusso_GotFocus()
    txtFlusso.BackColor = colArancione
End Sub

Private Sub txtFlusso_LostFocus()
    txtFlusso.BackColor = vbWhite
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

Private Sub txtIncremento_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtIncremento, lettera)
End Sub

Private Sub txtIncremento_GotFocus()
    txtIncremento.BackColor = colArancione
End Sub

Private Sub txtIncremento_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtIncremento_LostFocus()
    txtIncremento.BackColor = vbWhite
End Sub

Private Sub txtIncremento_Validate(Cancel As Boolean)
    If txtIncremento = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtIncremento.Text)
    End If
    If txtPesoFinale = "0" Then txtPesoFinale = ""
End Sub

Private Sub txtMinuti_GotFocus()
    txtMinuti.BackColor = colArancione
End Sub

Private Sub txtMinuti_LostFocus()
    txtMinuti.BackColor = vbWhite
End Sub

Private Sub txtOre_GotFocus()
    txtOre.BackColor = colArancione
End Sub

Private Sub txtOre_LostFocus()
    txtOre.BackColor = vbWhite
End Sub

Private Sub txtPesoFinale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPesoFinale, lettera)
End Sub

Private Sub txtPesoFinale_GotFocus()
    txtPesoFinale.BackColor = colArancione
End Sub

Private Sub txtPesoFinale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPesoFinale_LostFocus()
    txtPesoFinale.BackColor = vbWhite
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

Private Sub txtPesoIniziale_GotFocus()
    txtPesoIniziale.BackColor = colArancione
End Sub

Private Sub txtPesoIniziale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
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

Private Sub txtPesoIniziale_Validate(Cancel As Boolean)
    If txtPesoIniziale = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPesoIniziale.Text)
    End If
    If txtPesoIniziale = "0" Then txtPesoIniziale = ""
End Sub

Private Sub txtPesoSecco_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPesoSecco, lettera)
End Sub

Private Sub txtPesoSecco_GotFocus()
    txtPesoSecco.BackColor = colArancione
End Sub

Private Sub txtPesoSecco_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPesoSecco_LostFocus()
    txtPesoSecco.BackColor = vbWhite
End Sub

Private Sub txtPesoSecco_Validate(Cancel As Boolean)
    If txtPesoSecco = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPesoSecco.Text)
    End If
End Sub

Private Sub txtPotassio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtPotassio, lettera)
End Sub

Private Sub txtBicarbonato_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtBicarbonato, lettera)
End Sub

Private Sub txtPotassio_GotFocus()
    txtPotassio.BackColor = colArancione
End Sub

Private Sub txtBicarbonato_GotFocus()
    txtBicarbonato.BackColor = colArancione
End Sub

Private Sub txtPotassio_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtBicarbonato_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPotassio_LostFocus()
    txtPotassio.BackColor = vbWhite
End Sub

Private Sub txtBicarbonato_LostFocus()
    txtBicarbonato.BackColor = vbWhite
End Sub

Private Sub txtPotassio_Validate(Cancel As Boolean)
    If txtPotassio = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtPotassio.Text)
    End If
End Sub

Private Sub txtBicarbonato_Validate(Cancel As Boolean)
    If txtBicarbonato = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtBicarbonato.Text)
    End If
End Sub

Private Sub txtCalcio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtCalcio, lettera)
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

Private Sub txtPressioneMax_GotFocus(Index As Integer)
    txtPressioneMax(Index).BackColor = colArancione
End Sub

Private Sub txtPressioneMax_LostFocus(Index As Integer)
    txtPressioneMax(Index).BackColor = vbWhite
End Sub

Private Sub txtPressioneMin_GotFocus(Index As Integer)
    txtPressioneMin(Index).BackColor = colArancione
End Sub

Private Sub txtPressioneMin_LostFocus(Index As Integer)
    txtPressioneMin(Index).BackColor = vbWhite
End Sub

Private Sub txtFC_GotFocus(Index As Integer)
    txtFC(Index).BackColor = colArancione
End Sub

Private Sub txtFC_LostFocus(Index As Integer)
    txtFC(Index).BackColor = vbWhite
End Sub

Private Sub txtSodio_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtSodio, lettera)
End Sub

Private Sub txtSodio_GotFocus()
    txtSodio.BackColor = colArancione
End Sub

Private Sub txtSodio_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtSodio_LostFocus()
    txtSodio.BackColor = vbWhite
End Sub

Private Sub txtSodio_Validate(Cancel As Boolean)
    If txtSodio = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtSodio.Text)
    End If
End Sub

Private Sub txtFlussoSangue_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtFlussoSangue, lettera)
End Sub

Private Sub txtFlussoSangue_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtFlussoSangue_Validate(Cancel As Boolean)
    If txtFlussoSangue.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtFlussoSangue.Text)
    End If
End Sub

Private Sub txtFlusso_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtFlusso, lettera)
End Sub

Private Sub txtFlusso_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtFlusso_Validate(Cancel As Boolean)
    If txtFlusso.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtFlusso.Text)
    End If
End Sub

Private Sub txtSolInfCc_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtSolInfCc, lettera)
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
    If txtSolInfCc.Text = "" Then
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

Private Sub txtUltimoPeso_Change()
    If lettera = "" Or lettera = "." Then Exit Sub
    Call OnlyNumber(txtUltimoPeso, lettera)
End Sub

Private Sub txtUltimoPeso_GotFocus()
    txtUltimoPeso.BackColor = colArancione
End Sub

Private Sub txtUltimoPeso_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtUltimoPeso_LostFocus()
    txtUltimoPeso.BackColor = vbWhite
End Sub

Private Sub txtUltimoPeso_Validate(Cancel As Boolean)
    If txtUltimoPeso.Text = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtUltimoPeso.Text)
    End If
End Sub

Private Sub txtMinuti_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtOre_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPressioneMax_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtPressioneMax(Index), lettera)
End Sub

Private Sub txtPressioneMax_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtFC_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtFC(Index), lettera)
End Sub

Private Sub txtFC_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtPressioneMin_Change(Index As Integer)
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtPressioneMin(Index), lettera)
End Sub

Private Sub txtPressioneMin_KeyPress(Index As Integer, KeyAscii As Integer)
    lettera = Chr(KeyAscii)
End Sub

