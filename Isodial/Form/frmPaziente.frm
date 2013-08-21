VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmPaziente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ANAGRAFICA GENERALE"
   ClientHeight    =   8655
   ClientLeft      =   855
   ClientTop       =   1605
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabScheda 
      Height          =   7815
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Scheda Paziente"
      TabPicture(0)   =   "frmPaziente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(15)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(17)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(19)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(20)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(21)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(22)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(24)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblCodiceFiscale"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(26)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(27)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(28)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(30)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(31)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(43)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(45)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(25)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(29)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(46)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(48)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label1(49)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblEta"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(23)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "oData(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "oData(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "oData(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chkEsenteReddito"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtNome"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtCognome"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtCitta"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtCAP(0)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtProv(0)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtCAP(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtProv(1)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtIndirizzo"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtTelefono"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtEmail"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCellulare"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtFax"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtNumeroProcura"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtRilascioCarta"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtCodiceFiscale"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cboCentroProv"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtProfessione"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtNote"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cboDocumento"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cboStato"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cboGSanguigno"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtNumCarta"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtAllergia"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "panRH"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "cboAsl"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "picMostraStato"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtCodiceId"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cboDistretto"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cboEsenzione"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "cboComuneResidenza"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cboRegione"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtKm"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtTesseraSanitaria"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "cboAccompagnatore"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "chkTrasportoInAmbulanza"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "cmdTrova(0)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "panSesso"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "cboNazione"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).ControlCount=   81
      TabCaption(1)   =   "Medico di Base Associato"
      TabPicture(1)   =   "frmPaziente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRiceve"
      Tab(1).Control(1)=   "cboTipologia"
      Tab(1).Control(2)=   "chkPresenzaBarCode"
      Tab(1).Control(3)=   "txtCodiceRegionaleMedico"
      Tab(1).Control(4)=   "txtFaxMedico"
      Tab(1).Control(5)=   "txtEmailMedico"
      Tab(1).Control(6)=   "txtCellulareMedico"
      Tab(1).Control(7)=   "txtTelefonoMedico"
      Tab(1).Control(8)=   "txtStudioMedico"
      Tab(1).Control(9)=   "txtIndirizzoMedico"
      Tab(1).Control(10)=   "txtProvMedico"
      Tab(1).Control(11)=   "txtCapMedico"
      Tab(1).Control(12)=   "txtCittaMedico"
      Tab(1).Control(13)=   "txtNomeMedico"
      Tab(1).Control(14)=   "txtCognomeMedico"
      Tab(1).Control(15)=   "cmdTrova(1)"
      Tab(1).Control(16)=   "Label1(47)"
      Tab(1).Control(17)=   "Label1(50)"
      Tab(1).Control(18)=   "lblTipologiaMedico(47)"
      Tab(1).Control(19)=   "Label1(42)"
      Tab(1).Control(20)=   "Label1(38)"
      Tab(1).Control(21)=   "Label1(41)"
      Tab(1).Control(22)=   "Label1(40)"
      Tab(1).Control(23)=   "Label1(39)"
      Tab(1).Control(24)=   "Label1(37)"
      Tab(1).Control(25)=   "Label1(36)"
      Tab(1).Control(26)=   "Label1(35)"
      Tab(1).Control(27)=   "Label1(34)"
      Tab(1).Control(28)=   "Label1(33)"
      Tab(1).Control(29)=   "Label1(32)"
      Tab(1).ControlCount=   30
      Begin VB.TextBox txtRiceve 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   117
         Top             =   3000
         Width           =   3495
      End
      Begin VB.ComboBox cboTipologia 
         Enabled         =   0   'False
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
         Left            =   -67200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   3000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CheckBox chkPresenzaBarCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Barcode Cod.Fisc. su ricetta"
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
         Left            =   -68520
         TabIndex        =   116
         Top             =   2580
         Width           =   3975
      End
      Begin VB.TextBox txtCodiceRegionaleMedico 
         Enabled         =   0   'False
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
         MaxLength       =   7
         TabIndex        =   115
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox txtFaxMedico 
         Enabled         =   0   'False
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
         Left            =   -67440
         MaxLength       =   31
         TabIndex        =   114
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtEmailMedico 
         Enabled         =   0   'False
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
         MaxLength       =   31
         TabIndex        =   113
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtCellulareMedico 
         Enabled         =   0   'False
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
         Left            =   -67440
         MaxLength       =   31
         TabIndex        =   112
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox txtTelefonoMedico 
         Enabled         =   0   'False
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
         MaxLength       =   31
         TabIndex        =   111
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox txtStudioMedico 
         Enabled         =   0   'False
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
         Left            =   -67440
         MaxLength       =   31
         TabIndex        =   110
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtIndirizzoMedico 
         Enabled         =   0   'False
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
         MaxLength       =   31
         TabIndex        =   109
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtProvMedico 
         Enabled         =   0   'False
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
         Left            =   -64920
         MaxLength       =   2
         TabIndex        =   108
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtCapMedico 
         Enabled         =   0   'False
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
         Left            =   -67440
         MaxLength       =   5
         TabIndex        =   107
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtCittaMedico 
         Enabled         =   0   'False
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
         MaxLength       =   31
         TabIndex        =   106
         Top             =   900
         Width           =   3495
      End
      Begin VB.TextBox txtNomeMedico 
         Enabled         =   0   'False
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
         Left            =   -67440
         MaxLength       =   31
         TabIndex        =   105
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtCognomeMedico 
         Enabled         =   0   'False
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
         MaxLength       =   31
         TabIndex        =   104
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox cboNazione 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   470
         Width           =   3495
      End
      Begin VB.Frame panSesso 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   7920
         TabIndex        =   98
         Top             =   1320
         Width           =   1455
         Begin VB.OptionButton optSesso 
            Caption         =   "M"
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
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   10
            Width           =   735
         End
         Begin VB.OptionButton optSesso 
            Caption         =   "F"
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
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   3
            Top             =   10
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   1800
         Picture         =   "frmPaziente.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   860
         Width           =   450
      End
      Begin VB.CheckBox chkTrasportoInAmbulanza 
         Alignment       =   1  'Right Justify
         Caption         =   "Trasporto in Ambulanza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6450
         TabIndex        =   37
         Top             =   7365
         Width           =   3015
      End
      Begin VB.ComboBox cboAccompagnatore 
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
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   7320
         Width           =   3495
      End
      Begin VB.TextBox txtTesseraSanitaria 
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
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   29
         Top             =   6000
         Width           =   3735
      End
      Begin VB.TextBox txtKm 
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
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboRegione 
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
         ItemData        =   "frmPaziente.frx":0491
         Left            =   2520
         List            =   "frmPaziente.frx":0493
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2220
         Width           =   3495
      End
      Begin VB.ComboBox cboComuneResidenza 
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
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2640
         Width           =   3495
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
         ItemData        =   "frmPaziente.frx":0495
         Left            =   7920
         List            =   "frmPaziente.frx":0497
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox cboDistretto 
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
         ItemData        =   "frmPaziente.frx":0499
         Left            =   5400
         List            =   "frmPaziente.frx":049B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtCodiceId 
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
         Left            =   10920
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1380
         Width           =   735
      End
      Begin VB.PictureBox picMostraStato 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   58
         Top             =   6460
         Width           =   375
      End
      Begin VB.ComboBox cboAsl 
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
         ItemData        =   "frmPaziente.frx":049D
         Left            =   2520
         List            =   "frmPaziente.frx":05D6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Frame panRH 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   9120
         TabIndex        =   56
         Top             =   6360
         Width           =   2175
         Begin VB.OptionButton optRh 
            Caption         =   "Pos"
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
            Left            =   480
            TabIndex        =   32
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optRh 
            Caption         =   "Neg"
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
            Left            =   1320
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RH"
            BeginProperty Font 
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
            Left            =   0
            TabIndex        =   57
            Top             =   120
            Width           =   600
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtAllergia 
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
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   35
         Top             =   6915
         Width           =   3735
      End
      Begin VB.TextBox txtNumCarta 
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
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   20
         Top             =   4335
         Width           =   3735
      End
      Begin VB.ComboBox cboGSanguigno 
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
         ItemData        =   "frmPaziente.frx":0777
         Left            =   7920
         List            =   "frmPaziente.frx":0787
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   6450
         Width           =   735
      End
      Begin VB.ComboBox cboStato 
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
         ItemData        =   "frmPaziente.frx":0798
         Left            =   2520
         List            =   "frmPaziente.frx":079A
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   6480
         Width           =   3495
      End
      Begin VB.ComboBox cboDocumento 
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
         ItemData        =   "frmPaziente.frx":079C
         Left            =   2520
         List            =   "frmPaziente.frx":07A9
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4320
         Width           =   3495
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
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   34
         Top             =   6900
         Width           =   3495
      End
      Begin VB.TextBox txtProfessione 
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
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   27
         Top             =   5580
         Width           =   3735
      End
      Begin VB.ComboBox cboCentroProv 
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
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox txtCodiceFiscale 
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
         MaxLength       =   16
         TabIndex        =   26
         Top             =   5580
         Width           =   3495
      End
      Begin VB.TextBox txtRilascioCarta 
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
         MaxLength       =   30
         TabIndex        =   21
         Top             =   4740
         Width           =   3495
      End
      Begin VB.TextBox txtNumeroProcura 
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
         MaxLength       =   30
         TabIndex        =   18
         Top             =   3900
         Width           =   3495
      End
      Begin VB.TextBox txtFax 
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
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   17
         Top             =   3510
         Width           =   3735
      End
      Begin VB.TextBox txtCellulare 
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
         Left            =   7920
         MaxLength       =   35
         TabIndex        =   15
         Top             =   3075
         Width           =   3735
      End
      Begin VB.TextBox txtEmail 
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
         MaxLength       =   30
         TabIndex        =   16
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox txtTelefono 
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
         MaxLength       =   35
         TabIndex        =   14
         Top             =   3060
         Width           =   3495
      End
      Begin VB.TextBox txtIndirizzo 
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
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2655
         Width           =   3735
      End
      Begin VB.TextBox txtProv 
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
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   10
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtCAP 
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
         MaxLength       =   5
         TabIndex        =   9
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtProv 
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
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCAP 
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
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtCitta 
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
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtCognome 
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
         MaxLength       =   31
         TabIndex        =   0
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtNome 
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
         Left            =   7920
         MaxLength       =   31
         TabIndex        =   1
         Top             =   960
         Width           =   3735
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
         Left            =   9360
         TabIndex        =   25
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   450
         Index           =   1
         Left            =   -73200
         Picture         =   "frmPaziente.frx":07D4
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   360
         Width           =   450
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   101
         Top             =   1320
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
         Left            =   7920
         TabIndex        =   102
         Top             =   3840
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
         Left            =   7920
         TabIndex        =   103
         Top             =   4680
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Riceve"
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
         Height          =   240
         Index           =   47
         Left            =   -74880
         TabIndex        =   122
         Top             =   3000
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Regionale"
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
         Height          =   240
         Index           =   50
         Left            =   -74880
         TabIndex        =   120
         Top             =   2580
         Width           =   1890
      End
      Begin VB.Label lblTipologiaMedico 
         AutoSize        =   -1  'True
         Caption         =   "Tip.Medico"
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
         Height          =   240
         Index           =   47
         Left            =   -68520
         TabIndex        =   119
         Top             =   3000
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NAZIONE di RESIDENZA"
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
         Index           =   23
         Left            =   120
         TabIndex        =   100
         Top             =   330
         Width           =   2175
         WordWrap        =   -1  'True
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
         Left            =   5280
         TabIndex        =   97
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Accompagnatore"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   96
         Top             =   7365
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tessera San."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   48
         Left            =   6480
         TabIndex        =   95
         Top             =   6030
         Width           =   1485
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Km dal centro"
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
         Index           =   46
         Left            =   10200
         TabIndex        =   94
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gruppo San."
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   93
         Top             =   6480
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regione di Residenza"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   92
         Top             =   2250
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice ID"
         BeginProperty Font 
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
         Left            =   9720
         TabIndex        =   91
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Allergia"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   90
         Top             =   6915
         Width           =   825
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
         Index           =   31
         Left            =   120
         TabIndex        =   89
         Top             =   6915
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stato Paziente"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   88
         Top             =   6480
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Professione"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   87
         Top             =   5610
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro di Prov."
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   86
         Top             =   6030
         Width           =   1545
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
         Index           =   26
         Left            =   6480
         TabIndex        =   85
         Top             =   5190
         Width           =   1095
      End
      Begin VB.Label lblCodiceFiscale 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Fisc. - ENI - STP"
         BeginProperty Font 
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
         TabIndex        =   84
         Top             =   5610
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Distr."
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
         Left            =   4920
         TabIndex        =   83
         Top             =   5220
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ASL di residenza"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   82
         Top             =   5190
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "il"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   81
         Top             =   4800
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rilasciata/o da"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   80
         Top             =   4800
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   79
         Top             =   4365
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento di riconoscimento"
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
         Index           =   18
         Left            =   120
         TabIndex        =   78
         Top             =   4230
         Width           =   2175
         WordWrap        =   -1  'True
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
         Index           =   17
         Left            =   6480
         TabIndex        =   77
         Top             =   3960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procura Numero"
         BeginProperty Font 
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
         TabIndex        =   76
         Top             =   3900
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
         BeginProperty Font 
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
         TabIndex        =   75
         Top             =   3510
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   74
         Top             =   3510
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cellulari"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   73
         Top             =   3120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefoni"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   72
         Top             =   3100
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   71
         Top             =   2670
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.A.P."
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   70
         Top             =   2235
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prov."
         BeginProperty Font 
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
         Left            =   8760
         TabIndex        =   69
         Top             =   2235
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comune di Residenza"
         BeginProperty Font 
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
         TabIndex        =   68
         Top             =   2670
         Width           =   2280
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
         Left            =   120
         TabIndex        =   67
         Top             =   960
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
         Left            =   6480
         TabIndex        =   66
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data di Nascita"
         BeginProperty Font 
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
         TabIndex        =   65
         Top             =   1395
         Width           =   1620
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
         Left            =   4680
         TabIndex        =   64
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Città di Nascita"
         BeginProperty Font 
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
         TabIndex        =   63
         Top             =   1815
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sesso"
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   62
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prov."
         BeginProperty Font 
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
         Left            =   8760
         TabIndex        =   61
         Top             =   1815
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.A.P."
         BeginProperty Font 
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
         Left            =   6480
         TabIndex        =   60
         Top             =   1815
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Studio"
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
         Height          =   240
         Index           =   42
         Left            =   -68520
         TabIndex        =   55
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
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
         Height          =   240
         Index           =   38
         Left            =   -74880
         TabIndex        =   54
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
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
         Height          =   240
         Index           =   41
         Left            =   -74880
         TabIndex        =   53
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cellulare"
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
         Height          =   240
         Index           =   40
         Left            =   -68520
         TabIndex        =   52
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
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
         Height          =   240
         Index           =   39
         Left            =   -68520
         TabIndex        =   51
         Top             =   2160
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.A.P."
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
         Height          =   240
         Index           =   37
         Left            =   -68520
         TabIndex        =   50
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prov."
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
         Height          =   240
         Index           =   36
         Left            =   -65640
         TabIndex        =   49
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indirizzo"
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
         Height          =   240
         Index           =   35
         Left            =   -74880
         TabIndex        =   48
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Città"
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
         Height          =   240
         Index           =   34
         Left            =   -74880
         TabIndex        =   47
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Height          =   240
         Index           =   33
         Left            =   -68520
         TabIndex        =   46
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
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
         Height          =   240
         Index           =   32
         Left            =   -74880
         TabIndex        =   45
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   42
      Top             =   7800
      Width           =   11895
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
         Left            =   8880
         TabIndex        =   121
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdModuli 
         Caption         =   "&Moduli "
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
         Left            =   4440
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStampaCartella 
         Caption         =   "&Stampa Cartella Clinica"
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
         Left            =   5880
         TabIndex        =   40
         Top             =   240
         Width           =   2775
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
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuovoPaziente 
         Caption         =   "&Nuovo Paziente"
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
         Left            =   2280
         TabIndex        =   38
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPaziente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPaziente.frm
'
' <b>Descrizione</b>: Scheda Informazioni Generali associata alla tab PAZIENTI
'
' @remarks
'
' @author
'
' @date 22/02/2011 18.42
Option Explicit

'' rs della scheda
Dim rsPaziente As Recordset
Dim modifica As Boolean
Dim stoCaricando As Boolean
Dim stoPulendo As Boolean
Dim blnCaricamentoPaziente As Boolean
Dim lettera As String
Public intPazientiKey As Integer
Dim intMedicoKey As Integer
Dim blnModificato As Boolean
Dim KeyAppo As Integer
Dim NuovoPaziente As Boolean
'' indica se cancellare i turni nel caso di stato non in dialisi
Dim cancellaTurni As Boolean
'' indica se eliminare la data fine dialisi per il cambio dello stato
Dim variazioneDataFineDialisiSede As Boolean
'' indica se impostare la data inizio emodialisi in sede per il cambio della stato
Dim variazioneDataInizioDialisiSede As Boolean
'' indica se pulire le date di inizio e fine in sede quando cambia lo stato e passa in dialisi
Dim eliminaDateInizioFineDialisiSede As Boolean

Private Sub cboTipologia_Change()
    Call SetComboWidth(cboTipologia, 280)
End Sub

Private Sub Memorizza()
    Dim i As Integer
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim numKey As Integer
    
    i = 0

    If Completo Then
        If idNonValido Then
            Exit Sub
        End If
        If Not unicoCodiceFiscale Then
            Exit Sub
        End If
        Call SuperUcase(Me)
        If cboCentroProv.Text <> "" Then
            Call GestisciNuovo("CENTRI_PROVENIENZA", cboCentroProv)
        End If
        If txtKm = "" Then
            txtKm = 0
        End If
        
        
        v_Nomi = Array("COGNOME", "NOME", "CODICE_ID", "DATA_NASCITA", "SESSO", "CITTA_NASCITA", _
                    "CODICE_COMUNE_RESIDENZA", "CODICE_REGIONE", "NAZIONIID", "CAP_NASCITA", "CAP_RESIDENZA", "PROV_NASCITA", "PROV_RESIDENZA", _
                    "KM", "INDIRIZZO", "TELEFONO", "CELLULARE", "FAX", "EMAIL", _
                    "NUMERO_PROCURA", "DATA_PROCURA", "DATA_RILASCIO", "TIPO_DOCUMENTO", "CODICE_DOCUMENTO", "LUOGO_RILASCIO", _
                    "CODICE_ASL", "CODICE_FISCALE", "TESSERA_SANITARIA", "ESENZIONE_REDDITO", "ALLERGIA", "PROFESSIONE", _
                    "G_SANGUIGNO", "RH", "NOTE", "CODICE_MEDICO", "CODICE_FISCALE_CIFRATO", "TRASPORTO_IN_AMBULANZA", _
                    "CODICE_DISTRETTO", "CODICE_ESENZIONE", "CODICE_CENTRO_PROV", "CODICE_ACCOMPAGNATORE")
        v_Val = Array(txtCognome, txtNome, txtCodiceId, oData(0).data, IIf(optSesso(0).Value, "M", "F"), txtCitta, _
                    -1, -1, -1, txtCap(0), txtCap(1), txtProv(0), txtProv(1), _
                    txtKm, txtIndirizzo, txtTelefono, txtCellulare, txtFax, txtEmail, _
                    txtNumeroProcura, IIf(oData(1).data = "", Null, oData(1).data), IIf(oData(2).data = "", Null, oData(2).data), cboDocumento.ListIndex, txtNumCarta, txtRilascioCarta, _
                    -1, txtCodiceFiscale, txtTesseraSanitaria, IIf(chkEsenteReddito.Value = Checked, True, False), txtAllergia, txtProfessione, _
                    cboGSanguigno.ListIndex, GestisciOpt(optRh), txtNote, intMedicoKey, "", IIf(chkTrasportoInAmbulanza.Value = Checked, True, False), _
                    -1, -1, -1, -1)
                                        
        If cboComuneResidenza.ListIndex <> -1 Then
            v_Val(6) = cboComuneResidenza.ItemData(cboComuneResidenza.ListIndex)
        End If
        If cboRegione.ListIndex <> -1 Then
            v_Val(7) = cboRegione.ItemData(cboRegione.ListIndex)
        End If
        If cboNazione.ListIndex <> -1 Then
            v_Val(8) = cboNazione.ItemData(cboNazione.ListIndex)
        End If
        If cboAsl.ListIndex <> -1 Then
            v_Val(25) = cboAsl.ItemData(cboAsl.ListIndex)
        End If
        If cboDistretto.ListIndex <> -1 Then
            v_Val(37) = cboDistretto.ItemData(cboDistretto.ListIndex)
        End If
        If cboEsenzione.ListIndex <> -1 Then
            v_Val(38) = cboEsenzione.ItemData(cboEsenzione.ListIndex)
        End If
        If cboCentroProv.ListIndex <> -1 Then
            v_Val(39) = cboCentroProv.ItemData(cboCentroProv.ListIndex)
        End If
        If cboAccompagnatore.ListIndex <> -1 Then
            v_Val(40) = cboAccompagnatore.ItemData(cboAccompagnatore.ListIndex)
        End If
        
        Set rsPaziente = New Recordset
        
        If modifica Then
        
            rsPaziente.Open "SELECT * FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            KeyAppo = intPazientiKey
            Do While i <> UBound(v_Nomi) + 1
                rsPaziente(v_Nomi(i)) = v_Val(i)
                i = i + 1
            Loop
            rsPaziente.Update
            
        Else
        
            rsPaziente.Open "PAZIENTI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsPaziente.AddNew
            KeyAppo = GetNumero("PAZIENTI")
            rsPaziente("KEY") = KeyAppo
            Do While i <> UBound(v_Nomi) + 1
                rsPaziente(v_Nomi(i)) = v_Val(i)
                i = i + 1
            Loop
            rsPaziente.Update
            
        End If
        rsPaziente.Close
        
        Call GestisciStato(KeyAppo)
        If cancellaTurni Then
            Call EliminaTurni
        End If
        If eliminaDateInizioFineDialisiSede Then
            Call EliminaDateDialisiSede
        End If
        If variazioneDataInizioDialisiSede Then
            Call CambiaDataInizioDialisiSede
        End If
        If variazioneDataFineDialisiSede Then
            Call CambiaDataFineDialisiSede
        End If
        
        
        If TRACCIATO Then
            ' salva le modifiche anche in tracciatura
            v_Nomi = Array("KEY", "NOME", "COGNOME", "CODICE_FISCALE")
            v_Val = Array(KeyAppo, txtNome, txtCognome, txtCodiceFiscale)
            If modifica Then
                rsPaziente.Open "SELECT * FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnTrac, adOpenKeyset, adLockPessimistic, adCmdText
                rsPaziente.Update v_Nomi, v_Val
            Else
                rsPaziente.Open "PAZIENTI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
                rsPaziente.AddNew v_Nomi, v_Val
                rsPaziente.Update
            End If
        End If
        
        Set rsPaziente = Nothing
        
        modifica = True
        NuovoPaziente = False
        MsgBox "I dati sono stati memorizzzati nell'archivio", vbInformation, "Informazioni"
        
        blnModificato = False
    End If
End Sub

Private Sub cmdMemorizza_Click()
    Dim rsMedico As Recordset
    Dim s As Integer
    Dim v_ValMedico() As Variant
    Dim v_NomiMedico() As Variant
    Dim numKey As Integer
           
    If intMedicoKey = 0 And txtCognomeMedico.Text = "" Then     'nel caso in cui voglio memorizzare solo il paziente senza medico associato
        If intPazientiKey = 0 Then
            modifica = False
            Call Memorizza
            Exit Sub
        Else
            modifica = True
            Call Memorizza
            Exit Sub
        End If
    End If
        
    Call SuperUcase(Me)
    
    If intMedicoKey > 0 Then
        modifica = True
    End If
    
    If modifica Then            ' controllo per l' inserimento o modifica del medico di base
        numKey = intMedicoKey
    Else
        numKey = GetNumero("MEDICI_BASE")
    End If
    
    s = 0
                    
        v_NomiMedico = Array("KEY", "COGNOME", "NOME", "COMUNE", "INDIRIZZO", "CAP", "PROV", "TELEFONO", "STUDIO" _
                    , "CELLULARE", "FAX", "EMAIL", "CODICE", "PRESENZA_BARCODE", "CODICE_TIPO_MEDICO", "RICEVE")
        v_ValMedico = Array(numKey, txtCognomeMedico, txtNomeMedico, txtCittaMedico, txtIndirizzoMedico, txtCapMedico, txtProvMedico, txtTelefonoMedico, txtStudioMedico _
                    , txtCellulareMedico, txtFaxMedico, txtEmailMedico, txtCodiceRegionaleMedico, IIf(chkPresenzaBarCode.Value = Checked, True, False), -1, txtRiceve)
                    
        If cboTipologia.ListIndex <> -1 Then
            v_ValMedico(14) = cboTipologia.ItemData(cboTipologia.ListIndex)
        End If
                    
        Set rsMedico = New Recordset
        
        If modifica Then
            
            rsMedico.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMedicoKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            Do While s <> UBound(v_NomiMedico) + 1
                rsMedico(v_NomiMedico(s)) = v_ValMedico(s)
                s = s + 1
            Loop
            rsMedico.Update
            intMedicoKey = numKey
            If intPazientiKey = 0 Then
                intPazientiKey = KeyAppo
            End If
            If intMedicoKey > 0 And intPazientiKey = 0 Then
                modifica = False
                Call Memorizza
            Else
                modifica = True
                Call Memorizza
            End If
        
        Else
            
            rsMedico.Open "MEDICI_BASE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsMedico.AddNew
            Do While s <> UBound(v_NomiMedico) + 1
                rsMedico(v_NomiMedico(s)) = v_ValMedico(s)
                s = s + 1
            Loop
            rsMedico.Update
            intMedicoKey = numKey
            If intPazientiKey = 0 Then
                intPazientiKey = KeyAppo
            End If
            If intPazientiKey = 0 Then      'INSERIMENTO NUOVO MEDICO: controlla se inserisco o modifico un paziente
                modifica = False
                Call Memorizza
            Else
                modifica = True
                Call Memorizza
            End If
            
        End If
        rsMedico.Clone
        Set rsMedico = Nothing
        blnModificato = False

End Sub

Private Sub Form_Activate()
    Dim blnModificatoAppo As Boolean
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    blnModificatoAppo = blnModificato
    Call RicaricaComboBox("CENTRI_PROVENIENZA", "NOME", cboCentroProv)
    Call RicaricaComboBox("ASL", "NOME", cboAsl)
    Call RicaricaComboBox("TIPOLOGIE_ESENZIONE", "CODICE", cboEsenzione)
    If cboEsenzione.ListIndex = -1 Then cboEsenzione.ListIndex = 0
    Call RicaricaComboBox("COMUNI", "NOME", cboComuneResidenza)
    Call RicaricaComboBox("SELECT KEY, (COGNOME + ' '+ NOME) AS CAMPO FROM ACCOMPAGNATORI", "CAMPO", cboAccompagnatore)
    Call RicaricaComboBox("NAZIONI", "NOME", cboNazione)
    If cboNazione.ListIndex = -1 Then cboNazione.ListIndex = GetIndex(cboNazione, "Italia")
    If cboStato.ListIndex = -1 Then cboStato.ListIndex = 0
    Call RicaricaComboBox("REGIONI", "NOME", cboRegione)
    If cboRegione.ListIndex = -1 Then cboRegione.ListIndex = GetIndex(cboRegione, "Campania")
    blnModificato = blnModificatoAppo
    Call RicaricaComboBox("TIPOLOGIE_MEDICO", "NOME", cboTipologia)
    Select Case CaricaPazienteInAperturaForm(Me.Caption, blnModificato, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
             txtCognome.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
    
    If structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Or structIntestazione.sCodiceSTS = CODICESTS_EM_IRPINA Then
        lblTipologiaMedico(47).Visible = True
        cboTipologia.Visible = True
    End If
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    For i = 0 To 2
        oData(i).ConnectionString = strConnectionStringCentro
    Next i
    picMostraStato.Picture = LoadResPicture("pin1", 0)
    Call AnnullaVarStato
    modifica = False
    cancellaTurni = False
    variazioneDataFineDialisiSede = False
    variazioneDataInizioDialisiSede = False
    eliminaDateInizioFineDialisiSede = False
    tabScheda.Tab = 0
    Call ProponiId
    Call RicaricaComboBox("TIPO_STATO", "NOME", cboStato)
    blnModificato = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        intMedicoKey = 0
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click (0)
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Carica il primo id libero se il paziente non ha id
Private Sub ProponiId()
    Set rsPaziente = New Recordset
    rsPaziente.Open "SELECT MAX(CODICE_ID) AS MASSIMO FROM PAZIENTI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsPaziente("MASSIMO")) Then
        txtCodiceId = rsPaziente("MASSIMO") + 1
    Else
        txtCodiceId = 1
    End If
    Set rsPaziente = Nothing
End Sub

'' Gestisce il valore numerico degli opt
'
' @param opt vettore di opt da analizzare
' @return indice dell'opt selezionato o 2 se nessuno dei due è selezionato
Private Function GestisciOpt(ByRef opt As Object) As Byte
    If opt(0).Value = False And opt(1).Value = False Then
        GestisciOpt = 2
    Else
        GestisciOpt = IIf(opt(0).Value = True, 0, 1)
    End If
End Function

'' Carica i dati se dell'ospite
Private Sub CaricaStatoOspite()
    Dim rsDataset As New Recordset
    Dim i As Integer
    
    rsDataset.Open "SELECT * FROM PAZIENTI_OSPITI WHERE CODICE_PAZIENTE=" & intPazientiKey & " ORDER BY DATA_ARRIVO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    i = 0
    Do While (Not rsDataset.EOF) And i <= 3
        i = i + 1
        statoPaziente.dataArrivi(i) = rsDataset("DATA_ARRIVO")
        statoPaziente.dataPartenza(i) = rsDataset("DATA_PARTENZA")
        statoPaziente.centriProv(i) = rsDataset("CENTRO_PROVENIENZA")
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

'' Verifica se l'id inserito sia valido
Private Function idNonValido() As Boolean
    Dim massimo As Integer
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT MAX(CODICE_ID) AS MASSIMO FROM PAZIENTI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If IsNull(rsDataset("MASSIMO")) Then
        massimo = 0
    Else
        massimo = rsDataset("MASSIMO")
    End If
    rsDataset.Close
    
    If txtCodiceId = "" Then
        txtCodiceId = massimo + 1
        idNonValido = False
    Else
        rsDataset.Open "SELECT KEY,CODICE_ID FROM PAZIENTI WHERE CODICE_ID=" & txtCodiceId, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("KEY") <> intPazientiKey Then
                If MsgBox("Codice ID già in uso." & vbCrLf & "Si preferisce assegnare il valore " & massimo + 1 & " scelto dal sistema?", vbCritical + vbYesNo, "Attenzione") = vbYes Then
                    txtCodiceId = massimo + 1
                    idNonValido = False
                Else
                    idNonValido = True
                End If
            Else
                idNonValido = False
            End If
        Else
            idNonValido = False
        End If
    End If
    Set rsDataset = Nothing
End Function

'' Elimina le date di inizio e fine dialisi in sede
Private Sub EliminaDateDialisiSede()
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT * FROM ANAMNESI_NEFROLOGICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        rsDataset("DATA_INIZIO") = Null
        rsDataset("DATA_FINE") = Null
        rsDataset.Update
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

Private Sub CreaSchedaNefrologica(campo As String, data As Date)
    Dim rsDataset As New Recordset
    
    rsDataset.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew
    rsDataset("KEY") = GetNumero("ANAMNESI_NEFROLOGICHE")
    rsDataset("CODICE_PAZIENTE") = intPazientiKey
    rsDataset(campo) = data
    rsDataset.Update
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

'' Cambia la data di inizio dialisi in sede se lo stato passa a ospite
Private Sub CambiaDataInizioDialisiSede()
    Dim rsDataset As New Recordset
    Dim data As Date

    
    data = CDate(IIf(statoPaziente.dataArrivi(1) = "", CDate("01/01/1900"), statoPaziente.dataArrivi(1)))
    If data < CDate(IIf(statoPaziente.dataArrivi(2) = "", CDate("01/01/1900"), statoPaziente.dataArrivi(2))) Then
        data = CDate(statoPaziente.dataArrivi(2))
    End If
    If data < CDate(IIf(statoPaziente.dataArrivi(3) = "", CDate("01/01/1900"), statoPaziente.dataArrivi(3))) Then
        data = CDate(statoPaziente.dataArrivi(3))
    End If

    
    If MsgBox("Indicare il " & data & " come INIZIO EMODIALISI IN SEDE?", vbQuestion + vbYesNo, "Cambio stato paziente") = vbYes Then
        rsDataset.Open "SELECT * FROM ANAMNESI_NEFROLOGICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If IsNull(rsDataset("DATA_INIZIO")) Then
                rsDataset("DATA_INIZIO") = data
                rsDataset.Update
                MsgBox "Data inizio emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
            Else
                If MsgBox("E' già presente la data del " & rsDataset("DATA_INIZIO") & " come INIZIO EMODIALISI IN SEDE. SOSTITUIRLA?", vbQuestion + vbYesNo, "Cambio stato paziente") = vbYes Then
                    rsDataset("DATA_INIZIO") = data
                    rsDataset.Update
                    MsgBox "Data inizio emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
                End If
            End If
        Else
            Call CreaSchedaNefrologica("DATA_INIZIO", data)
            MsgBox "Data inizio emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
        End If
    End If
    
    Set rsDataset = Nothing
End Sub

'' Cambia la data di fine dialisi in sede se lo stato passa a deceduto o trapiantato o trasferito o ospite
Private Sub CambiaDataFineDialisiSede()
    Dim rsDataset As New Recordset
    Dim data As Date

    If statoPaziente.statoPaz = TPDECEDUTO Or statoPaziente.statoPaz = TPTRAPIANTO Or statoPaziente.statoPaz = TPTRASFERITO Then
        data = CDate(statoPaziente.dataStato)
    ElseIf statoPaziente.statoPaz = TPOSPITE Then
        data = CDate(IIf(statoPaziente.dataPartenza(1) = "", CDate("01/01/1900"), statoPaziente.dataPartenza(1)))
        If data < CDate(IIf(statoPaziente.dataPartenza(2) = "", CDate("01/01/1900"), statoPaziente.dataPartenza(2))) Then
            data = CDate(statoPaziente.dataPartenza(2))
        End If
        If data < CDate(IIf(statoPaziente.dataPartenza(3) = "", CDate("01/01/1900"), statoPaziente.dataPartenza(3))) Then
            data = CDate(statoPaziente.dataPartenza(3))
        End If
    End If
    
    If MsgBox("Indicare il " & data & " come FINE EMODIALISI IN SEDE?", vbQuestion + vbYesNo, "Cambio stato paziente") = vbYes Then
        rsDataset.Open "SELECT * FROM ANAMNESI_NEFROLOGICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If IsNull(rsDataset("DATA_FINE")) Then
                rsDataset("DATA_FINE") = data
                rsDataset.Update
                MsgBox "Data fine emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
            Else
                If MsgBox("E' già presente la data del " & rsDataset("DATA_FINE") & " come FINE EMODIALISI IN SEDE. SOSTITUIRLA?", vbQuestion + vbYesNo, "Cambio stato paziente") = vbYes Then
                    rsDataset("DATA_FINE") = data
                    rsDataset.Update
                    MsgBox "Data fine emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
                End If
            End If
        Else
            Call CreaSchedaNefrologica("DATA_FINE", data)
            MsgBox "Data inizio emodialisi in sede aggiornata", vbInformation, "Cambio stato paziente"
        End If
    End If
    
    Set rsDataset = Nothing
End Sub

'' Verifica che tutti i dati necessati sia inseriti correttamente prima di memorizzare
Private Function Completo() As Boolean
    Completo = False
    Dim nome As String
    If txtCognome = "" Then
        nome = "COGNOME"
    ElseIf txtNome = "" Then
        nome = "NOME"
    ElseIf oData(0).data = "" Then
        nome = "DATA DI NASCITA"
    ElseIf cboComuneResidenza.ListIndex = -1 And UCase(cboNazione.Text) = UCase("ITALIA") Then
        nome = "CITTA RESIDENZA"
    ElseIf cboRegione.ListIndex = -1 And UCase(cboNazione.Text) = UCase("ITALIA") Then
        nome = "REGIONE RESIDENZA"
    ElseIf cboNazione.ListIndex = -1 Then
        nome = "NAZIONE RESIDENZA"
    ElseIf cboAsl.ListIndex = -1 And UCase(cboNazione.Text) = UCase("ITALIA") Then
        nome = "A.S.L."
    ElseIf txtCodiceFiscale = "" Then
        If UCase(cboNazione.Text) = UCase("Italia") Then
            nome = "CODICE FISCALE"
        Else
            nome = "CODICE ENI o STP"
        End If
    ElseIf optSesso(0).Value = False And optSesso(1).Value = False Then
        nome = "SESSO"
    ElseIf cboRegione.ListIndex = 3 Then
        ' solo quelli della campania devono avere il distretto obbligatorio
        If cboDistretto.ListIndex <> -1 Then
            Completo = True
            Exit Function
        Else
            nome = "DISTRETTO"
        End If
    Else
        Completo = True
        Exit Function
    End If
    MsgBox "Inserire i dati obbligatori" & vbCrLf & "Campo: " & nome, vbCritical, "Attenzione"
End Function

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    stoPulendo = True
    blnCaricamentoPaziente = False
    modifica = False
    cancellaTurni = False
    variazioneDataFineDialisiSede = False
    variazioneDataInizioDialisiSede = False
    eliminaDateInizioFineDialisiSede = False
    Call PulisciForm(Me)
    intPazientiKey = 0
    intMedicoKey = 0
    KeyAppo = 0
    chkPresenzaBarCode.Value = False
    lblEta = ""
    For i = 0 To 2
        oData(i).Pulisci
    Next i
    For i = 0 To 1
        optSesso(i).ForeColor = vbBlack
        optRh(i).ForeColor = vbBlack
        optRh(i).Value = False
        optSesso(i).Value = False
    Next i
    Call AnnullaVarStato
    cboNazione.ListIndex = GetIndex(cboNazione, "Italia")
    cboRegione.ListIndex = GetIndex(cboRegione, "Campania")
    cboEsenzione.ListIndex = GetIndex(cboEsenzione, "NESSUNA")
    cboStato.ListIndex = 0
    chkEsenteReddito.Value = Unchecked
    chkTrasportoInAmbulanza.Value = Unchecked
    txtCognome.SetFocus
    stoPulendo = False
    Call ProponiId
    Call DisabilitaMedico
    blnModificato = False
End Sub

'' Limita l'inserimento a 30 caratteri
Private Sub cboCentroProv_KeyPress(KeyAscii As Integer)
    If Len(cboCentroProv.Text) >= 30 Then
        Beep
        KeyAscii = 0
    End If
End Sub

'' Elimina i turni dei pazienti deceduti, trasferiti o trapiantato
Private Sub EliminaTurni()
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        rsDataset.Delete
        rsDataset.Update
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Function ControlloCodiceFiscale(codice As String) As Boolean
    Dim vett(1 To 90, 1 To 2) As Integer
    Call CaricaTabella(vett)
    Dim somma As Integer
    Dim resto As Integer
    Dim i As Integer
    For i = 1 To 15 Step 2
        somma = somma + vett(Asc(CStr(Mid(codice, i, 1))), 2)
    Next i
    For i = 2 To 14 Step 2
        somma = somma + vett(Asc(CStr(Mid(codice, i, 1))), 1)
    Next i
    resto = somma Mod 26
    If CStr((Mid(codice, 16, 1))) = Chr(resto + 65) Then
        ControlloCodiceFiscale = True
    Else
        ControlloCodiceFiscale = False
    End If
End Function

Private Sub CaricaTabella(v_tabella() As Integer)
    Dim i As Integer
    Dim k As Integer
    k = 2
    For i = Asc("0") To Asc("9")
        v_tabella(i, 1) = CInt(Chr(i))
        Select Case CInt(Chr(i))
            Case 0: v_tabella(i, 2) = 1
            Case 1: v_tabella(i, 2) = 0
            Case 2: v_tabella(i, 2) = 5
            Case 5
                v_tabella(i, 2) = 13
                k = k + 4
            Case Else
                v_tabella(i, 2) = 5 + k
                k = k + 2
        End Select
    Next i
    Dim count As Integer
    k = 2
    count = 0
    For i = Asc("A") To Asc("Z")
        v_tabella(i, 1) = count
        count = count + 1
        Select Case (Chr(i))
            Case "A": v_tabella(i, 2) = 1
            Case "B": v_tabella(i, 2) = 0
            Case "C": v_tabella(i, 2) = 5
            Case "F"
                v_tabella(i, 2) = 13
                k = k + 4
            Case Else
                v_tabella(i, 2) = 5 + k
                k = k + 2
        End Select
    Next i
    v_tabella(Asc("K"), 2) = 2
    v_tabella(Asc("L"), 2) = 4
    v_tabella(Asc("M"), 2) = 18
    v_tabella(Asc("N"), 2) = 20
    v_tabella(Asc("O"), 2) = 11
    v_tabella(Asc("P"), 2) = 3
    v_tabella(Asc("Q"), 2) = 6
    v_tabella(Asc("R"), 2) = 8
    v_tabella(Asc("S"), 2) = 12
    v_tabella(Asc("T"), 2) = 14
    v_tabella(Asc("U"), 2) = 16
    v_tabella(Asc("V"), 2) = 10
    v_tabella(Asc("W"), 2) = 22
    v_tabella(Asc("X"), 2) = 25
    v_tabella(Asc("Y"), 2) = 24
    v_tabella(Asc("Z"), 2) = 23
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdNuovoPaziente_Click()
    NuovoPaziente = True
    Call PulisciTutto
    oPazientiKey.OnClosingForm (Me.Caption)
End Sub

Private Sub cmdModuli_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
    Else
        Unload frmModuliWord
        Load frmModuliWord
        frmModuliWord.LetCodicePaziente = intPazientiKey
        frmModuliWord.Show 1
    End If
End Sub

Private Sub cboAsl_Click()
    If stoPulendo Then Exit Sub
    If cboAsl.ListIndex = -1 Then Exit Sub
    Call RicaricaComboBox("SELECT * FROM DISTRETTI WHERE CODICE_ASL=" & cboAsl.ItemData(cboAsl.ListIndex), "NOME", cboDistretto)
    blnModificato = True
End Sub

Private Sub cboAsl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 And Shift Then
        MsgBox strConnectionStringCentro & " " & strConnectionStringTracciatura
    End If
End Sub

Private Sub cboStato_Click()
    On Error GoTo gestione
    
    If stoCaricando Or cboStato.ListIndex = -1 Then Exit Sub
    statoPaziente.statoPaz = cboStato.ListIndex
    If cboStato.ListIndex = 0 Or cboStato.ListIndex = 5 Then Exit Sub
    If statoPaziente.statoPaz = TPOSPITE Then
        Call CaricaStatoOspite
    End If
        
    frmStatoPaz.Show 1
    
    blnModificato = True
    Exit Sub
gestione:
    If Err.Number = 401 Then
        Exit Sub
    Else
        MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
    End If
End Sub

Private Sub cboNazione_Click()
    If UCase(cboNazione.Text) = UCase("Italia") Then
        Label1(25).Enabled = True
        cboRegione.Enabled = True
        Label1(8).Enabled = True
        cboComuneResidenza.Enabled = True
        Label1(22).Enabled = True
        cboAsl.Enabled = True
        Label1(24).Enabled = True
        cboDistretto.Enabled = True
        cboRegione.ListIndex = GetIndex(cboRegione, "Campania")
        If Not stoCaricando Then blnModificato = True
    Else
        Dim blnProsegui As Boolean
        
        If stoPulendo Or blnCaricamentoPaziente Or intPazientiKey = 0 Then
            blnProsegui = True
        Else
            If cboRegione.Text <> "" Or cboAsl.Text <> "" Or cboDistretto.Text <> "" Or cboComuneResidenza.Text <> "" Then
                If MsgBox("Sostituendo la NAZIONE di PROVENIENZA verranno cancellati i dati" & vbCrLf & "relativi alla CITTA' - REGIONE - ASL - DISTRETTO. SI CONFERMA?", vbQuestion + vbYesNo + vbDefaultButton2, "ATTENZIONE!!!") = vbYes Then
                    blnProsegui = True
                    blnModificato = True
                Else
                    blnProsegui = False
                    blnModificato = False
                End If
            Else
                blnProsegui = True
            End If

        End If
            
        If blnProsegui Then
            Label1(25).Enabled = False
            cboRegione.Enabled = False
            Label1(8).Enabled = False
            cboComuneResidenza.Enabled = False
            Label1(22).Enabled = False
            cboAsl.Enabled = False
            Label1(24).Enabled = False
            cboDistretto.Enabled = False
            cboRegione.ListIndex = -1
            cboComuneResidenza.ListIndex = -1
            cboAsl.ListIndex = -1
            cboDistretto.ListIndex = -1
            blnModificato = True
        Else
            stoCaricando = True
            cboNazione.ListIndex = GetIndex(cboNazione, "italia")
            stoCaricando = False
        End If
    End If
End Sub

'' Verifica che il codice fiscale inserito sia unico nel db
Private Function unicoCodiceFiscale() As Boolean
    Dim rsDataset As New Recordset
    
    unicoCodiceFiscale = True
    rsDataset.Open "SELECT CODICE_FISCALE, key FROM PAZIENTI WHERE NOT KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        If UCase(txtCodiceFiscale) = UCase(rsDataset("CODICE_FISCALE")) Then
            MsgBox "IMPOSSIBILE MEMORIZZARE!!!" & vbCrLf & "Paziente già presente in archivio", vbCritical, "Attenzione"
            unicoCodiceFiscale = False
            Exit Do
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
End Function

'' Gestisce le informazioni sullo stato del paziente (data, stato, donatore, centri)
Private Sub GestisciStato(codicePaziente As Integer)
    On Error Resume Next
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    Dim i As Integer
    Dim numDate As Integer
    Dim cmCommand As New Command
    
    numDate = 0
    For i = 1 To 3
        If IsDate(statoPaziente.dataArrivi(i)) Then
            numDate = numDate + 1
        End If
    Next i
                
    laData = DateValue(Format(numDate + 1, "00") & "/" & Format(11 + IIf(numDate < 0, numDate, 0), "00") & "/" & Year(Now))
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & codicePaziente, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If rsDataset("STATO") <> statoPaziente.statoPaz Then
        Select Case statoPaziente.statoPaz
            Case 0
                eliminaDateInizioFineDialisiSede = True
            Case 1, 2, 3
                cancellaTurni = True
                If IsDate(statoPaziente.dataStato) Then
                    variazioneDataFineDialisiSede = True
                End If
            Case 4
                If IsDate(statoPaziente.dataArrivi(1)) Or IsDate(statoPaziente.dataArrivi(2)) Or IsDate(statoPaziente.dataArrivi(3)) Then
                    variazioneDataInizioDialisiSede = True
                End If
                If IsDate(statoPaziente.dataPartenza(1)) Or IsDate(statoPaziente.dataPartenza(2)) Or IsDate(statoPaziente.dataPartenza(3)) Then
                    variazioneDataFineDialisiSede = True
                End If
        End Select
    Else
        Select Case statoPaziente.statoPaz
            Case 1, 2, 3
                If IsDate(statoPaziente.dataStato) Then
                    If DateValue(statoPaziente.dataStato) <> rsDataset("STATODATA") Then
                        variazioneDataFineDialisiSede = True
                    End If
                End If
            Case 4
                rsAppo.Open "SELECT * FROM PAZIENTI_OSPITI WHERE CODICE_PAZIENTE=" & codicePaziente & " ORDER BY DATA_ARRIVO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If rsAppo.RecordCount <> numDate Then
                    variazioneDataInizioDialisiSede = True
                    variazioneDataFineDialisiSede = True
                Else
                    variazioneDataFineDialisiSede = False
                    variazioneDataInizioDialisiSede = False
                    i = 1
                    Do While Not rsAppo.EOF And i <= numDate
                        If DateValue(statoPaziente.dataArrivi(i)) <> rsAppo("DATA_ARRIVO") Then
                            variazioneDataInizioDialisiSede = True
                        End If
                        If DateValue(statoPaziente.dataPartenza(i)) <> rsAppo("DATA_PARTENZA") Then
                            variazioneDataFineDialisiSede = True
                        End If
                        i = i + 1
                        rsAppo.MoveNext
                    Loop
                End If
                rsAppo.Close
        End Select
    End If
    rsDataset.Close
    
    'If date > laData Then          ERRORE DEI TURNI ASSEGNATI
    '    cmCommand.ActiveConnection = cnPrinc
    '    cmCommand.CommandType = adCmdText
    '    cmCommand.CommandText = "DELETE * FROM turni WHERE CODICE_PAZIENTE=" & codicePaziente
    '    cmCommand.Execute
    'End If
    
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & codicePaziente, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsDataset("STATO") = statoPaziente.statoPaz
    rsDataset("STATODATA") = IIf(statoPaziente.dataStato = "", Null, statoPaziente.dataStato)
    rsDataset("STATODONATORE") = statoPaziente.donatore
    rsDataset.Update
    rsDataset.Close
    
    If Not statoPaziente.statoPaz = tpDIALISI Then
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
        cmCommand.CommandText = "DELETE * FROM PAZIENTI_OSPITI WHERE CODICE_PAZIENTE=" & codicePaziente
        cmCommand.Execute
    End If
    
    If statoPaziente.statoPaz = TPOSPITE Then
        rsDataset.Open "PAZIENTI_OSPITI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        For i = 1 To numDate
            rsDataset.AddNew
            rsDataset("KEY") = GetNumero("PAZIENTI_OSPITI")
            rsDataset("CODICE_PAZIENTE") = codicePaziente
            rsDataset("DATA_ARRIVO") = statoPaziente.dataArrivi(i)
            rsDataset("DATA_PARTENZA") = statoPaziente.dataPartenza(i)
            rsDataset("CENTRO_PROVENIENZA") = statoPaziente.centriProv(i)
            rsDataset.Update
        Next i
        rsDataset.Close
    End If
End Sub

Private Sub cmdStampaCartella_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
    Else
        structIntestazione.sPaziente = txtCognome & " " & txtNome
        structIntestazione.sDataPaziente = oData(0).data
    
        Unload frmStampaCartellaClinica
        Load frmStampaCartellaClinica
        frmStampaCartellaClinica.LetCodiceId = txtCodiceId
        frmStampaCartellaClinica.LetCodicePaziente = intPazientiKey
        frmStampaCartellaClinica.Show 1
    End If
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    tTrova.Tipo = IIf(Index = 0, tpPAZIENTE, tpMEDICOBASE)
    tTrova.condizione = ""
    tTrova.condStato = ""
    tTrova.isOpenFromInfoGenerali = True
    frmTrova.Show 1
    If Index = 0 Then
        Select Case tTrova.keyReturn
            Case -1
                ' nuovo paziente
                Call PulisciTutto
                NuovoPaziente = True
            Case 0
                ' indietro
                Unload Me
            Case Else
                Call PulisciTutto
                intPazientiKey = tTrova.keyReturn
                Call CaricaPaziente
        End Select
    Else
        intMedicoKey = tTrova.keyReturn
        Call CaricaMedico
    End If
    tTrova.isOpenFromInfoGenerali = False
End Sub

'' Carica i dati del medico
Private Sub CaricaMedico()
    Dim rsDataset As Recordset
    
    If intMedicoKey = 0 And txtCognomeMedico.Text = "" Then     ' controllo per vedere se c'è il medico e il cognome del medico
        Call DisabilitaMedico                                   ' in tal caso disabilito tutto
        Exit Sub
    ElseIf intMedicoKey = 0 Then                                ' nel caso in cui il medico è presente, oppure è presente ma non lo carico
        Exit Sub                                                ' esce dalla sub
    End If
    
    If intMedicoKey = -1 Then
        modifica = False
        intMedicoKey = 0
        txtCognomeMedico = ""
        txtNomeMedico = ""
        txtCittaMedico = ""
        txtCapMedico = ""
        txtProvMedico = ""
        txtTelefonoMedico = ""
        txtCellulareMedico = ""
        txtFaxMedico = ""
        txtEmailMedico = ""
        txtIndirizzoMedico = ""
        txtStudioMedico = ""
        txtCodiceRegionaleMedico = ""
        chkPresenzaBarCode.Value = False
        txtRiceve = ""
        cboTipologia.ListIndex = 4
        Call AbilitaMedico
        txtCognomeMedico.SetFocus
    Exit Sub
    End If
    
    Set rsDataset = New Recordset
    
    rsDataset.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMedicoKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    txtCognomeMedico = rsDataset("COGNOME") & ""
    txtNomeMedico = rsDataset("NOME") & ""
    txtCittaMedico = rsDataset("COMUNE") & ""
    txtCapMedico = rsDataset("CAP") & ""
    txtProvMedico = rsDataset("PROV") & ""
    txtTelefonoMedico = rsDataset("TELEFONO") & ""
    txtCellulareMedico = rsDataset("CELLULARE") & ""
    txtFaxMedico = rsDataset("FAX") & ""
    txtEmailMedico = rsDataset("EMAIL") & ""
    txtIndirizzoMedico = rsDataset("INDIRIZZO") & ""
    txtStudioMedico = rsDataset("STUDIO") & ""
    txtCodiceRegionaleMedico = rsDataset("CODICE") & ""
    chkPresenzaBarCode.Value = IIf(CBool(rsDataset("PRESENZA_BARCODE")), Checked, Unchecked)
    txtRiceve = rsDataset("RICEVE") & ""
    cboTipologia.ListIndex = GetCboListIndex(rsDataset("CODICE_TIPO_MEDICO"), cboTipologia)
    Call AbilitaMedico
    
    Set rsDataset = Nothing
    
    blnModificato = False
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
        
    If intPazientiKey = 0 Then Exit Sub

    blnCaricamentoPaziente = True
    Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
    cancellaTurni = False
    variazioneDataFineDialisiSede = False
    variazioneDataInizioDialisiSede = False
    eliminaDateInizioFineDialisiSede = False
    ' discolora tutti per precauzione
    Call ColoraSel(optSesso, 0, 2, vbBlack, vbBlack)
    Call ColoraSel(optRh, 0, 2, vbBlack, vbBlack)
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    txtCognome = rsDataset("COGNOME")
    txtNome = rsDataset("NOME")
    txtCodiceId = rsDataset("CODICE_ID") & ""
    txtCap(0) = rsDataset("CAP_NASCITA")
    txtCap(1) = rsDataset("CAP_RESIDENZA")
    txtCellulare = rsDataset("CELLULARE")
    txtCitta = rsDataset("CITTA_NASCITA")
    cboNazione.ListIndex = GetCboListIndex(rsDataset("NAZIONIID"), cboNazione)
    If cboNazione.ListIndex = -1 Then cboNazione.ListIndex = GetIndex(cboNazione, "Italia")
    cboRegione.ListIndex = GetCboListIndex(rsDataset("CODICE_REGIONE"), cboRegione)
    cboComuneResidenza.ListIndex = GetCboListIndex(rsDataset("CODICE_COMUNE_RESIDENZA"), cboComuneResidenza)
    cboAccompagnatore.ListIndex = GetCboListIndex(rsDataset("CODICE_ACCOMPAGNATORE"), cboAccompagnatore)
    txtCodiceFiscale = rsDataset("CODICE_FISCALE")
    txtTesseraSanitaria = rsDataset("TESSERA_SANITARIA")
    txtEmail = rsDataset("EMAIL")
    chkEsenteReddito.Value = IIf(CBool(rsDataset("ESENZIONE_REDDITO")), Checked, Unchecked)
    cboEsenzione.ListIndex = GetCboListIndex(rsDataset("CODICE_ESENZIONE"), cboEsenzione)
    txtFax = rsDataset("FAX")
    txtIndirizzo = rsDataset("INDIRIZZO")
    txtNote = rsDataset("NOTE")
    txtNumCarta = rsDataset("CODICE_DOCUMENTO")
    txtNumeroProcura = rsDataset("NUMERO_PROCURA")
    txtProfessione = rsDataset("PROFESSIONE")
    txtProv(0) = rsDataset("PROV_NASCITA")
    txtProv(1) = rsDataset("PROV_RESIDENZA")
    txtKm = rsDataset("KM")
    cboAsl.ListIndex = GetCboListIndex(rsDataset("CODICE_ASL"), cboAsl)
    cboDistretto.ListIndex = GetCboListIndex(IIf(IsNull(rsDataset("CODICE_DISTRETTO")), 0, rsDataset("CODICE_DISTRETTO")), cboDistretto)
    txtRilascioCarta = rsDataset("LUOGO_RILASCIO")
    txtTelefono = rsDataset("TELEFONO")
    chkTrasportoInAmbulanza.Value = IIf(CBool(rsDataset("TRASPORTO_IN_AMBULANZA")), Checked, Unchecked)
    If rsDataset("SESSO") = "" Then
        optSesso(0).Value = False
        optSesso(1).Value = False
    ElseIf rsDataset("SESSO") = "M" Then
        optSesso(0).Value = True
    Else
        optSesso(1).Value = True
    End If
    If rsDataset("RH") <> 2 Then
        optRh(rsDataset("RH")).Value = True
    Else
        optRh(0).Value = False
        optRh(1).Value = False
    End If
    txtAllergia = rsDataset("ALLERGIA")
    cboDocumento.ListIndex = rsDataset("TIPO_DOCUMENTO")
    cboGSanguigno.ListIndex = rsDataset("G_SANGUIGNO")
    stoCaricando = True
    cboStato.ListIndex = GetCboListIndex(rsDataset("STATO"), cboStato)
    stoCaricando = False
    cboCentroProv.ListIndex = GetCboListIndex(rsDataset("CODICE_CENTRO_PROV"), cboCentroProv)
    oData(0).data = rsDataset("DATA_NASCITA")
    ' carica i riferimenti del medico del paziente attivando l'evento change di intMedicoKey
    intMedicoKey = rsDataset("CODICE_MEDICO")
    Call CaricaMedico
    If rsDataset("DATA_PROCURA") <> "" Then
        oData(1).data = rsDataset("DATA_PROCURA")
    End If
    If rsDataset("DATA_RILASCIO") <> "" Then
        oData(2).data = rsDataset("DATA_RILASCIO")
    End If
    
    Dim somma As Integer
    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
    
    ' carica anche le var e i vettori dello stato
    If cboStato.ListIndex <> 0 And cboStato.ListIndex <> -1 Then
        statoPaziente.dataStato = rsDataset("STATODATA") & ""
        statoPaziente.statoPaz = rsDataset("STATO")
        statoPaziente.donatore = rsDataset("STATODONATORE")
    End If
    rsDataset.Close
    
    If statoPaziente.statoPaz = TPOSPITE Then
        Call CaricaStatoOspite
    End If
    
    Set rsDataset = Nothing
    modifica = True
    blnModificato = False
    blnCaricamentoPaziente = False
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    Dim somma As Integer
    If Index = 0 And oData(Index).data <> "" Then
        If Month(oData(0).data) > Month(date) Then
            somma = -1
        ElseIf Month(oData(0).data) = Month(date) And Day(oData(0).data) > Day(date) Then
            somma = -1
        Else
            somma = 0
        End If
        lblEta = Year(date) - Year(oData(0).data) + somma
    End If
    blnModificato = True
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

Private Sub optRh_Click(Index As Integer)
    Call ColoraSel(optRh, Index, 2)
    blnModificato = True
End Sub

Private Sub optSesso_Click(Index As Integer)
    Call ColoraSel(optSesso, Index, 2)
    blnModificato = True
End Sub

Private Sub picMostraStato_Click()
    ' evita in dialisi e ambulatoriale
    If cboStato.ListIndex = 0 Or cboStato.ListIndex = 5 Then Exit Sub
    frmStatoPaz.Show 1
End Sub

Private Sub picMostraStato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMostraStato.Picture = LoadResPicture("pin2", 0)
End Sub

Private Sub picMostraStato_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMostraStato.Picture = LoadResPicture("pin1", 0)
End Sub

Private Sub tabScheda_Click(PreviousTab As Integer)
    If tabScheda.Tab = 0 Then
        cboCentroProv.SelStart = 0
        Call DisabilitaMedico
    Else
        If intMedicoKey = 0 And NuovoPaziente = False Then
            cmdTrova(1).Enabled = True
            Exit Sub
        ElseIf NuovoPaziente = True Then
            cmdTrova(1).Enabled = False
        End If

        If txtCognome.Text <> "" And txtNome.Text <> "" And oData(0).data <> "" And txtCodiceFiscale.Text <> "" Then
            Call AbilitaMedico
            End If
        End If
End Sub
Private Sub AbilitaMedico()
    cmdTrova(1).Enabled = True
    txtCognomeMedico.Enabled = True
    txtNomeMedico.Enabled = True
    txtCittaMedico.Enabled = True
    txtCapMedico.Enabled = True
    txtProvMedico.Enabled = True
    txtIndirizzoMedico.Enabled = True
    txtStudioMedico.Enabled = True
    txtTelefonoMedico.Enabled = True
    txtCellulareMedico.Enabled = True
    txtEmailMedico.Enabled = True
    txtFaxMedico.Enabled = True
    txtCodiceRegionaleMedico.Enabled = True
    chkPresenzaBarCode.Enabled = True
    txtRiceve.Enabled = True
    cboTipologia.Enabled = True
    Label1(32).Enabled = True
    Label1(33).Enabled = True
    Label1(34).Enabled = True
    Label1(35).Enabled = True
    Label1(36).Enabled = True
    Label1(37).Enabled = True
    Label1(38).Enabled = True
    Label1(39).Enabled = True
    Label1(40).Enabled = True
    Label1(41).Enabled = True
    Label1(42).Enabled = True
    Label1(47).Enabled = True
    Label1(50).Enabled = True
    lblTipologiaMedico(47).Enabled = True
End Sub

Private Sub DisabilitaMedico()
    cmdTrova(1).Enabled = False
    txtCognomeMedico.Enabled = False
    txtNomeMedico.Enabled = False
    txtCittaMedico.Enabled = False
    txtCapMedico.Enabled = False
    txtProvMedico.Enabled = False
    txtIndirizzoMedico.Enabled = False
    txtStudioMedico.Enabled = False
    txtTelefonoMedico.Enabled = False
    txtCellulareMedico.Enabled = False
    txtEmailMedico.Enabled = False
    txtFaxMedico.Enabled = False
    txtCodiceRegionaleMedico.Enabled = False
    chkPresenzaBarCode.Enabled = False
    txtRiceve.Enabled = False
    cboTipologia.Enabled = False
    Label1(32).Enabled = False
    Label1(33).Enabled = False
    Label1(34).Enabled = False
    Label1(35).Enabled = False
    Label1(36).Enabled = False
    Label1(37).Enabled = False
    Label1(38).Enabled = False
    Label1(39).Enabled = False
    Label1(40).Enabled = False
    Label1(41).Enabled = False
    Label1(42).Enabled = False
    Label1(47).Enabled = False
    Label1(50).Enabled = False
    lblTipologiaMedico(47).Enabled = False
End Sub

Private Sub txtAllergia_GotFocus()
    txtAllergia.BackColor = colArancione
End Sub

Private Sub txtAllergia_LostFocus()
    txtAllergia.BackColor = vbWhite
End Sub

Private Sub txtCap_LostFocus(Index As Integer)
    txtCap(Index).BackColor = vbWhite
End Sub

Private Sub txtCapMedico_GotFocus()
    txtCapMedico.BackColor = colArancione
End Sub

Private Sub txtCapMedico_LostFocus()
    txtCapMedico.BackColor = vbWhite
End Sub

Private Sub txtCellulare_LostFocus()
    txtCellulare.BackColor = vbWhite
End Sub



Private Sub txtCellulareMedico_GotFocus()
    txtCellulareMedico.BackColor = colArancione
End Sub

Private Sub txtCellulareMedico_LostFocus()
    txtCellulareMedico.BackColor = vbWhite
End Sub

Private Sub txtCitta_LostFocus()
    txtCitta.BackColor = vbWhite
End Sub

Private Sub txtCittaMedico_GotFocus()
    txtCittaMedico.BackColor = colArancione
End Sub

Private Sub txtCittaMedico_LostFocus()
    txtCittaMedico.BackColor = vbWhite
End Sub

Private Sub txtCodiceFiscale_GotFocus()
    txtCodiceFiscale.BackColor = colArancione
End Sub

Private Sub txtCodiceFiscale_LostFocus()
    txtCodiceFiscale.BackColor = vbWhite
End Sub

Private Sub txtCodiceFiscale_Validate(Cancel As Boolean)
    If UCase(Mid(txtCodiceFiscale, 1, 3)) = "ENI" Or UCase(Mid(txtCodiceFiscale, 1, 3)) = "STP" Then
        Exit Sub
    End If
        
    If txtCodiceFiscale = "" Then
        Cancel = False
    Else
        If Len(txtCodiceFiscale) = 16 Then
            If UCase(cboNazione.Text) = UCase("Italia") Then
                Cancel = Not ControlloCodiceFiscale(UCase(txtCodiceFiscale))
            Else
                Cancel = False
            End If
        Else
            If UCase(cboNazione.Text) = UCase("Italia") Then
                MsgBox "CODICE FISCALE ERRATO" & vbCrLf & "Inserire correttamente tutte le 16 cifre/lettere", vbCritical, "Attenzione"
            Else
                MsgBox "CODICE STP o ENI ERRATO" & vbCrLf & "Inserire correttamente tutte le 16 cifre/lettere", vbCritical, "Attenzione"
            End If
            Cancel = True
            Exit Sub
        End If
    End If
    If Cancel Then
        MsgBox "Il valore inserito è errato", vbCritical, "Attenzione"
        txtCodiceFiscale.SelStart = 0
        txtCodiceFiscale.SelLength = Len(txtCodiceFiscale)
    End If
End Sub

Private Sub txtAllergia_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCodiceFiscale_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCodiceId_GotFocus()
    txtCodiceId.BackColor = colArancione
End Sub

Private Sub txtCodiceId_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCodiceId_LostFocus()
    txtCodiceId.BackColor = vbWhite
End Sub

Private Sub txtCodiceRegionaleMedico_GotFocus()
    txtCodiceRegionaleMedico.BackColor = colArancione
End Sub

Private Sub txtCodiceRegionaleMedico_LostFocus()
    txtCodiceRegionaleMedico.BackColor = vbWhite
End Sub

Private Sub txtCognome_GotFocus()
    txtCognome.BackColor = colArancione
End Sub

Private Sub txtcogNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCap_GotFocus(Index As Integer)
    txtCap(Index).BackColor = colArancione
End Sub

Private Sub txtCellulare_GotFocus()
    txtCellulare.BackColor = colArancione
End Sub

Private Sub txtCitta_GotFocus()
    txtCitta.BackColor = colArancione
End Sub

Private Sub txtCognome_LostFocus()
    txtCognome.BackColor = vbWhite
End Sub

Private Sub txtCognomeMedico_GotFocus()
    txtCognomeMedico.BackColor = colArancione
End Sub

Private Sub txtCognomeMedico_LostFocus()
    txtCognomeMedico.BackColor = vbWhite
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.BackColor = colArancione
End Sub

Private Sub txtEmail_LostFocus()
    txtEmail.BackColor = vbWhite
End Sub

Private Sub txtEmailMedico_GotFocus()
    txtEmailMedico.BackColor = colArancione
End Sub

Private Sub txtEmailMedico_LostFocus()
    txtEmailMedico.BackColor = vbWhite
End Sub

Private Sub txtFax_GotFocus()
    txtFax.BackColor = colArancione
End Sub

Private Sub txtFax_LostFocus()
    txtFax.BackColor = vbWhite
End Sub

Private Sub txtFaxMedico_GotFocus()
    txtFaxMedico.BackColor = colArancione
End Sub

Private Sub txtFaxMedico_LostFocus()
    txtFaxMedico.BackColor = vbWhite
End Sub

Private Sub txtIndirizzo_GotFocus()
    txtIndirizzo.BackColor = colArancione
End Sub

Private Sub txtIndirizzo_LostFocus()
    txtIndirizzo.BackColor = vbWhite
End Sub

Private Sub txtIndirizzoMedico_GotFocus()
    txtIndirizzoMedico.BackColor = colArancione
End Sub

Private Sub txtIndirizzoMedico_LostFocus()
    txtIndirizzoMedico.BackColor = vbWhite
End Sub

Private Sub txtKm_GotFocus()
    txtKm.BackColor = colArancione
End Sub

Private Sub txtKm_LostFocus()
    txtKm.BackColor = vbWhite
End Sub

Private Sub txtNome_GotFocus()
    txtNome.BackColor = colArancione
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = vbWhite
End Sub

Private Sub txtNomeMedico_GotFocus()
    txtNomeMedico.BackColor = colArancione
End Sub

Private Sub txtNomeMedico_LostFocus()
    txtNomeMedico.BackColor = vbWhite
End Sub

Private Sub txtNote_GotFocus()
    txtNote.BackColor = colArancione
End Sub

Private Sub txtNote_LostFocus()
    txtNote.BackColor = vbWhite
End Sub

Private Sub txtNumCarta_GotFocus()
    txtNumCarta.BackColor = colArancione
End Sub

Private Sub txtNumCarta_LostFocus()
    txtNumCarta.BackColor = vbWhite
End Sub

Private Sub txtNumeroProcura_GotFocus()
    txtNumeroProcura.BackColor = colArancione
End Sub

Private Sub txtNumeroProcura_LostFocus()
    txtNumeroProcura.BackColor = vbWhite
End Sub

Private Sub txtProfessione_GotFocus()
    txtProfessione.BackColor = colArancione
End Sub

Private Sub txtProfessione_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtProfessione_LostFocus()
    txtProfessione.BackColor = vbWhite
End Sub

Private Sub txtProv_GotFocus(Index As Integer)
    txtProv(Index).BackColor = colArancione
End Sub

Private Sub txtProv_LostFocus(Index As Integer)
    txtProv(Index).BackColor = vbWhite
End Sub

Private Sub txtProvMedico_GotFocus()
    txtProvMedico.BackColor = colArancione
End Sub

Private Sub txtProvMedico_LostFocus()
    txtProvMedico.BackColor = vbWhite
End Sub

Private Sub txtRiceve_GotFocus()
    txtRiceve.BackColor = colArancione
End Sub

Private Sub txtRiceve_LostFocus()
    txtRiceve.BackColor = vbWhite
End Sub

Private Sub txtRilascioCarta_GotFocus()
    txtRilascioCarta.BackColor = colArancione
End Sub

Private Sub txtRilascioCarta_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtNumCarta_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtRilascioCarta_LostFocus()
    txtRilascioCarta.BackColor = vbWhite
End Sub

Private Sub txtStudioMedico_GotFocus()
    txtStudioMedico.BackColor = colArancione
End Sub

Private Sub txtStudioMedico_LostFocus()
    txtStudioMedico.BackColor = vbWhite
End Sub

Private Sub txtTelefono_GotFocus()
    txtTelefono.BackColor = colArancione
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.BackColor = vbWhite
End Sub

Private Sub txtTelefonoMedico_GotFocus()
    txtTelefonoMedico.BackColor = colArancione
End Sub

Private Sub txtTelefonoMedico_LostFocus()
    txtTelefonoMedico.BackColor = vbWhite
End Sub

Private Sub txtTesseraSanitaria_GotFocus()
    txtTesseraSanitaria.BackColor = colArancione
End Sub

Private Sub txtTesseraSanitaria_LostFocus()
    txtTesseraSanitaria.BackColor = vbWhite
End Sub

Private Sub txtNumeroProcura_Change()
    If lettera = "" Then Exit Sub
    Call OnlyNumber(txtNumeroProcura, lettera)
    blnModificato = True
End Sub

Private Sub txtNumeroProcura_KeyPress(KeyAscii As Integer)
    lettera = Chr(KeyAscii)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCellulare_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

'******** Gestione Modificato

Private Sub txtCognome_Change()
    blnModificato = True
End Sub

Private Sub txtNome_Change()
    blnModificato = True
End Sub

Private Sub txtCodiceId_Change()
    blnModificato = True
End Sub

Private Sub txtCitta_Change()
    blnModificato = True
End Sub

Private Sub txtCAP_Change(Index As Integer)
    blnModificato = True
End Sub

Private Sub txtProv_Change(Index As Integer)
    blnModificato = True
End Sub

Private Sub txtKm_Change()
    blnModificato = True
End Sub

Private Sub txtIndirizzo_Change()
    blnModificato = True
End Sub

Private Sub txtTelefono_Change()
    blnModificato = True
End Sub

Private Sub txtCellulare_Change()
    blnModificato = True
End Sub

Private Sub txtEmail_Change()
    blnModificato = True
End Sub

Private Sub txtFax_Change()
    blnModificato = True
End Sub

Private Sub txtNumCarta_Change()
    blnModificato = True
End Sub

Private Sub txtRilascioCarta_Change()
    blnModificato = True
End Sub

Private Sub txtCodiceFiscale_Change()
    blnModificato = True
End Sub

Private Sub txtProfessione_Change()
    blnModificato = True
End Sub

Private Sub txtTesseraSanitaria_Change()
    blnModificato = True
End Sub

Private Sub txtNote_Change()
    blnModificato = True
End Sub

Private Sub txtAllergia_Change()
    blnModificato = True
End Sub

Private Sub chkTrasportoInAmbulanza_Click()
    blnModificato = True
End Sub

Private Sub cboAccompagnatore_Click()
    blnModificato = True
End Sub

Private Sub cboCentroProv_Click()
    blnModificato = True
End Sub

Private Sub cboCentroProv_Change()
    blnModificato = True
End Sub

Private Sub cboComuneResidenza_Click()
    blnModificato = True
End Sub

Private Sub cboDistretto_Click()
    blnModificato = True
End Sub

Private Sub cboDocumento_Click()
    blnModificato = True
End Sub

Private Sub cboEsenzione_Click()
    blnModificato = True
End Sub

Private Sub cboGSanguigno_Click()
    blnModificato = True
End Sub

Private Sub cboRegione_Click()
    If stoPulendo Then Exit Sub
    If cboRegione.ListIndex = -1 Then Exit Sub
    Call RicaricaComboBox("SELECT * FROM COMUNI WHERE REGIONIID=" & cboRegione.ItemData(cboRegione.ListIndex), "NOME", cboComuneResidenza)
    blnModificato = True
End Sub

Private Sub chkEsenteReddito_Click()
    blnModificato = True
End Sub

Private Sub lblCognomeMed_Change()
    blnModificato = True
End Sub

