VERSION 5.00
Begin VB.Form frmFatEle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fattura Elettronica"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   59
      Top             =   8160
      Width           =   7695
      Begin VB.TextBox txtBolloFattura 
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
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   28
         Top             =   240
         Width           =   612
      End
      Begin VB.TextBox txtAutorizzazioneBollo 
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
         MaxLength       =   14
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblProgrInvio 
         Caption         =   "N° Progressivo Invio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblAutorizzazione 
         Caption         =   "N° Autorizzazione Bollo Virtuale "
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
         Left            =   120
         TabIndex        =   61
         Top             =   140
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Bollo su Fattura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   22
         Left            =   4200
         TabIndex        =   60
         Top             =   240
         Width           =   1692
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   53
      Top             =   9240
      Width           =   7695
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
         Left            =   5080
         TabIndex        =   73
         Top             =   240
         Width           =   1380
      End
      Begin VB.CommandButton cmdEsci 
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
         Left            =   6480
         TabIndex        =   30
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton GeneraFE 
         Caption         =   "Genera Fattura Elettronica (XML)"
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
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tutti i campi vanno valorizzati"
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
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Coordinate Bonifico Bancario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   47
      Top             =   6720
      Width           =   7695
      Begin VB.TextBox txtIbanNum 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtIbanAlfa 
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
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtIbanAlfa 
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtIbanNum 
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
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtIbanNum 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtIbanNum 
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
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtIntestatario 
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
         MaxLength       =   50
         TabIndex        =   20
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IBAN"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Intestatario c/c"
         BeginProperty Font 
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
         TabIndex        =   48
         Top             =   480
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Committente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   38
      Top             =   4320
      Width           =   7695
      Begin VB.ComboBox cboProvCommittente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   5880
         TabIndex        =   72
         Top             =   1440
         Width           =   804
      End
      Begin VB.TextBox txtCodiceDestinatario 
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
         Left            =   6720
         MaxLength       =   6
         TabIndex        =   14
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txtIndirizzoFattura 
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   15
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtCapFattura 
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
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPartitaIvaFattura 
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
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
         Height          =   288
         ItemData        =   "frmFatEle.frx":0000
         Left            =   1320
         List            =   "frmFatEle.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox cboComune 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtCodFiscaleFattura 
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
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Cod. Destinatario"
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
         Left            =   4800
         TabIndex        =   57
         Top             =   480
         Width           =   1935
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
         Index           =   21
         Left            =   120
         TabIndex        =   46
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CAP"
         BeginProperty Font 
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
         Left            =   5160
         TabIndex        =   45
         Top             =   990
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Città"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   480
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
         Index           =   18
         Left            =   5160
         TabIndex        =   43
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P. Iva"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   16
         Left            =   120
         TabIndex        =   41
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "ASL a cui fatturare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.F."
         BeginProperty Font 
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
         Left            =   5280
         TabIndex        =   39
         Top             =   1920
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prestatore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4332
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cboProvPrestatore 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   6240
         TabIndex        =   71
         Top             =   1250
         Width           =   804
      End
      Begin VB.TextBox txtMail 
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
         MaxLength       =   40
         TabIndex        =   69
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtCapSociale 
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
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3840
         Width           =   1900
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   372
         Left            =   1920
         TabIndex        =   64
         Top             =   3360
         Width           =   1572
         Begin VB.OptionButton srlno 
            Caption         =   "no"
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
            Left            =   720
            TabIndex        =   66
            Top             =   80
            Width           =   615
         End
         Begin VB.OptionButton srlsi 
            Caption         =   "si"
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
            Left            =   0
            TabIndex        =   65
            Top             =   80
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   372
         Left            =   1680
         TabIndex        =   63
         Top             =   3000
         Width           =   1932
         Begin VB.OptionButton Liquidaz_no 
            Caption         =   "no"
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
            Left            =   960
            TabIndex        =   68
            Top             =   80
            Width           =   615
         End
         Begin VB.OptionButton Liquidaz_si 
            Caption         =   "si"
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
            Left            =   240
            TabIndex        =   67
            Top             =   80
            Width           =   615
         End
      End
      Begin VB.ComboBox cboProvUffReg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   6760
         TabIndex        =   7
         Top             =   2640
         Width           =   800
      End
      Begin VB.ComboBox cboRegimeFiscale 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "frmFatEle.frx":0004
         Left            =   5620
         List            =   "frmFatEle.frx":0011
         TabIndex        =   10
         Top             =   3840
         Width           =   2000
      End
      Begin VB.CheckBox chkSocioPiu 
         Caption         =   "Più Soci"
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
         Left            =   5560
         TabIndex        =   9
         Top             =   3432
         Width           =   1335
      End
      Begin VB.CheckBox chkSocioUnico 
         Caption         =   "Socio Unico"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   3432
         Width           =   1815
      End
      Begin VB.TextBox txtNumRea 
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
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtRagione 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   260
         Width           =   5415
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
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   2
         Top             =   780
         Width           =   3615
      End
      Begin VB.TextBox txtCap 
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
         TabIndex        =   3
         Top             =   780
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1250
         Width           =   3615
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
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1700
         Width           =   2055
      End
      Begin VB.TextBox txtIva 
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
         Left            =   5340
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1700
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   70
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "Regime Fiscale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3936
         TabIndex        =   58
         Top             =   3840
         Width           =   1668
      End
      Begin VB.Label Label4 
         Caption         =   "In Liquidazione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   108
         TabIndex        =   56
         Top             =   3084
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "S.R.L."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   108
         TabIndex        =   55
         Top             =   3432
         Width           =   732
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capitale Sociale"
         BeginProperty Font 
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
         Left            =   96
         TabIndex        =   52
         Top             =   3840
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° REA"
         BeginProperty Font 
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
         TabIndex        =   51
         Top             =   2664
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Uff. Registrazione Società->Prov."
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
         Index           =   0
         Left            =   4740
         TabIndex        =   50
         Top             =   2580
         Width           =   1908
         WordWrap        =   -1  'True
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
         Index           =   1
         Left            =   100
         TabIndex        =   37
         Top             =   780
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CAP"
         BeginProperty Font 
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
         Left            =   5700
         TabIndex        =   36
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comune"
         BeginProperty Font 
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
         Left            =   100
         TabIndex        =   35
         Top             =   1250
         Width           =   855
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
         Index           =   4
         Left            =   5640
         TabIndex        =   34
         Top             =   1250
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice fiscale"
         BeginProperty Font 
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
         Left            =   100
         TabIndex        =   33
         Top             =   1700
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P. Iva"
         BeginProperty Font 
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
         Left            =   4680
         TabIndex        =   32
         Top             =   1700
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Denominazione Centro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   100
         TabIndex        =   31
         Top             =   250
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmFatEle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lettera As String
Dim rsDataset As Recordset
Dim modifica As Boolean

Private Function Completo() As Boolean
Dim nome As String
    Completo = False
    
    If txtRagione.Text = "" Then
        nome = "RAGIONE SOCIALE"
    ElseIf txtIndirizzo.Text = "" Then
        nome = "INDIRIZZO del CENTRO"
    ElseIf txtCap.Text = "" Then
        nome = "CAP del CENTRO"
    ElseIf txtCitta.Text = "" Then
        nome = "COMUNE del CENTRO"
    ElseIf cboProvPrestatore.Text = "" Then
        nome = "PROV del CENTRO"
    ElseIf txtCodiceFiscale.Text = "" Then
        nome = "CODICE FISCALE del CENTRO"
    ElseIf txtIva = "" Then
        nome = "PARTITA IVA CENTRO"
    ElseIf txtMail = "" Then
        nome = "EMAIL"
    ElseIf cboProvUffReg.Text = "" Then
        nome = "UFFICIO REGISTRAZIONE SOCIETA'->PROV."
    ElseIf srlno.Value = False And srlsi.Value = False Then
        nome = "S.R.L."
    ElseIf srlsi.Value = True And chkSocioUnico.Value = Unchecked And chkSocioPiu.Value = Unchecked Then
        nome = "SOCIO UNICO - PIU' SOCI"
    ElseIf Liquidaz_no.Value = False And Liquidaz_si.Value = False Then
        nome = "IN LIQUIDAZIONE"
    ElseIf cboRegimeFiscale.ListIndex = -1 Then
        nome = "REGIME FISCALE"
    ElseIf txtCapSociale.Text = "" Then
        nome = "CAPITALE SOCIALE"
    ElseIf txtNumRea.Text = "" Then
        nome = "N° REA"
    
    ElseIf cboAsl.ListIndex = -1 Then
        nome = "ASL a cui FATTURARE"
    ElseIf txtCodiceDestinatario.Text = "" Then
        nome = "CODICE del DESTINATARIO"
    ElseIf txtIndirizzoFattura.Text = "" Then
        nome = "INDIRIZZO del COMMITTENTE"
    ElseIf txtCapFattura.Text = "" Then
        nome = "CAP del COMMITTENTE"
    ElseIf cboComune.ListIndex = -1 Then
        nome = "COMUNE del COMMITTENTE"
    ElseIf cboProvCommittente.Text = "" Then
        nome = "PROV del COMMITTENTE"
    ElseIf txtPartitaIvaFattura = "" Then
        nome = "PARTITA IVA COMMITTENTE"
    ElseIf txtCodFiscaleFattura.Text = "" Then
        nome = "CODICE FISCALE COMMITTENTE"

    ElseIf txtIntestatario.Text = "" Then
        nome = "INTESTATARIO C/C"
    ElseIf txtIbanAlfa(0).Text = "" Then
        nome = "IBAN"
    ElseIf txtIbanAlfa(1).Text = "" Then
        nome = "IBAN"
    ElseIf txtIbanNum(0).Text = "" Then
        nome = "IBAN"
    ElseIf txtIbanNum(1).Text = "" Then
        nome = "IBAN"
    ElseIf txtIbanNum(2).Text = "" Then
        nome = "IBAN"
    ElseIf txtIbanNum(3).Text = "" Then
        nome = "IBAN"
        
    ElseIf txtAutorizzazioneBollo.Text = "" Then
        nome = "AUTORIZZAZIONE BOLLO"
    ElseIf txtBolloFattura.Text = "" Then
        nome = "BOLLO SU FATTURA"
    Else
        Completo = True
        Exit Function
    End If
    MsgBox "Inserire i dati obbligatori" & vbCrLf & "Campo: " & nome, vbInformation, "ATTENZIONE!!!"
End Function

Private Sub chkSocioPiu_GotFocus()
    chkSocioUnico.Value = Unchecked
End Sub

Private Sub chkSocioUnico_GotFocus()
    chkSocioPiu.Value = Unchecked
End Sub

Private Sub cmdMemorizza_Click()
    If Len(txtAutorizzazioneBollo.Text) < 14 Then
        MsgBox "Il N° di AUTORIZZAZIONE BOLLO VIRTUALE NON può essere inferiore a 14 cifre", vbInformation, "ATTENZIONE!!!"
        Exit Sub
    End If
    If Completo Then
        Call MemorizzaIntestazione
        Call MemorizzaFattura
        MsgBox "I dati sono stati memorizzati nell'archivio", vbInformation, "Informazione"
    End If
End Sub

Private Sub MemorizzaIntestazione()
    Dim v_Val() As Variant
    Dim v_nome() As Variant
    
    v_nome = Array("KEY", "RAGIONE_SOCIALE", "INDIRIZZO", "CAP", "CITTA", "PROV", "CODICE_FISCALE", "IVA", "MAIL", "PR_UFF_REG", "SRL", "SOCIO", "LIQUIDAZIONE", "REG_FISCALE", "CAP_SOCIALE", "NUM_REA")
    v_Val = Array(1, txtRagione, txtIndirizzo, txtCap, txtCitta, cboProvPrestatore.Text, txtCodiceFiscale, txtIva, txtMail, cboProvUffReg.Text, GestisciSrl, GestisciSocio, LiquidazioneSiNo, GestisciRegimeFiscale, txtCapSociale, txtNumRea)
    Set rsDataset = New Recordset
    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        If modifica Then
            rsDataset.Update v_nome, v_Val
        Else
            rsDataset.AddNew v_nome, v_Val
            rsDataset.Update
        End If
    Set rsDataset = Nothing
    Call CaricaVarPublic
    
End Sub

Private Function GestisciSrl() As String
    If srlsi.Value = True Then
        GestisciSrl = 1
    ElseIf srlno.Value = True Then
         GestisciSrl = 0
    End If
End Function

Private Function GestisciRegimeFiscale() As String
    If cboRegimeFiscale.Text = "Altro" Then
        GestisciRegimeFiscale = "RF18"
    ElseIf cboRegimeFiscale.Text = "Contribuenti Minimi" Then
        GestisciRegimeFiscale = "RF02"
    ElseIf cboRegimeFiscale.Text = "Ordinario" Then
        GestisciRegimeFiscale = "RF01"
    End If
End Function

Private Function GestisciSocio() As String
    If chkSocioUnico.Value = Checked Then
        GestisciSocio = "SU"
    ElseIf chkSocioPiu.Value = Checked Then
        GestisciSocio = "SM"
    End If
End Function

Private Function LiquidazioneSiNo() As String
    If Liquidaz_si.Value = True Then
        LiquidazioneSiNo = "LS"
    ElseIf Liquidaz_no.Value = True Then
        LiquidazioneSiNo = "LN"
    End If
End Function

Private Sub MemorizzaFattura()
    Dim v_Val() As Variant
    Dim v_nome() As Variant
    Dim strIban As String
    
    Call SuperUcase(Me)
    
    strIban = txtIbanAlfa(0) & txtIbanNum(0) & txtIbanAlfa(1) & txtIbanNum(1) & txtIbanNum(2) & txtIbanNum(3)
    v_nome = Array("KEY", "CODICE_ASL", "COD_DESTINATARIO", "INDIRIZZO", "CAP", "CODICE_COMUNE", "PROV", "P_IVA", "CODICE_FISCALE", "INTESTATARIO_CC", "IBAN", "NUMERO_AUTORIZZAZIONE", "IMPORTO_BOLLO") ', "PROGR_INVIO")
    v_Val = Array(1, cboAsl.ItemData(cboAsl.ListIndex), txtCodiceDestinatario, txtIndirizzoFattura, txtCapFattura, cboComune.ItemData(cboComune.ListIndex), cboProvCommittente.Text, txtPartitaIvaFattura, txtCodFiscaleFattura, txtIntestatario, strIban, txtAutorizzazioneBollo, txtBolloFattura) ', txtProgrInvio)
        
    Set rsDataset = New Recordset
    rsDataset.Open "INTESTAZIONE_FATTURA", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        If modifica Then
            rsDataset.Update v_nome, v_Val
        Else
            rsDataset.AddNew v_nome, v_Val
            rsDataset.Update
        End If
    Set rsDataset = Nothing
    
End Sub

Private Sub Form_Activate()
    Call CaricaIntestazione
    Call CaricaParametriFattura
End Sub

Private Sub CaricaIntestazione()

    Call RicaricaComboBox("SIGLE_PROVINCIE", "NOME", cboProvPrestatore)
    Call RicaricaComboBox("SIGLE_PROVINCIE", "NOME", cboProvUffReg)

    Set rsDataset = New Recordset
    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        txtRagione = rsDataset("RAGIONE_SOCIALE")
        txtIndirizzo = rsDataset("INDIRIZZO")
        txtCap = rsDataset("CAP")
        txtCitta = rsDataset("CITTA")
        cboProvPrestatore.Text = rsDataset("PROV") & ""
        txtCodiceFiscale = rsDataset("CODICE_FISCALE")
        txtIva = rsDataset("IVA")
        txtMail = rsDataset("MAIL") & ""
        cboProvUffReg.Text = rsDataset("PR_UFF_REG") & ""
        txtCapSociale = VirgolaOrPunto(rsDataset("CAP_SOCIALE") & "", ",")
        txtNumRea = rsDataset("NUM_REA") & ""
        Call CaricaSrl
        Call CaricaSocio
        Call CaricaLiquidazione
        Call CaricaRegimeFiscale
        modifica = True
    Else
        srlsi.Value = False
        srlno.Value = False
        Liquidaz_si.Value = False
        Liquidaz_no.Value = False
        modifica = False
    End If
    rsDataset.Clone
    Set rsDataset = Nothing
End Sub

Private Sub CaricaSrl()
    If rsDataset("SRL") = 1 Then
        srlsi.Value = True
        chkSocioUnico.Enabled = True
        chkSocioPiu.Enabled = True
    ElseIf rsDataset("SRL") = 0 Then
        srlno.Value = True
    ElseIf rsDataset("SRL") = 2 Then
        srlsi.Value = False
        srlno.Value = False
    End If
End Sub

Private Sub CaricaRegimeFiscale()
    If rsDataset("REG_FISCALE") = "RF18" Then
        cboRegimeFiscale.ListIndex = 0
    ElseIf rsDataset("REG_FISCALE") = "RF02" Then
        cboRegimeFiscale.ListIndex = 1
    ElseIf rsDataset("REG_FISCALE") = "RF01" Then
        cboRegimeFiscale.ListIndex = 2
    End If
End Sub

Private Sub CaricaSocio()
    If rsDataset("SOCIO") = "SU" Then
        chkSocioUnico.Value = Checked
    ElseIf rsDataset("SOCIO") = "SM" Then
        chkSocioPiu.Value = Checked
    End If
End Sub

Private Sub CaricaLiquidazione()
    If rsDataset("LIQUIDAZIONE") = "LS" Then
        Liquidaz_si.Value = True
    ElseIf rsDataset("LIQUIDAZIONE") = "LN" Then
        Liquidaz_no.Value = True
    ElseIf IsNull(rsDataset("LIQUIDAZIONE")) Then
        Liquidaz_si.Value = False
        Liquidaz_no.Value = False
    End If
End Sub

Private Sub CaricaParametriFattura()
Dim strIban As String
Dim strSql As String
        
    Call RicaricaComboBox("SIGLE_PROVINCIE", "NOME", cboProvCommittente)
    Call RicaricaComboBox("ASL", "NOME", cboAsl)
    Call RicaricaComboBox("COMUNI", "NOME", cboComune)
    
    strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.KEY AS ASLKEY, COMUNI.KEY AS COMUNIKEY " & _
            "FROM       (INTESTAZIONE_FATTURA " & _
            "           LEFT OUTER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
            "           LEFT OUTER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE"
    Set rsDataset = New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        cboAsl.ListIndex = GetCboListIndex(rsDataset("ASLKEY"), cboAsl)
        txtCodiceDestinatario = rsDataset("COD_DESTINATARIO") & ""
        txtIndirizzoFattura = rsDataset("INDIRIZZO")
        txtCapFattura = rsDataset("CAP")
        cboComune.ListIndex = GetCboListIndex(rsDataset("COMUNIKEY"), cboComune)
        cboProvCommittente.Text = rsDataset("PROV") & ""
        txtPartitaIvaFattura = rsDataset("P_IVA")
        txtCodFiscaleFattura = rsDataset("CODICE_FISCALE")
        txtIntestatario = rsDataset("INTESTATARIO_CC")
        strIban = rsDataset("IBAN")
        txtIbanAlfa(0) = Mid(strIban, 1, 2)
        txtIbanNum(0) = Mid(strIban, 3, 2)
        txtIbanAlfa(1) = Mid(strIban, 5, 1)
        txtIbanNum(1) = Mid(strIban, 6, 5)
        txtIbanNum(2) = Mid(strIban, 11, 5)
        txtIbanNum(3) = Mid(strIban, 16, 12)
        txtAutorizzazioneBollo = rsDataset("NUMERO_AUTORIZZAZIONE")
        txtBolloFattura = VirgolaOrPunto(rsDataset("IMPORTO_BOLLO"), ",") & ""
'        If VarType(rsDataset("PROGR_INVIO")) = 1 Then 'se il campo è null VarType assume valore 1
'            txtProgrInvio = 1
'        End if
         lblProgrInvio = "N° Progressivo Invio -> " & rsDataset("PROGR_INVIO")
        modifica = True
    Else
        modifica = False
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Sub cmdEsci_Click()
    Unload frmFatEle
End Sub

Private Sub GeneraFE_Click()
    If Completo Then
        If MsgBox("La generazione della fattura elettronica comporta l'attribuzione definitiva" & vbCrLf & "e non modificabile del N° Progressivo d'Invio - SI E' SICURI DI PROCEDERE?", vbQuestion + vbYesNo + vbDefaultButton2, "PRESTARE ATTENZIONE!!!") = vbNo Then
            Exit Sub
        End If
        OKGeneraFE = True
        Unload frmFatEle
    End If
End Sub

Private Sub srlno_Click()
    chkSocioUnico.Enabled = False
    chkSocioPiu.Enabled = False
    chkSocioUnico.Value = Unchecked
    chkSocioPiu.Value = Unchecked
End Sub

Private Sub srlsi_Click()
    chkSocioUnico.Enabled = True
    chkSocioPiu.Enabled = True
End Sub

Private Sub txtAutorizzazioneBollo_GotFocus()
    txtAutorizzazioneBollo.BackColor = colArancione
End Sub

Private Sub txtAutorizzazioneBollo_LostFocus()
    txtAutorizzazioneBollo.BackColor = vbWhite
    If Len(txtAutorizzazioneBollo.Text) < 14 Then
        MsgBox "Il N° di AUTORIZZAZIONE BOLLO VIRTUALE NON può essere inferiore a 14 cifre/caratteri", vbInformation, "ATTENZIONE!!!"
        txtAutorizzazioneBollo.SetFocus
    End If
End Sub

Private Sub txtBolloFattura_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtBolloFattura, lettera)
End Sub

Private Sub txtBolloFattura_GotFocus()
    txtBolloFattura.BackColor = colArancione
End Sub

Private Sub txtBolloFattura_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtBolloFattura_LostFocus()
    txtBolloFattura.BackColor = vbWhite
End Sub

Private Sub txtCap_GotFocus()
    txtCap.BackColor = colArancione
End Sub

Private Sub txtCap_LostFocus()
    txtCap.BackColor = vbWhite
End Sub

Private Sub txtCapFattura_GotFocus()
    txtCapFattura.BackColor = colArancione
End Sub

Private Sub txtCapFattura_LostFocus()
    txtCapFattura.BackColor = vbWhite
End Sub

Private Sub txtCapSociale_Change()
    If lettera = "." Or lettera = "" Then Exit Sub
    Call OnlyNumber(txtCapSociale, lettera)
End Sub

Private Sub txtCapSociale_GotFocus()
    txtCapSociale.BackColor = colArancione
End Sub

Private Sub txtCapSociale_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
End Sub

Private Sub txtCapSociale_LostFocus()
    txtCapSociale.BackColor = vbWhite
    If Len(txtCapSociale.Text) < 4 Then
        MsgBox "Il valore del CAPITALE SOCIALE NON può essere inferiore a 4 cifre", vbInformation, "ATTENZIONE!!!"
        txtCapSociale.SetFocus
    End If
End Sub

Private Sub txtCitta_GotFocus()
    txtCitta.BackColor = colArancione
End Sub

Private Sub txtCitta_LostFocus()
    txtCitta.BackColor = vbWhite
End Sub

Private Sub txtCodFiscaleFattura_GotFocus()
    txtCodFiscaleFattura.BackColor = colArancione
End Sub

Private Sub txtCodFiscaleFattura_LostFocus()
    txtCodFiscaleFattura.BackColor = vbWhite
End Sub

Private Sub txtCodiceDestinatario_GotFocus()
    txtCodiceDestinatario.BackColor = colArancione
End Sub

Private Sub txtCodiceDestinatario_LostFocus()
    txtCodiceDestinatario.BackColor = vbWhite
End Sub

Private Sub txtCodiceFiscale_GotFocus()
    txtCodiceFiscale.BackColor = colArancione
End Sub

Private Sub txtCodiceFiscale_LostFocus()
    txtCodiceFiscale.BackColor = vbWhite
End Sub

Private Sub txtIndirizzo_GotFocus()
    txtIndirizzo.BackColor = colArancione
End Sub

Private Sub txtIndirizzo_LostFocus()
    txtIndirizzo.BackColor = vbWhite
End Sub

Private Sub txtIndirizzoFattura_GotFocus()
    txtIndirizzoFattura.BackColor = colArancione
End Sub

Private Sub txtIndirizzoFattura_LostFocus()
    txtIndirizzoFattura.BackColor = vbWhite
End Sub

Private Sub txtIntestatario_GotFocus()
    txtIntestatario.BackColor = colArancione
End Sub

Private Sub txtIntestatario_LostFocus()
    txtIntestatario.BackColor = vbWhite
End Sub

Private Sub txtIva_GotFocus()
    txtIva.BackColor = colArancione
End Sub

Private Sub txtIva_LostFocus()
    txtIva.BackColor = vbWhite
End Sub

Private Sub txtMail_GotFocus()
    txtMail.BackColor = colArancione
End Sub

Private Sub txtMail_LostFocus()
    txtMail.BackColor = vbWhite
End Sub

Private Sub txtNumRea_GotFocus()
    txtNumRea.BackColor = colArancione
End Sub

Private Sub txtNumRea_LostFocus()
    txtNumRea.BackColor = vbWhite
End Sub

Private Sub txtPartitaIvaFattura_GotFocus()
    txtPartitaIvaFattura.BackColor = colArancione
End Sub

Private Sub txtPartitaIvaFattura_LostFocus()
    txtPartitaIvaFattura.BackColor = vbWhite
End Sub

Private Sub txtRagione_GotFocus()
    txtRagione.BackColor = colArancione
End Sub

Private Sub txtRagione_LostFocus()
    txtRagione.BackColor = vbWhite
End Sub

Private Sub txtIbanNum_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtIbanAlfa_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("A") To Asc("z"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtIbanAlfa_GotFocus(Index As Integer)
    txtIbanAlfa(Index).BackColor = colArancione
End Sub

Private Sub txtIbanAlfa_LostFocus(Index As Integer)
    txtIbanAlfa(Index).BackColor = vbWhite
End Sub

Private Sub txtIbanNum_GotFocus(Index As Integer)
    txtIbanNum(Index).BackColor = colArancione
End Sub

Private Sub txtIbanNum_LostFocus(Index As Integer)
    txtIbanNum(Index).BackColor = vbWhite
End Sub

'Private Sub txtProgrInvio_GotFocus(Index As Integer)
'    txtProgrInvio(Index).BackColor = colArancione
'End Sub

'Private Sub txtProgrInvio_LostFocus(Index As Integer)
'    txtProgrInvio(Index).BackColor = vbWhite
'End Sub



