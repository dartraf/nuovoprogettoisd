VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIntestazioneCentro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Intestazione Centro"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5052
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox cboProv 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmIntestazioneCentro.frx":0000
         Left            =   6240
         List            =   "frmIntestazioneCentro.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1680
         Width           =   804
      End
      Begin VB.TextBox txtSitoWeb 
         BeginProperty Font 
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
         TabIndex        =   11
         Top             =   3600
         Width           =   5415
      End
      Begin VB.ComboBox cboDistrettoAppartenenza 
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
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4560
         Width           =   735
      End
      Begin VB.ComboBox cboAslAppartenenza 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox txtCodiceSts 
         BeginProperty Font 
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
         MaxLength       =   6
         TabIndex        =   12
         Top             =   4080
         Width           =   1215
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
         TabIndex        =   10
         Top             =   3120
         Width           =   3015
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
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2640
         Width           =   2175
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
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
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
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   6
         Top             =   2160
         Width           =   2055
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
         TabIndex        =   5
         Top             =   1680
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
         TabIndex        =   4
         Top             =   1200
         Width           =   735
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
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtTipo 
         BeginProperty Font 
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
         TabIndex        =   2
         Top             =   720
         Width           =   5415
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
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sito Web"
         BeginProperty Font 
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
         TabIndex        =   45
         Top             =   3600
         Width           =   888
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
         Height          =   432
         Index           =   14
         Left            =   120
         TabIndex        =   34
         Top             =   140
         Width           =   1788
      End
      Begin VB.Label Label1 
         Caption         =   "Distretto di appartenenza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   13
         Left            =   5160
         TabIndex        =   33
         Top             =   4488
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "ASL di appartenenza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Index           =   12
         Left            =   120
         TabIndex        =   32
         Top             =   4440
         Width           =   1428
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Struttura"
         BeginProperty Font 
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
         TabIndex        =   31
         Top             =   4080
         Width           =   1668
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
         TabIndex        =   30
         Top             =   3120
         Width           =   600
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
         Left            =   4350
         TabIndex        =   29
         Top             =   2640
         Width           =   600
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
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   1500
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
         Index           =   6
         Left            =   4560
         TabIndex        =   27
         Top             =   2160
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         BeginProperty Font 
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
         TabIndex        =   26
         Top             =   2160
         Width           =   945
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
         TabIndex        =   25
         Top             =   1680
         Width           =   555
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
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   855
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
         TabIndex        =   23
         Top             =   1215
         Width           =   465
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
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
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
         TabIndex        =   21
         Top             =   720
         Width           =   492
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Width           =   7575
      Begin VB.CommandButton cmdScegli 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5800
         TabIndex        =   44
         ToolTipText     =   "Cerca Logo Aziendale"
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdScegli 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5800
         TabIndex        =   43
         ToolTipText     =   "Cerca Logo Grande"
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdScegli 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5800
         TabIndex        =   42
         ToolTipText     =   "Cerca Logo Piccolo"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtLogoAziendale 
         BeginProperty Font 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtLogoQualita 
         BeginProperty Font 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   16
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtLogoISO 
         BeginProperty Font 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   15
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox chkLogo 
         Caption         =   "Stampa"
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
         Left            =   6240
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkLogoAziendale 
         Caption         =   "Stampa"
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
         Left            =   6240
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkLogoQualita 
         Caption         =   "Stampa"
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
         Left            =   6240
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logo Grande"
         BeginProperty Font 
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
         TabIndex        =   41
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logo Aziendale"
         BeginProperty Font 
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
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logo Piccolo"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   7575
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
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   1575
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
         Left            =   6000
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlApri 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIntestazioneCentro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmIntestazioneCentro.frm
'
' <b>Descrizione</b>: Scheda Intestazione Centro associata alla tab INTESTAZIONE_STAMPA
'
' @remarks
'
' @author
'
' @date 07/02/2011 18.32
Option Explicit

'' rs della scheda
Dim rsDataset As Recordset
'' indica se si è in fase di modifica
Dim modifica As Boolean

'' Ricarica le combo e la scheda
Private Sub Form_Activate()

    Me.ZOrder
    
    Call RicaricaComboBox("ASL", "NOME", cboAslAppartenenza)
    Call RicaricaComboBox("SIGLE_PROVINCIE", "NOME", cboProv)

    ' carica i dati
    Set rsDataset = New Recordset
    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        txtRagione = rsDataset("RAGIONE_SOCIALE")
        txtTipo = rsDataset("TIPO")
        txtIndirizzo = rsDataset("INDIRIZZO")
        txtCap = rsDataset("CAP")
        txtCitta = rsDataset("CITTA")
        cboProv.Text = rsDataset("PROV") & ""
        txtTelefono = rsDataset("TELEFONO")
        txtFax = rsDataset("FAX")
        txtCodiceFiscale = rsDataset("CODICE_FISCALE")
        txtIva = rsDataset("IVA")
        txtSitoWeb = rsDataset("SITO_WEB") & ""
        txtMail = rsDataset("MAIL")
        txtCodiceSts = rsDataset("CODICE_STS") & ""
        chkLogo.Value = IIf(CBool(rsDataset("LOGO")), Checked, Unchecked)
        chkLogoQualita.Value = IIf(CBool(rsDataset("LOGO_QUALITA")), Checked, Unchecked)
        chkLogoAziendale.Value = IIf(CBool(rsDataset("LOGO_AZIENDALE")), Checked, Unchecked)
        txtLogoISO = rsDataset("NOME_LOGOISO")
        txtLogoQualita = rsDataset("NOME_LOGOQUALITA")
        txtLogoAziendale = rsDataset("NOME_LOGOAZIENDALE")
        cboAslAppartenenza.ListIndex = GetCboListIndex(rsDataset("CODICE_ASL"), cboAslAppartenenza)
        cboDistrettoAppartenenza.ListIndex = GetCboListIndex(rsDataset("CODICE_DISTRETTO"), cboDistrettoAppartenenza)
        modifica = True
    Else
        modifica = False
    End If
    Set rsDataset = Nothing
End Sub

'' Verifica prima di salvare se tutti i dati sono stati inseriti
Private Function Completo() As Boolean
    Dim nome As String
    Completo = False
    If txtRagione = "" Then
        nome = "RAGIONE SOCIALE"
    ElseIf txtTipo = "" Then
        nome = "TIPO"
    ElseIf txtIndirizzo = "" Then
        nome = "INDIRIZZO"
    ElseIf txtCap = "" Then
        nome = "CAP"
    ElseIf txtCitta = "" Then
        nome = "COMUNE"
    ElseIf cboProv.Text = "" Then
        nome = "PROVINCIA"
    ElseIf txtCodiceFiscale = "" Then
        nome = "CODICE FISCALE"
    ElseIf txtIva = "" Then
        nome = "PARTITA IVA"
    ElseIf cboAslAppartenenza.ListIndex = -1 Then
        nome = "ASL APPARTENENZA"
    ElseIf cboDistrettoAppartenenza.ListIndex = -1 Then
        nome = "DISTRETTO APPARTENENZA"
    Else
        Completo = True
        Exit Function
    End If
    MsgBox "Inserire i dati obbligatori" & vbCrLf & "Campo: " & nome, vbCritical, "Attenzione"
End Function

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub ControllaImg()
    Set rsDataset = New Recordset

    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        If txtLogoISO <> rsDataset("NOME_LOGOISO") Then
            If Dir(structApri.pathExe & "\" & rsDataset("NOME_LOGOISO")) <> "" Then
                Kill structApri.pathExe & "\" & rsDataset("NOME_LOGOISO")
            End If
        End If
        If txtLogoQualita <> rsDataset("NOME_LOGOQUALITA") Then
            If Dir(structApri.pathExe & "\" & rsDataset("NOME_LOGOQUALITA")) <> "" Then
                Kill structApri.pathExe & "\" & rsDataset("NOME_LOGOQUALITA")
            End If
        End If
        If txtLogoAziendale <> rsDataset("NOME_LOGOAZIENDALE") Then
            If Dir(structApri.pathExe & "\" & rsDataset("NOME_LOGOAZIENDALE")) <> "" Then
                Kill structApri.pathExe & "\" & rsDataset("NOME_LOGOAZIENDALE")
            End If
        End If
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_nome() As Variant
    
    If Completo Then
        Call ControllaImg
        v_nome = Array("KEY", "RAGIONE_SOCIALE", "TIPO", "INDIRIZZO", "CAP", "CITTA", "PROV", "TELEFONO", "FAX", "IVA", "CODICE_FISCALE", "SITO_WEB", "MAIL", "LOGO", "LOGO_AZIENDALE", "LOGO_QUALITA", "NOME_LOGOISO", "NOME_LOGOQUALITA", "NOME_LOGOAZIENDALE", "CODICE_STS", "CODICE_ASL", "CODICE_DISTRETTO")
        v_Val = Array(1, txtRagione, txtTipo, txtIndirizzo, txtCap, txtCitta, cboProv.Text, txtTelefono, txtFax, txtIva, txtCodiceFiscale, txtSitoWeb, txtMail, IIf(chkLogo.Value = Checked, True, False), IIf(chkLogoAziendale.Value = Checked, True, False), IIf(chkLogoQualita.Value = Checked, True, False), txtLogoISO, txtLogoQualita, txtLogoAziendale, txtCodiceSts, cboAslAppartenenza.ItemData(cboAslAppartenenza.ListIndex), cboDistrettoAppartenenza.ItemData(cboDistrettoAppartenenza.ListIndex))
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
        MsgBox "I dati sono stati memorizzati nell'archivio", vbInformation, "Informazioni"
    End If
End Sub

Private Sub cmdScegli_Click(Index As Integer)
    On Error GoTo gestione
    
    Dim nomePathFile As String
    Dim nome As String
    Dim txt As TextBox
    
    Select Case Index
        Case 0
            Set txt = txtLogoISO
        Case 1
            Set txt = txtLogoQualita
        Case 2
            Set txt = txtLogoAziendale
    End Select
    
    cdlApri.CancelError = True
    cdlApri.Filter = "File immagine jpg|*.jpg"
    cdlApri.FilterIndex = 1
    cdlApri.ShowOpen
    nomePathFile = cdlApri.FileName
    nome = cdlApri.FileTitle
    
    If Len(nome) > 25 Then
        MsgBox "La lunghezza del nome del file non deve superare i 25 caratteri", vbCritical, "Attenzione"
    Else
        FileCopy nomePathFile, structApri.pathExe & "\" & nome
        txt = nome
    End If

    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number = 70 Then
        MsgBox "SELEZIONE NON PERMESSA dalla cartella di installazione del software", vbCritical, "ATTENZIONE!!!"
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub cboAslAppartenenza_Click()
    Call RicaricaComboBox("SELECT * FROM DISTRETTI WHERE CODICE_ASL=" & cboAslAppartenenza.ItemData(cboAslAppartenenza.ListIndex), "NOME", cboDistrettoAppartenenza)
End Sub

Private Sub txtRagione_GotFocus()
    txtRagione.BackColor = colArancione
End Sub

Private Sub txtRagione_LostFocus()
    txtRagione.BackColor = vbWhite
End Sub

Private Sub txtSitoWeb_GotFocus()
    txtSitoWeb.BackColor = colArancione
End Sub

Private Sub txtSitoWeb_LostFocus()
    txtSitoWeb.BackColor = vbWhite
End Sub

Private Sub TXTTIPO_GotFocus()
    txtTipo.BackColor = colArancione
End Sub

Private Sub TXTTIPO_LostFocus()
    txtTipo.BackColor = vbWhite
End Sub

Private Sub txtCap_GotFocus()
    txtCap.BackColor = colArancione
End Sub

Private Sub txtCap_LostFocus()
    txtCap.BackColor = vbWhite
End Sub

Private Sub txtCitta_GotFocus()
    txtCitta.BackColor = colArancione
End Sub

Private Sub txtCitta_LostFocus()
    txtCitta.BackColor = vbWhite
End Sub

Private Sub txtCodiceFiscale_GotFocus()
    txtCodiceFiscale.BackColor = colArancione
End Sub

Private Sub txtCodiceFiscale_LostFocus()
    txtCodiceFiscale.BackColor = vbWhite
End Sub

Private Sub txtCodiceSts_GotFocus()
    txtCodiceSts.BackColor = colArancione
End Sub

Private Sub txtCodiceSts_LostFocus()
    txtCodiceSts.BackColor = vbWhite
End Sub

Private Sub txtFax_GotFocus()
    txtFax.BackColor = colArancione
End Sub

Private Sub txtFax_LostFocus()
    txtFax.BackColor = vbWhite
End Sub

Private Sub txtLogoAziendale_GotFocus()
    txtLogoAziendale.BackColor = colArancione
End Sub

Private Sub txtLogoAziendale_LostFocus()
    txtLogoAziendale.BackColor = vbWhite
End Sub

Private Sub txtLogoISO_GotFocus()
    txtLogoISO.BackColor = colArancione
End Sub

Private Sub txtLogoISO_LostFocus()
    txtLogoISO.BackColor = vbWhite
End Sub

Private Sub txtLogoQualita_GotFocus()
    txtLogoQualita.BackColor = colArancione
End Sub

Private Sub txtLogoQualita_LostFocus()
    txtLogoQualita.BackColor = vbWhite
End Sub

Private Sub txtIndirizzo_GotFocus()
    txtIndirizzo.BackColor = colArancione
End Sub

Private Sub txtIndirizzo_LostFocus()
    txtIndirizzo.BackColor = vbWhite
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

Private Sub txtTelefono_GotFocus()
    txtTelefono.BackColor = colArancione
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.BackColor = vbWhite
End Sub
