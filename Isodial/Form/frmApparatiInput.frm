VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmApparatiInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scheda Apparato"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   6620
         Picture         =   "frmApparatiInput.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1720
         Width           =   375
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   1320
         Picture         =   "frmApparatiInput.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1730
         Width           =   375
      End
      Begin VB.ComboBox cboTipoRene 
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
         ItemData        =   "frmApparatiInput.frx":08B2
         Left            =   3860
         List            =   "frmApparatiInput.frx":08BF
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1200
      End
      Begin VB.TextBox txtpostazione 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox chkAttivaAlert 
         Caption         =   "ATTIVA ALERT su MANUTENZIONI ORDINARIE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   390
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin VB.ComboBox cboModello 
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
         Index           =   2
         Left            =   6980
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cboModalitaAcquisizione 
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
         Left            =   7920
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtPeriodoAmmortamento 
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
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   13
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtNoteCollaudo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   4440
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   3720
         Width           =   5895
      End
      Begin VB.TextBox txtMatricola 
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
         Left            =   6960
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtNumeroApparato 
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
         Left            =   4410
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cboTipoApparato 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtNumeroInventario 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin DataTimeBox.uDataTimeBox oDataDismissione 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   15
         Top             =   3720
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataFabbricazione 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   12
         Top             =   2760
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataAcquisizione 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   10
         Top             =   2280
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataCollaudo 
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   14
         Top             =   3240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frequenza Manutenzione Ordinaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   120
         TabIndex        =   37
         Top             =   4680
         Width           =   3975
         Begin VB.ComboBox cboSicurezza 
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
            Height          =   315
            ItemData        =   "frmApparatiInput.frx":08DA
            Left            =   1560
            List            =   "frmApparatiInput.frx":08F6
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   960
            Width           =   2295
         End
         Begin VB.ComboBox cboFunzionalita 
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
            Height          =   315
            ItemData        =   "frmApparatiInput.frx":0956
            Left            =   1560
            List            =   "frmApparatiInput.frx":0972
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Funzionalità"
            BeginProperty Font 
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
            TabIndex        =   39
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sicurezza"
            BeginProperty Font 
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
            TabIndex        =   38
            Top             =   960
            Width           =   1020
         End
      End
      Begin DataTimeBox.uDataTimeBox oDataRottamazione 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Top             =   4200
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label lblManutentore 
         BackColor       =   &H8000000E&
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
         Height          =   315
         Left            =   7080
         TabIndex        =   44
         Top             =   1800
         Width           =   3250
      End
      Begin VB.Label lblProduttore 
         BackColor       =   &H8000000E&
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
         Height          =   315
         Left            =   1780
         TabIndex        =   43
         Top             =   1800
         Width           =   3250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Rene"
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
         Index           =   17
         Left            =   2680
         TabIndex        =   42
         Top             =   1340
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postazione"
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
         Index           =   16
         Left            =   120
         TabIndex        =   41
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Rottamazione"
         BeginProperty Font 
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
         Top             =   4250
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Ammortamento (anni)"
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
         Index           =   18
         Left            =   5280
         TabIndex        =   36
         Top             =   2790
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Collaudo"
         BeginProperty Font 
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
         TabIndex        =   34
         Top             =   3270
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Acquisizione"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   33
         Top             =   2310
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Fabbricazione"
         BeginProperty Font 
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
         TabIndex        =   32
         Top             =   2790
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Manutentore"
         BeginProperty Font 
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
         TabIndex        =   31
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produttore"
         BeginProperty Font 
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
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modello"
         BeginProperty Font 
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
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modalità di Acquisizione"
         BeginProperty Font 
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
         Left            =   5280
         TabIndex        =   28
         Top             =   2310
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note Collaudo"
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
         Index           =   3
         Left            =   6700
         TabIndex        =   27
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Matricola"
         BeginProperty Font 
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
         Left            =   5280
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Apparato/Rene"
         BeginProperty Font 
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
         Left            =   2450
         TabIndex        =   25
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Dismissione"
         BeginProperty Font 
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
         TabIndex        =   24
         Top             =   3750
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria Apparato"
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
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   750
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Inventario"
         BeginProperty Font 
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
         TabIndex        =   22
         Top             =   380
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   6120
      Width           =   10455
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
         Height          =   600
         Left            =   7440
         TabIndex        =   20
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
         Height          =   600
         Left            =   9120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmApparatiInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsNumeroProgressivo As Recordset
Dim rsMemorizzaApparecchiature As Recordset
Dim rsCercaApparato As Recordset
Dim rsPresenzaManutenzione As Recordset
Dim NumeroApparato As Integer
Dim ModificaApparato As Boolean
Dim ProxRevFun As String
Dim ProxRevSic As String
Dim mDataCollaudo As Variant
Dim mProxRevFun As Variant
Dim mProxRevSic As Variant
Dim PostazionePrec As String
Dim cboTipoApparatoPrec As String
Dim cboTipoRenePrec As String
Dim MantieniDato As Integer
Dim KeyProduttore As Integer
Dim keyManutentore As Integer

Private Sub cboModalitaAcquisizione_GotFocus(Index As Integer)
    cboModalitaAcquisizione(1).BackColor = colArancione
End Sub

Private Sub cboModalitaAcquisizione_LostFocus(Index As Integer)

    If Len(cboModalitaAcquisizione(1)) > 15 Then
        MsgBox "NON è possibile memorizzare più di 15 caratteri", vbCritical, "ATTENZIONE!!!"
        cboModalitaAcquisizione(1).Text = ""
        cboModalitaAcquisizione(1).SetFocus
        Exit Sub
    End If
    
    If cboModalitaAcquisizione(1).Text <> "" Then
        Call GestisciNuovo("APPARATI_MOD_ACQ", cboModalitaAcquisizione(1))
    End If

    cboModalitaAcquisizione(1).BackColor = vbWhite
    
End Sub

Private Sub cboModello_GotFocus(Index As Integer)
    cboModello(2).BackColor = colArancione
End Sub

Private Sub cboModello_LostFocus(Index As Integer)
    
    If Len(cboModello(2)) > 30 Then
        MsgBox "NON è possibile memorizzare più di 30 caratteri", vbCritical, "ATTENZIONE!!!"
        cboModello(2).Text = ""
        cboModello(2).SetFocus
        Exit Sub
    End If
    
    If cboModello(2).Text <> "" Then
        Call GestisciNuovo("APPARATI_MODELLO", cboModello(2))
    End If
    
    cboModello(2).BackColor = vbWhite
    
End Sub

Private Sub cboTipoApparato_GotFocus(Index As Integer)
        cboTipoApparato(0).BackColor = colArancione
End Sub

Private Sub cboTipoApparato_LostFocus(Index As Integer)

    If Len(cboTipoApparato(0)) > 30 Then
        MsgBox "NON è possibile memorizzare oltre 30 caratteri", vbCritical, "ATTENZIONE!!!"
        cboTipoApparato(0).Text = ""
        cboTipoApparato(0).SetFocus
        Exit Sub
    ElseIf cboTipoApparato(0).Text = "RENE ARTIFICIALE" Then
        Label1(16).Enabled = True
        Label1(17).Enabled = True
        txtpostazione.Enabled = True
        cboTipoRene.Enabled = True
        cboTipoRene.ListIndex = 0
    ElseIf cboTipoApparato(0).Text <> "RENE ARTIFICIALE" Then
        Label1(16).Enabled = False
        Label1(17).Enabled = False
        txtpostazione.Enabled = False
        cboTipoRene.Enabled = False
        cboTipoRene.ListIndex = -1
        txtpostazione = ""
    ElseIf cboTipoApparato(0).Text <> "" Then
        Call GestisciNuovo("APPARATI_TIPO", cboTipoApparato(0))
    End If

    cboTipoApparato(0).BackColor = vbWhite
    
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    'Salvo la key dell' apparato per evitare che si perda
    'quando carico il frmTrova
    MantieniDato = tTrova.keyReturn
    
    If Index = 1 Then
        ModificaProduttore = True
        tTrova.Tipo = tpPRODUTTORE_MANUTENTORE
        tTrova.condizione = ""
        tTrova.condStato = ""
        Unload frmTrova
        frmTrova.Show 1
        KeyProduttore = tTrova.keyReturn
        lblProduttore.Caption = tTrova.NomeStriga
        ModificaProduttore = False
    Else
        ModificaManutentore = True
        tTrova.Tipo = tpPRODUTTORE_MANUTENTORE
        tTrova.condizione = ""
        tTrova.condStato = ""
        Unload frmTrova
        frmTrova.Show 1
        keyManutentore = tTrova.keyReturn
        lblManutentore.Caption = tTrova.NomeStriga
        ModificaManutentore = False
    End If
    
    tTrova.keyReturn = MantieniDato
End Sub

Private Sub oDataCollaudo_LostFocus(Index As Integer)
    If cboTipoApparato(0).Text = "RENE ARTIFICIALE" And oDataCollaudo(3).data <> "" And oDataRottamazione(0).data = "" Then
       oDataRottamazione(0).data = DateAdd("yyyy", 8, oDataCollaudo(3).data)
    End If
End Sub

Private Sub cmdChiudi_Click()
    If MantieniKeyReturn > 0 Then
        ModificaApparato = False
        Unload frmApparatiInput
    Else
        MantieniKeyReturn = -2
        ModificaApparato = False
        Unload frmApparatiInput
    End If
End Sub

Private Sub cmdMemorizza_Click()
Dim v_Nomi() As Variant
Dim v_Val() As Variant
'Dim numKey As Integer
Dim valore As Integer

    '' Controlli sui campi
    If NumInvent Then
        Exit Sub
    ElseIf NumApp Then
        Exit Sub
    ElseIf txtNumeroInventario.Text = "" Then
        MsgBox "Inserire il N° di Inventario", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf txtNumeroApparato = "" Then
        MsgBox "Inserire il Numero dell'Apparato o del Rene Artificiale", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf cboTipoApparato(0).Text = "" Then
        MsgBox "Inserire la Categoria a cui appartiene l'Apparato", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf (IsPossibleDelete("TURNI", "CODICE_RENE", KeyApparato) = False Or IsPossibleDelete("STORICO_DIALISI_GIORNALIERA", "CODICE_RENE", KeyApparato) = False) And cboTipoApparato(0) <> "RENE ARTIFICIALE" And tTrova.keyReturn <> 0 Then
        MsgBox "MODIFICA CATEGORIA APPARATO NON PERMESSA!!! - Dati in relazione con altre gestioni dell'applicativo", vbInformation, "ATTENZIONE!!!"
        cboTipoApparato(0) = cboTipoApparatoPrec
        txtpostazione = PostazionePrec
        cboTipoRene.Text = cboTipoRenePrec
        Label1(16).Enabled = True
        Label1(17).Enabled = True
        txtpostazione.Enabled = True
        cboTipoRene.Enabled = True
        Exit Sub
    ElseIf cboModello(2).Text = "" Then
        MsgBox "Inserire il Modello", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf txtpostazione = "" And cboTipoApparato(0) = "RENE ARTIFICIALE" Then
        MsgBox "Inserire la Postazione del Rene", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf txtMatricola.Text = "" Then
        MsgBox "Inserire la Matricola", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf lblProduttore.Caption = "" Then
        MsgBox "Inserire il Produttore", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf oDataAcquisizione(2).txtBox = "" Then
        MsgBox "Inserire la Data di Acquisizione", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf cboModalitaAcquisizione(1).Text = "" Then
        MsgBox "Inserire la Modalità di Acquisizione", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf oDataRottamazione(0).txtBox = "" Then
        MsgBox "Inserire la Data di Rottamazione", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf cboFunzionalita.ListIndex = -1 Then
        MsgBox "Inserire la Frequenza per la Manutenzione Ordinaria della FUNZIONALITA'", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf cboSicurezza.ListIndex = -1 Then
        MsgBox "Inserire la Frequenza per la Manutenzione Ordinaria della SICUREZZA", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf CDate(oDataAcquisizione(2).data) > date Then
        MsgBox "Data di Acquisizione successiva alla Data Odierna", vbCritical, "ATTENZIONE!!!"
        Exit Sub
    ElseIf oDataFabbricazione(0).txtBox <> "" Then
        If CDate(oDataFabbricazione(0).data) > date Then
            MsgBox "Data di Fabbricazione successiva alla Data Odierna", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        End If
    ElseIf txtPeriodoAmmortamento = "" Then
        txtPeriodoAmmortamento = 0
    ElseIf NumPost Then
        txtpostazione = PostazionePrec
        Exit Sub
    End If
    
    'controlla campo per campo se la data digitata è superiore alla data di sistema oppure
    
    'se il campo contiene la data controlla se si può modificare
    If oDataCollaudo(3).txtBox <> "" Then
        If CDate(oDataCollaudo(3).data) > date Then
            MsgBox "Data di Collaudo successiva alla Data Odierna", vbCritical, "ATTENZIONE!!!"
            Exit Sub
        ElseIf oDataCollaudo(3).data <> mDataCollaudo And ModificaApparato And PresenzaManutenzioneOrdinaria Then
            MsgBox "NON E' POSSIBILE MODIFICARE LA DATA DI COLLAUDO - Presenza di schede di Manutenzione Ordinaria ", vbCritical, "ATTENZIONE!!!"
            oDataCollaudo(3).data = mDataCollaudo
            Exit Sub
        End If
    'se il campo NON contiene la data controlla se si può modificare
    ElseIf ModificaApparato And PresenzaManutenzioneOrdinaria Then
        MsgBox "NON E' POSSIBILE MODIFICARE LA DATA DI COLLAUDO - Presenza di schede di Manutenzione Ordinaria ", vbCritical, "ATTENZIONE!!!"
        oDataCollaudo(3).data = mDataCollaudo
        Exit Sub
    End If
    
    If oDataDismissione(1).txtBox <> "" Then
       If CDate(oDataDismissione(1).data) > date Then
          MsgBox "Data di Dismissione successiva alla Data Odierna", vbCritical, "ATTENZIONE!!!"
          Exit Sub
      End If
    End If
        
    Call SuperUcase(Me)
        
    Set rsMemorizzaApparecchiature = New Recordset
 
    Call CalcoloProxRevFun
    Call CalcoloProxRevSic
    
    If ModificaApparato = True Then
        numKey = NumeroApparato
    Else
        numKey = GetNumero("APPARATI")
    End If
    
    If cboTipoRene = "HCV POS" Then
        valore = 1
    ElseIf cboTipoRene = "HBV POS" Then
        valore = 2
    Else
        valore = 0
    End If

    v_Nomi = Array("KEY", "NUMERO_INVENTARIO", "NUMERO_APPARATO", "TIPO_APPARATO", "MODELLO", "POSTAZIONE", "TIPO", "MATRICOLA", "PRODUTTORE", "MANUTENTORE", "DATA_FABBRICAZIONE" _
                    , "DATA_COLLAUDO", "NOTE_COLLAUDO", "DATA_DISMISSIONE", "MODALITA_ACQUISIZIONE", "DATA_ACQUISIZIONE", "DATA_ROTTAMAZIONE", "PERIODO_AMMORTAMENTO" _
                    , "FUNZIONALITA", "SICUREZZA", "PROXREVFUN", "PROXREVSIC", "ALERT", "KEY_PRODUTTORE", "KEY_MANUTENTORE")
                    
        
    v_Val = Array(numKey, txtNumeroInventario, txtNumeroApparato, cboTipoApparato(0).Text, cboModello(2).Text, UCase(txtpostazione), valore, txtMatricola, lblProduttore.Caption, lblManutentore.Caption, IIf(oDataFabbricazione(0).data = "", Null, oDataFabbricazione(0).data) _
                    , IIf(oDataCollaudo(3).data = "", Null, oDataCollaudo(3).data), txtNoteCollaudo, IIf(oDataDismissione(1).data = "", Null, oDataDismissione(1).data), cboModalitaAcquisizione(1).Text, IIf(oDataAcquisizione(2).data = "", Null, oDataAcquisizione(2).data), IIf(oDataRottamazione(0).data = "", Null, oDataRottamazione(0).data), txtPeriodoAmmortamento _
                    , cboFunzionalita.ListIndex, cboSicurezza.ListIndex, IIf(ProxRevFun = "", Null, ProxRevFun), IIf(ProxRevSic = "", Null, ProxRevSic), IIf(chkAttivaAlert.Value = Checked, True, False), KeyProduttore, keyManutentore)

    If ModificaApparato = True Then
        rsMemorizzaApparecchiature.Open "SELECT * FROM APPARATI WHERE KEY=" & NumeroApparato, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsMemorizzaApparecchiature.Update v_Nomi, v_Val
    Else
        rsMemorizzaApparecchiature.Open "APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsMemorizzaApparecchiature.AddNew v_Nomi, v_Val
    End If

    Set rsMemorizzaApparecchiature = Nothing

    Call Pulisci

    txtNumeroInventario = GetNumero("APPARATI")
        
    If ModificaApparato = True Then
        ModificaApparato = False
    Else
        ModificaApparato = False
    End If
    Unload frmApparatiInput
End Sub

'' Calcolo per la prossima revisione funzionale
Private Sub CalcoloProxRevFun()
    If oDataCollaudo(3).data <> "" And IsPossibleDelete("MANUTENZIONE_APPARATI", "CODICE_APPARATO", KeyApparato) Then
        Select Case cboFunzionalita.ListIndex
            Case Is = 0
                ' funzione per sommare la date
                ' d=day, m=month, y=year
                ProxRevFun = DateAdd("m", 1, oDataCollaudo(3).data)
            Case Is = 1
                ProxRevFun = DateAdd("m", 2, oDataCollaudo(3).data)
            Case Is = 2
                ProxRevFun = DateAdd("m", 3, oDataCollaudo(3).data)
            Case Is = 3
                ProxRevFun = DateAdd("m", 4, oDataCollaudo(3).data)
            Case Is = 4
                ProxRevFun = DateAdd("m", 6, oDataCollaudo(3).data)
            Case Is = 5
                ' calcolo l' aggiunta dell' anno con la somma dei mesi
                ' in quanto la funzione "year" aggiunge il giorno
                ProxRevFun = DateAdd("m", 12, oDataCollaudo(3).data)
            Case Is = 6
                ProxRevFun = DateAdd("m", 24, oDataCollaudo(3).data)
            Case Is = 7
                ProxRevFun = DateAdd("m", 36, oDataCollaudo(3).data)
        End Select
    Else
      ' se ci sono schede di manutenzione non modifica la data della prossima revisione funz.
        ProxRevFun = mProxRevFun
    End If
End Sub

'' Calcolo per la prossima revisione Sicurezza
Private Sub CalcoloProxRevSic()
    If oDataCollaudo(3).data <> "" And IsPossibleDelete("MANUTENZIONE_APPARATI", "CODICE_APPARATO", KeyApparato) Then
        Select Case cboSicurezza.ListIndex
            Case Is = 0
                ' funzione per sommare la date
                ' d=day, m=month, y=year
                ProxRevSic = DateAdd("m", 1, oDataCollaudo(3).data)
            Case Is = 1
                ProxRevSic = DateAdd("m", 2, oDataCollaudo(3).data)
            Case Is = 2
                ProxRevSic = DateAdd("m", 3, oDataCollaudo(3).data)
            Case Is = 3
                ProxRevSic = DateAdd("m", 4, oDataCollaudo(3).data)
            Case Is = 4
                ProxRevSic = DateAdd("m", 6, oDataCollaudo(3).data)
            Case Is = 5
                ' calcolo l' aggiunta dell' anno con la somma dei mesi
                ' in quanto la funzione "year" aggiunge il giorno
                ProxRevSic = DateAdd("m", 12, oDataCollaudo(3).data)
            Case Is = 6
                ProxRevSic = DateAdd("m", 24, oDataCollaudo(3).data)
            Case Is = 7
                ProxRevSic = DateAdd("m", 36, oDataCollaudo(3).data)
        End Select
    Else
      ' se ci sono schede di manutenzione non modifica la data della prossima revisione sic.
        ProxRevSic = mProxRevSic
    End If
End Sub

'' Controllo sulla presenza di schede di manuntenzione ordinaria
Private Function PresenzaManutenzioneOrdinaria() As Boolean

    Set rsPresenzaManutenzione = New Recordset
    rsPresenzaManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE CODICE_APPARATO= " & KeyApparato & " ORDER BY KEY DESC ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not (rsPresenzaManutenzione.EOF And rsPresenzaManutenzione.BOF) Then
        Do While Not rsPresenzaManutenzione.EOF
            If rsPresenzaManutenzione("TIPO_MANUTENZIONE") = "ORD. FUNZ." Or rsPresenzaManutenzione("TIPO_MANUTENZIONE") = "ORD. SICUR." Or rsPresenzaManutenzione("TIPO_MANUTENZIONE") = "ORD. FUN. SIC." Then
                PresenzaManutenzioneOrdinaria = True
                Exit Function
            End If
            rsPresenzaManutenzione.MoveNext
        Loop
    End If
    
    Set rsPresenzaManutenzione = Nothing
    
End Function

'' Controllo sull'univocità del numero d'inventario
Private Function NumInvent() As Boolean
    Dim massimo As Integer
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT MAX(NUMERO_INVENTARIO) AS MASSIMO FROM APPARATI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If IsNull(rsDataset("MASSIMO")) Then
        massimo = 0
    Else
        massimo = rsDataset("MASSIMO")
    End If
    rsDataset.Close
    
    If txtNumeroInventario = "" Then
        txtNumeroInventario = massimo + 1
        NumInvent = False
    Else
        rsDataset.Open "SELECT KEY,NUMERO_INVENTARIO FROM APPARATI WHERE NUMERO_INVENTARIO=" & txtNumeroInventario, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If rsDataset("KEY") <> tTrova.keyReturn Then
                If MsgBox("Numero Inventario già in uso." & vbCrLf & "Si preferisce assegnare il valore " & massimo + 1 & " scelto dal sistema?", vbCritical + vbYesNo, "ATTENZIONE!!!!!!") = vbYes Then
                    txtNumeroInventario = massimo + 1
                    NumInvent = False
                Else
                    txtNumeroInventario.SetFocus
                    NumInvent = True
                End If
            Else
                NumInvent = False
            End If
        Else
            NumInvent = False
        End If
    End If
    
    Set rsDataset = Nothing

End Function

'' Controllo sull'esistenza della postazione
Private Function NumPost() As Boolean
   Dim rsDataset As New Recordset
   
   rsDataset.Open "SELECT KEY,POSTAZIONE FROM APPARATI WHERE POSTAZIONE ='" & txtpostazione & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

   If Not (rsDataset.EOF And rsDataset.BOF) Then
    If rsDataset("KEY") <> tTrova.keyReturn And cboTipoApparato(0) = "RENE ARTIFICIALE" Then
        If MsgBox("Postazione già presente." & vbCrLf & "Vuoi duplicarla?", vbQuestion + vbYesNo + vbDefaultButton2, "ATTENZIONE!!!") = vbYes Then
            NumPost = False
        Else
            NumPost = True
        End If
    Else
        NumPost = False
    End If
   End If
   
   rsDataset.Close
   Set rsDataset = Nothing
End Function

'' Controllo sull'univocità del n° di apparato
Private Function NumApp() As Boolean
   If txtNumeroApparato = "" Then
    Exit Function
   End If
   
   Dim rsDataset As New Recordset
   
   rsDataset.Open "SELECT KEY,NUMERO_APPARATO FROM APPARATI WHERE NUMERO_APPARATO =" & txtNumeroApparato, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

   If Not (rsDataset.EOF And rsDataset.BOF) Then
    If rsDataset("KEY") <> tTrova.keyReturn Then
        MsgBox "Numero di Apparato/Rene già esistente", vbCritical, "ATTENZIONE!!!"
        NumApp = True
    End If
   Else
    NumApp = False
   End If
   
   rsDataset.Close
   Set rsDataset = Nothing
End Function

Private Sub Pulisci()
    txtNumeroApparato.Text = ""
    cboTipoApparato(0).Text = ""
    txtpostazione = ""
    cboTipoRene.ListIndex = 0
    cboModello(2).Text = ""
    txtMatricola.Text = ""
    lblProduttore.Caption = ""
    lblManutentore.Caption = ""
    oDataFabbricazione(0).Pulisci
    oDataDismissione(1).Pulisci
    cboModalitaAcquisizione(1).Text = ""
    oDataAcquisizione(2).Pulisci
    oDataCollaudo(3).Pulisci
    oDataRottamazione(0).Pulisci
    txtNoteCollaudo.Text = ""
    txtPeriodoAmmortamento.Text = ""
    txtNumeroInventario.SetFocus
    txtNumeroInventario_GotFocus
    NumeroApparato = 0
    tTrova.keyReturn = 0
    MantieniDato = 0
    KeyProduttore = 0
    keyManutentore = 0
    cboFunzionalita.ListIndex = -1
    cboSicurezza.ListIndex = -1
    ProxRevFun = ""
    ProxRevSic = ""
End Sub

Private Sub Form_Activate()
    Call RicaricaComboBox("APPARATI_TIPO", "NOME", cboTipoApparato(0))
    Call RicaricaComboBox("APPARATI_MODELLO", "NOME", cboModello(2))
    Call RicaricaComboBox("APPARATI_MOD_ACQ", "NOME", cboModalitaAcquisizione(1))
End Sub

Private Sub Form_Load()

    If tTrova.keyReturn = 0 And tInput.mantieniDati = True Then  'predispone il form all'inserimento del rene
        Label1(16).Enabled = True                                'in rottamazione da sostituire
        Label1(17).Enabled = True
        txtpostazione.Enabled = True
        cboTipoRene.Enabled = True
        txtNumeroInventario = GetNumero("APPARATI")
        cboTipoApparato(0) = "RENE ARTIFICIALE"
        txtpostazione = tInput.v_valori(1)
        cboTipoRene.ListIndex = 0

    ElseIf tTrova.keyReturn = 0 Then
        txtNumeroInventario = GetNumero("APPARATI")
    Else
        NumeroApparato = tTrova.keyReturn
        Call CaricaApparato
    End If
End Sub

Private Sub CaricaApparato()
    
    Set rsCercaApparato = New Recordset
    
    rsCercaApparato.Open "SELECT * FROM APPARATI WHERE KEY =" & NumeroApparato, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    txtNumeroInventario.Text = rsCercaApparato("NUMERO_INVENTARIO")
    txtNumeroApparato.Text = rsCercaApparato("NUMERO_APPARATO")
    cboTipoApparato(0).Text = rsCercaApparato("TIPO_APPARATO")
    cboTipoApparatoPrec = rsCercaApparato("TIPO_APPARATO")
    cboModello(2).Text = rsCercaApparato("MODELLO")
    txtpostazione.Text = rsCercaApparato("POSTAZIONE")
    PostazionePrec = rsCercaApparato("POSTAZIONE")
    
    If rsCercaApparato("TIPO_APPARATO") = "RENE ARTIFICIALE" Then
        Label1(16).Enabled = True
        Label1(17).Enabled = True
        txtpostazione.Enabled = True
        cboTipoRene.Enabled = True
        If rsCercaApparato("TIPO") = 0 Then
            cboTipoRene.Text = "NEG"
        ElseIf rsCercaApparato("TIPO") = 1 Then
            cboTipoRene.Text = "HCV POS"
        Else
            cboTipoRene.Text = "HBV POS"
        End If
        cboTipoRenePrec = cboTipoRene.Text
    End If
 
    txtMatricola.Text = rsCercaApparato("MATRICOLA")
    lblProduttore.Caption = rsCercaApparato("PRODUTTORE")
    lblManutentore.Caption = rsCercaApparato("MANUTENTORE")
    oDataFabbricazione(0).txtBox = rsCercaApparato("DATA_FABBRICAZIONE") & ""
    oDataDismissione(1).txtBox = rsCercaApparato("DATA_DISMISSIONE") & ""
    cboModalitaAcquisizione(1).Text = rsCercaApparato("MODALITA_ACQUISIZIONE")
    oDataAcquisizione(2).txtBox = rsCercaApparato("DATA_ACQUISIZIONE") & ""
    oDataCollaudo(3).txtBox = rsCercaApparato("DATA_COLLAUDO") & ""
    oDataRottamazione(0).txtBox = rsCercaApparato("DATA_ROTTAMAZIONE") & ""
    txtNoteCollaudo.Text = rsCercaApparato("NOTE_COLLAUDO")
    txtPeriodoAmmortamento.Text = rsCercaApparato("PERIODO_AMMORTAMENTO")
    cboFunzionalita.ListIndex = rsCercaApparato("FUNZIONALITA")
    cboSicurezza.ListIndex = rsCercaApparato("SICUREZZA")
    chkAttivaAlert.Value = IIf(CBool(rsCercaApparato("ALERT")), Checked, Unchecked)
        
    mDataCollaudo = oDataCollaudo(3).txtBox
    mProxRevFun = rsCercaApparato("PROXREVFUN")
    mProxRevSic = rsCercaApparato("PROXREVSIC")
    
    Set rsCercaApparato = Nothing
    ModificaApparato = True
    
End Sub

Private Sub txtMatricola_GotFocus()
    txtMatricola.BackColor = colArancione
End Sub

Private Sub txtMatricola_LostFocus()
    txtMatricola.BackColor = vbWhite
End Sub

Private Sub txtNoteCollaudo_GotFocus()
    txtNoteCollaudo.BackColor = colArancione
End Sub

Private Sub txtNoteCollaudo_LostFocus()
    txtNoteCollaudo.BackColor = vbWhite
End Sub

Private Sub txtNumeroApparato_GotFocus()
    txtNumeroApparato.BackColor = colArancione
End Sub

Private Sub txtNumeroApparato_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumeroApparato_LostFocus()
    txtNumeroApparato.BackColor = vbWhite
End Sub

Private Sub txtNumeroInventario_GotFocus()
    txtNumeroInventario.BackColor = colArancione
End Sub

Private Sub txtNumeroInventario_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumeroInventario_LostFocus()
    txtNumeroInventario.BackColor = vbWhite
End Sub

Private Sub txtPeriodoAmmortamento_GotFocus()
    txtPeriodoAmmortamento.BackColor = colArancione
End Sub

Private Sub txtPeriodoAmmortamento_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPeriodoAmmortamento_LostFocus()
    txtPeriodoAmmortamento.BackColor = vbWhite
End Sub
