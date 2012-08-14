VERSION 5.00
Object = "{EB7F7146-0A68-4457-8036-5793F0EB1EB8}#31.0#0"; "SuperTextBox.ocx"
Begin VB.Form frmParametriFattura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parametri Fattura"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Intestazione Fattura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   7815
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
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtAutorizzazione 
         BeginProperty Font 
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
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
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
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   3375
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3375
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1920
         Width           =   2295
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
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1440
         Width           =   735
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
         Left            =   6000
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   3375
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
         Index           =   10
         Left            =   5400
         TabIndex        =   42
         Top             =   1920
         Width           =   420
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
         TabIndex        =   40
         Top             =   2400
         Width           =   1935
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
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1080
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
         Index           =   13
         Left            =   120
         TabIndex        =   30
         Top             =   4470
         Width           =   1200
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
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   600
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
         Left            =   5280
         TabIndex        =   28
         Top             =   1440
         Width           =   555
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
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   480
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
         Left            =   5280
         TabIndex        =   26
         Top             =   990
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
         TabIndex        =   25
         Top             =   990
         Width           =   870
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Importi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   36
      Top             =   6120
      Width           =   7815
      Begin SuperTextBox.uSuperTextBox txtQuotaAggiuntiva 
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   480
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   6
      End
      Begin SuperTextBox.uSuperTextBox txtQuotaNazionale 
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         Top             =   480
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   6
      End
      Begin SuperTextBox.uSuperTextBox txtRimborsoSpeseViaggio 
         Height          =   285
         Left            =   2640
         TabIndex        =   19
         Top             =   960
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   6
      End
      Begin SuperTextBox.uSuperTextBox txtImportoTicket 
         Height          =   285
         Left            =   6000
         TabIndex        =   20
         Top             =   960
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   6
      End
      Begin SuperTextBox.uSuperTextBox txtImportoBollo 
         Height          =   285
         Left            =   2640
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
         _ExtentX        =   2143
         _ExtentY        =   503
         IsMultiLine     =   0   'False
         OnlyNumber      =   -1  'True
         IsPossibleSpacing=   0   'False
         IsDecimal       =   -1  'True
         MaxLenght       =   6
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
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Quota Nazionale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4080
         TabIndex        =   41
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Rimborso spese viaggi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Quota Regionale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ticket"
         BeginProperty Font 
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
         Left            =   4080
         TabIndex        =   37
         Top             =   960
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   7920
      Width           =   7815
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
         Left            =   4680
         TabIndex        =   22
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
         Left            =   6480
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coordinate Bonifico Bancario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   7815
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
         TabIndex        =   9
         Top             =   480
         Width           =   5775
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
         TabIndex        =   13
         Top             =   960
         Width           =   735
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
         TabIndex        =   14
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
         Index           =   3
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   15
         Top             =   960
         Width           =   1455
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
         TabIndex        =   10
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
         TabIndex        =   12
         Top             =   960
         Width           =   255
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
         Index           =   0
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   11
         Top             =   960
         Width           =   375
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
         Index           =   14
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1545
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
         Index           =   15
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dicitura da stampare in fattura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   35
      Top             =   4800
      Width           =   7815
      Begin VB.TextBox txtDicitura 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmParametriFattura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmParametriFattura.frm
'
' <b>Descrizione</b>: Scheda Parametri Fattura associata alla tab INTESTAZIONE_FATTURA
'
' @remarks
'
' @author
'
' @date 22/02/2011 18.34
Option Explicit

'' rs della scheda
Dim rsDataset As Recordset
Dim modifica As Boolean
Dim lettera As String

Private Sub Form_Activate()
    Me.ZOrder
End Sub

Private Sub Form_Load()
    Dim strIban As String
    Dim strSql As String
    
    Dim rsDatasetControllo As New Recordset
    
    Call RicaricaComboBox("ASL", "NOME", cboAsl)
    Call RicaricaComboBox("COMUNI", "NOME", cboComune)
    
    
    rsDatasetControllo.Open "SELECT CODICE_ASL FROM INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDatasetControllo("CODICE_ASL") = 5 Then
        Label1(11).Visible = True
        txtImportoBollo.Visible = True
    Else
        Label1(11).Visible = False
        txtImportoBollo.Visible = False
        Frame4.Top = 7440
        Frame5.Height = 1455
        frmParametriFattura.Height = 8730
    End If
    rsDatasetControllo.Close
    Set rsDatasetControllo = Nothing
    
    
    strSql = "SELECT    INTESTAZIONE_FATTURA.*, ASL.KEY AS ASLKEY, COMUNI.KEY AS COMUNIKEY " & _
            "FROM       (INTESTAZIONE_FATTURA " & _
            "           LEFT OUTER JOIN ASL ON ASL.KEY=INTESTAZIONE_FATTURA.CODICE_ASL) " & _
            "           LEFT OUTER JOIN COMUNI ON COMUNI.KEY=INTESTAZIONE_FATTURA.CODICE_COMUNE"
    Set rsDataset = New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        cboAsl.ListIndex = GetCboListIndex(rsDataset("ASLKEY"), cboAsl)
        txtIndirizzo = rsDataset("INDIRIZZO")
        txtCap = rsDataset("CAP")
        cboComune.ListIndex = GetCboListIndex(rsDataset("COMUNIKEY"), cboComune)
        txtProv = rsDataset("PROV")
        txtIva = rsDataset("P_IVA")
        txtCodiceFiscale = rsDataset("CODICE_FISCALE")
        txtAutorizzazione = rsDataset("NUMERO_AUTORIZZAZIONE")
        txtDicitura = rsDataset("DICITURA")
        txtImportoTicket.Text = rsDataset("TICKET")
        txtQuotaAggiuntiva.Text = rsDataset("QUOTA_AGGIUNTIVA")
        txtQuotaNazionale.Text = rsDataset("QUOTA_NAZIONALE")
        txtIntestatario = rsDataset("INTESTATARIO_CC")
        txtRimborsoSpeseViaggio.Text = rsDataset("RIMBORSO_SPESE_VIAGGIO")
        txtImportoBollo.Text = rsDataset("IMPORTO_BOLLO")
        strIban = rsDataset("IBAN")
        txtIbanAlfa(0) = Mid(strIban, 1, 2)
        txtIbanNum(0) = Mid(strIban, 3, 2)
        txtIbanAlfa(1) = Mid(strIban, 5, 1)
        txtIbanNum(1) = Mid(strIban, 6, 5)
        txtIbanNum(2) = Mid(strIban, 11, 5)
        txtIbanNum(3) = Mid(strIban, 16, 12)
        modifica = True
    Else
        modifica = False
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Function Completo() As Boolean
    Dim nome As String
    Completo = False
    If cboAsl.ListIndex = -1 Then
        nome = "ASL A CUI FATTURARE"
    ElseIf txtIndirizzo = "" Then
        nome = "INDIRIZZO"
    ElseIf txtCap = "" Then
        nome = "CAP"
    ElseIf cboComune.ListIndex = -1 Then
        nome = "CITTA'"
    ElseIf txtProv = "" Then
        nome = "PROVINCIA"
    ElseIf txtIva = "" Then
        nome = "PARTITA IVA"
    ElseIf txtImportoTicket.Text = "" Then
        nome = "TICKET"
    ElseIf txtQuotaAggiuntiva.Text = "" Then
        nome = "QUOTA REGIONALE"
    ElseIf txtQuotaNazionale.Text = "" Then
        nome = "QUOTA NAZIONALE"
    ElseIf txtImportoBollo.Text = "" Then
        nome = "IMPORTO BOLLO"
    Else
        Completo = True
        Exit Function
    End If
    MsgBox "Inserire i dati obbligatori" & vbCrLf & "Campo: " & nome, vbCritical, "Attenzione"
End Function

Private Function CampiCorretti() As Boolean
    Dim strNome As String
    CampiCorretti = True
    
    If Not txtQuotaAggiuntiva.IsCorrect Then
        strNome = "QUOTA REGIONALE"
        txtQuotaAggiuntiva.SetFocusSelected
    ElseIf Not txtQuotaNazionale.IsCorrect Then
        strNome = "QUOTA NAZIONALE"
        txtQuotaNazionale.SetFocusSelected
    ElseIf Not txtRimborsoSpeseViaggio.IsCorrect Then
        strNome = "RIMBORSO SPESE VIAGGIO"
        txtRimborsoSpeseViaggio.SetFocusSelected
    ElseIf Not txtImportoTicket.IsCorrect Then
        strNome = "TICKET"
        txtImportoTicket.SetFocusSelected
    ElseIf Not txtImportoBollo.IsCorrect Then
        strNome = "IMPORTO BOLLO"
        txtImportoBollo.SetFocusSelected
    End If
    
    
    
    If strNome <> "" Then
        MsgBox "Il campo " & strNome & " non è stato inserito correttamente.", vbCritical, "Attenzione"
        CampiCorretti = False
    End If
End Function


Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_nome() As Variant
    Dim strIban As String
    
    If Completo Then
        
        If Not CampiCorretti Then Exit Sub
    
        strIban = txtIbanAlfa(0) & txtIbanNum(0) & txtIbanAlfa(1) & txtIbanNum(1) & txtIbanNum(2) & txtIbanNum(3)
        v_nome = Array("KEY", "CODICE_ASL", "INDIRIZZO", "CAP", "CODICE_COMUNE", "PROV", "P_IVA", "CODICE_FISCALE", "INTESTATARIO_CC", "IBAN", "DICITURA", "TICKET", "QUOTA_AGGIUNTIVA", "QUOTA_NAZIONALE", "RIMBORSO_SPESE_VIAGGIO", "NUMERO_AUTORIZZAZIONE", "IMPORTO_BOLLO")
        v_Val = Array(1, cboAsl.ItemData(cboAsl.ListIndex), txtIndirizzo, txtCap, cboComune.ItemData(cboComune.ListIndex), txtProv, txtIva, txtCodiceFiscale, txtIntestatario, strIban, txtDicitura, txtImportoTicket.GetDecimal, txtQuotaAggiuntiva.GetDecimal, txtQuotaNazionale.GetDecimal, txtRimborsoSpeseViaggio.GetDecimal, txtAutorizzazione, txtImportoBollo.GetDecimal)
        
        Set rsDataset = New Recordset
        rsDataset.Open "INTESTAZIONE_FATTURA", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
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



Private Sub txtAutorizzazione_GotFocus()
    txtAutorizzazione.BackColor = colArancione
End Sub

Private Sub txtAutorizzazione_LostFocus()
    txtAutorizzazione.BackColor = vbWhite
End Sub

Private Sub txtCap_GotFocus()
    txtCap.BackColor = colArancione
End Sub

Private Sub txtCap_LostFocus()
    txtCap.BackColor = vbWhite
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

Private Sub txtIva_GotFocus()
    txtIva.BackColor = colArancione
End Sub

Private Sub txtIva_LostFocus()
    txtIva.BackColor = vbWhite
End Sub

Private Sub txtDicitura_GotFocus()
    txtDicitura.BackColor = colArancione
End Sub

Private Sub txtDicitura_LostFocus()
    txtDicitura.BackColor = vbWhite
End Sub

Private Sub txtProv_GotFocus()
    txtProv.BackColor = colArancione
End Sub

Private Sub txtProv_LostFocus()
    txtProv.BackColor = vbWhite
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

Private Sub txtIntestatario_GotFocus()
    txtIntestatario.BackColor = colArancione
End Sub

Private Sub txtIntestatario_LostFocus()
    txtIntestatario.BackColor = vbWhite
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

