VERSION 5.00
Begin VB.Form frmAccompagnatori 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scheda Accompagnatori"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Autovettura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   10215
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
         Left            =   6000
         MaxLength       =   15
         TabIndex        =   17
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtProprietario 
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
         MaxLength       =   25
         TabIndex        =   18
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtTarga 
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
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proprietario"
         BeginProperty Font 
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
         TabIndex        =   38
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Tipo 
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
         Index           =   8
         Left            =   5400
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Targa"
         BeginProperty Font 
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
         TabIndex        =   36
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   10215
      Begin VB.PictureBox picDataRilascio 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   8520
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   14
         ToolTipText     =   "Cerca data"
         Top             =   320
         Width           =   360
      End
      Begin VB.TextBox txtPatente 
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
         MaxLength       =   10
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtEnteEmittente 
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
         TabIndex        =   15
         Top             =   840
         Width           =   1935
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
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   375
         Width           =   825
      End
      Begin VB.Label lblDataRilascio 
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
         Left            =   7200
         TabIndex        =   13
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data di Rilascio"
         BeginProperty Font 
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
         Left            =   5400
         TabIndex        =   33
         Top             =   375
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rilasciata da"
         BeginProperty Font 
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
         TabIndex        =   32
         Top             =   840
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accompagnatori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   1320
         Picture         =   "frmAccompagnatori.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtComuneProvNascita 
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
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
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
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
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
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   0
         Top             =   480
         Width           =   3255
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   3240
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   3
         ToolTipText     =   "Cerca data"
         Top             =   930
         Width           =   360
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
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1920
         Width           =   615
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
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1920
         Width           =   855
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
         MaxLength       =   16
         TabIndex        =   9
         Top             =   2400
         Width           =   2055
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
         MaxLength       =   25
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
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
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
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
         Left            =   6720
         MaxLength       =   25
         TabIndex        =   1
         Top             =   480
         Width           =   3255
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
         TabIndex        =   10
         Top             =   2880
         Width           =   2055
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
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1485
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cellulare"
         BeginProperty Font 
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
         Left            =   5400
         TabIndex        =   29
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label lblDataNascita 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1215
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
         Index           =   32
         Left            =   120
         TabIndex        =   28
         Top             =   480
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
         Index           =   33
         Left            =   5400
         TabIndex        =   27
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Città di Residenza"
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
         Index           =   34
         Left            =   5400
         TabIndex        =   26
         Top             =   1320
         Width           =   1125
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
         Index           =   35
         Left            =   120
         TabIndex        =   25
         Top             =   1950
         Width           =   870
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
         Index           =   36
         Left            =   8040
         TabIndex        =   24
         Top             =   1950
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.A.P"
         BeginProperty Font 
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
         Left            =   5400
         TabIndex        =   23
         Top             =   1950
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Fiscale"
         BeginProperty Font 
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
         Left            =   120
         TabIndex        =   22
         Top             =   2430
         Width           =   1575
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1620
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
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   5760
      Width           =   10215
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
         Left            =   7080
         TabIndex        =   41
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
         Left            =   8760
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAccompagnatori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAccompagnatori.frm
'
' <b>Descrizione</b>: Scheda Accompagnatori associata alla tab ACCOMPAGNATORI
'
' @remarks
'
' @author
'
' @date 01/02/2011 21.12
Option Explicit
'' rs della scheda
Dim rsDataset As Recordset
'' indica se si è in fase di modifica
Dim modifica As Boolean
Dim intAccompagnatoreKey As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

'' Carica le impostazioni iniziali
Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    picData.Picture = LoadResPicture("cal1", 0)
    picDataRilascio.Picture = LoadResPicture("cal1", 0)
End Sub

'' Determina se la scheda è completa prima del salvataggio
Private Function Completo() As Boolean
    Completo = False
    If txtCognome.Text = "" Then
        MsgBox "Inserire il COGNOME dell'accompagnatore", vbCritical, "Attenzione"
        txtCognome.SetFocus
        Exit Function
    ElseIf txtNome.Text = "" Then
        MsgBox "Inserire il NOME dell'accompagnatore", vbCritical, "Attenzione"
        txtNome.SetFocus
        Exit Function
    End If
    Completo = True
End Function

'' Pulisce tutta la scheda
Private Sub PulisciTutto()
    intAccompagnatoreKey = 0
    modifica = False
    lblDataNascita = ""
    lblDataRilascio = ""
    Call PulisciForm(Me)
    txtCognome.SetFocus
End Sub

'' Carica i dati dell'accompagnatore
Private Sub CaricaAccompagnatore()
    If intAccompagnatoreKey = 0 Then Exit Sub
    modifica = True
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM ACCOMPAGNATORI WHERE KEY=" & intAccompagnatoreKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    txtCognome = rsDataset("COGNOME")
    txtNome = rsDataset("NOME")
    txtComuneProvNascita = rsDataset("CITTA_NASCITA")
    lblDataNascita = rsDataset("DATA_NASCITA") & ""
    txtCitta = rsDataset("CITTA")
    txtIndirizzo = rsDataset("INDIRIZZO")
    txtCap = rsDataset("CAP")
    txtProv = rsDataset("PROV")
    txtCodiceFiscale = rsDataset("CODICE_FISCALE")
    txtCellulare = rsDataset("CELLULARE")
    txtTelefono = rsDataset("TELEFONO")
    txtEnteEmittente = rsDataset("ENTE_EMITTENTE")
    txtPatente = rsDataset("PATENTE")
    lblDataRilascio = rsDataset("DATA_RILASCIO") & ""
    txtTarga = rsDataset("TARGA")
    txtTipo = rsDataset("TIPO")
    txtProprietario = rsDataset("PROPRIETARIO")
    Set rsDataset = Nothing
End Sub

Private Sub cmdMemorizza_Click()
   If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Or structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Then
   Else
       MsgBox "MODULO OPZIONALE ATTIVABILE A RICHIESTA", vbInformation, "INFORMAZIONE"
       Exit Sub
   End If
       
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim numKey As Integer
    
    If Completo Then
        Call SuperUcase(Me)
        Set rsDataset = New Recordset
        ' setta i valori
        If modifica Then
            numKey = intAccompagnatoreKey
        Else
            numKey = GetNumero("ACCOMPAGNATORI")
        End If
        v_Nomi = Array("KEY", "COGNOME", "NOME", "CITTA_NASCITA", "DATA_NASCITA", "CODICE_FISCALE", "INDIRIZZO", "CAP", "CITTA", "PROV", "TELEFONO", "CELLULARE", "PATENTE", "DATA_RILASCIO", "ENTE_EMITTENTE", "TARGA", "TIPO", "PROPRIETARIO")
        v_Val = Array(numKey, txtCognome, txtNome, txtComuneProvNascita, IIf(lblDataNascita = "", Null, lblDataNascita), txtCodiceFiscale, txtIndirizzo, txtCap, txtCitta, txtProv, txtTelefono, txtCellulare, txtPatente, IIf(lblDataRilascio = "", Null, lblDataRilascio), txtEnteEmittente, txtTarga, txtTipo, txtProprietario)
        If modifica Then
            rsDataset.Open "SELECT * FROM ACCOMPAGNATORI WHERE KEY=" & intAccompagnatoreKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDataset.Update v_Nomi, v_Val
        Else
            rsDataset.Open "ACCOMPAGNATORI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.AddNew v_Nomi, v_Val
        End If
        Set rsDataset = Nothing
        Call PulisciTutto
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdTrova_Click()
    tTrova.Tipo = tpACCOMPAGNATORI
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intAccompagnatoreKey = tTrova.keyReturn
    Call CaricaAccompagnatore
End Sub

Private Sub picData_Click()
    frmCalendario.Show 1
    If laData <> "" Then lblDataNascita = laData
End Sub

Private Sub picData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub picDataRilascio_Click()
    frmCalendario.Show 1
    If laData <> "" Then lblDataRilascio = laData
End Sub

Private Sub picDataRilascio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDataRilascio.Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picDataRilascio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDataRilascio.Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub txtCap_GotFocus()
    txtCap.BackColor = colArancione
End Sub

Private Sub txtCAP_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
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

'' Controlla la validita del codice fiscale
Private Sub txtCodiceFiscale_Validate(Cancel As Boolean)
    If txtCodiceFiscale = "" Then
        Cancel = False
    Else
        If Len(txtCodiceFiscale) = 16 Then
            Cancel = Not ControlloCodiceFiscale(UCase(txtCodiceFiscale))
        Else
            MsgBox "Devi inserire solo 16 lettere/cifre", vbCritical, "Attenzione"
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

'' Per il controllo sul codice fiscale
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

'' Per il controllo sul codice fiscale
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

Private Sub txtComuneProvNascita_GotFocus()
    txtComuneProvNascita.BackColor = colArancione
End Sub

Private Sub txtComuneProvNascita_LostFocus()
    txtComuneProvNascita.BackColor = vbWhite
End Sub

Private Sub txtCodiceFiscale_GotFocus()
    txtCodiceFiscale.BackColor = colArancione
End Sub

Private Sub txtCodiceFiscale_LostFocus()
    txtCodiceFiscale.BackColor = vbWhite
End Sub

Private Sub txtCognome_GotFocus()
    txtCognome.BackColor = colArancione
End Sub

Private Sub txtCognome_LostFocus()
    txtCognome.BackColor = vbWhite
End Sub

Private Sub txtEnteEmittente_GotFocus()
    txtEnteEmittente.BackColor = colArancione
End Sub

Private Sub txtEnteEmittente_LostFocus()
    txtEnteEmittente.BackColor = vbWhite
End Sub

Private Sub txtIndirizzo_GotFocus()
    txtIndirizzo.BackColor = colArancione
End Sub

Private Sub txtIndirizzo_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtIndirizzo_LostFocus()
    txtIndirizzo.BackColor = vbWhite
End Sub

Private Sub txtNome_GotFocus()
    txtNome.BackColor = colArancione
End Sub

Private Sub txtNome_LostFocus()
    txtNome.BackColor = vbWhite
End Sub

Private Sub txtPatente_GotFocus()
    txtPatente.BackColor = colArancione
End Sub

Private Sub txtPatente_LostFocus()
    txtPatente.BackColor = vbWhite
End Sub

Private Sub txtProv_GotFocus()
    txtProv.BackColor = colArancione
End Sub

Private Sub txtProv_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtProv_LostFocus()
    txtProv.BackColor = vbWhite
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.BackColor = vbWhite
End Sub

Private Sub txtTelefono_GotFocus()
    txtTelefono.BackColor = colArancione
End Sub

Private Sub txtCellulare_LostFocus()
    txtCellulare.BackColor = vbWhite
End Sub

Private Sub txtCellulare_GotFocus()
    txtCellulare.BackColor = colArancione
End Sub

Private Sub txtProprietario_LostFocus()
    txtProprietario.BackColor = vbWhite
End Sub

Private Sub txtProprietario_GotFocus()
    txtProprietario.BackColor = colArancione
End Sub

Private Sub TXTTIPO_LostFocus()
    txtTipo.BackColor = vbWhite
End Sub

Private Sub TXTTIPO_GotFocus()
    txtTipo.BackColor = colArancione
End Sub

Private Sub txtTarga_LostFocus()
    txtTarga.BackColor = vbWhite
End Sub

Private Sub txtTarga_GotFocus()
    txtTarga.BackColor = colArancione
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case vbKeyReturn
            Call InvioTab(KeyAscii)
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtCellulare_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc(" "), vbKeyBack
        Case vbKeyReturn
            Call InvioTab(KeyAscii)
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

