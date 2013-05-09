VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmApparatiInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserimento Gestioni Apparati"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.ComboBox cboManutentore 
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
         Left            =   6960
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.ComboBox cboProduttore 
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
         TabIndex        =   6
         Top             =   1800
         Width           =   3375
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1320
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   2175
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   14
         Top             =   3720
         Width           =   615
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
         Height          =   1005
         Left            =   6960
         MaxLength       =   107
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   3255
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
         TabIndex        =   5
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
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   2
         Top             =   840
         Width           =   3375
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
         Left            =   6960
         Sorted          =   -1  'True
         TabIndex        =   3
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
      Begin DataTimeBox.uDataTimeBox oDataScadenza 
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   9
         Top             =   2280
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
         Left            =   2520
         TabIndex        =   8
         Top             =   2280
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
         Left            =   7320
         TabIndex        =   11
         Top             =   2760
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
         Left            =   2520
         TabIndex        =   12
         Top             =   3240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Ammortamento"
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
         TabIndex        =   31
         Top             =   3720
         Width           =   2175
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
         TabIndex        =   29
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
         Left            =   5280
         TabIndex        =   28
         Top             =   2790
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
         TabIndex        =   27
         Top             =   2310
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modalit� Acquisizione"
         BeginProperty Font 
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
         TabIndex        =   23
         Top             =   2790
         Width           =   2295
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
         Height          =   240
         Index           =   3
         Left            =   5280
         TabIndex        =   22
         Top             =   3270
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
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N� Apparato"
         BeginProperty Font 
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
         TabIndex        =   20
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Scadenza"
         BeginProperty Font 
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
         Left            =   5280
         TabIndex        =   19
         Top             =   2310
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Apparato"
         BeginProperty Font 
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
         Left            =   5280
         TabIndex        =   18
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N� Inventario"
         BeginProperty Font 
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
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   4440
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
         TabIndex        =   15
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
         TabIndex        =   16
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
Dim NumeroApparato As Integer
Dim ModificaApparato As Boolean

Private Sub cboManutentore_GotFocus(Index As Integer)
    cboManutentore(1).BackColor = colArancione
End Sub

Private Sub cboManutentore_LostFocus(Index As Integer)
        
    If Len(cboManutentore(1)) > 30 Then
        MsgBox "Impossibile memorizzare pi� di 30 caratteri", vbInformation, "Informazione"
        cboManutentore(1).Text = ""
        cboManutentore(1).SetFocus
        Exit Sub
    End If
    
    If cboManutentore(1).Text <> "" Then
        Call GestisciNuovo("APPARATI_PRODUT_MANUTENT", cboManutentore(1))
        Call RicaricaComboBox("APPARATI_PRODUT_MANUTENT", "NOME", cboProduttore(0))
    End If
    
    cboManutentore(1).BackColor = vbWhite
    
End Sub

Private Sub cboModalitaAcquisizione_GotFocus(Index As Integer)
    cboModalitaAcquisizione(1).BackColor = colArancione
End Sub

Private Sub cboModalitaAcquisizione_LostFocus(Index As Integer)

    If Len(cboModalitaAcquisizione(1)) > 15 Then
        MsgBox "Impossibile memorizzare pi� di 15 caratteri", vbInformation, "Informazione"
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
        MsgBox "Impossibile memorizzare pi� di 30 caratteri", vbInformation, "Informazione"
        cboModello(2).Text = ""
        cboModello(2).SetFocus
        Exit Sub
    End If
    
    If cboModello(2).Text <> "" Then
        Call GestisciNuovo("APPARATI_MODELLO", cboModello(2))
    End If
    
    cboModello(2).BackColor = vbWhite
    
End Sub

Private Sub cboProduttore_GotFocus(Index As Integer)
    cboProduttore(0).BackColor = colArancione
End Sub

Private Sub cboProduttore_LostFocus(Index As Integer)
    
    If Len(cboProduttore(0)) > 30 Then
        MsgBox "Impossibile memorizzare pi� di 30 caratteri", vbInformation, "Informazione"
        cboProduttore(0).Text = ""
        cboProduttore(0).SetFocus
        Exit Sub
    End If
    
    If cboProduttore(0).Text <> "" Then
        Call GestisciNuovo("APPARATI_PRODUT_MANUTENT", cboProduttore(0))
        Call RicaricaComboBox("APPARATI_PRODUT_MANUTENT", "NOME", cboManutentore(1))
    End If
    
    cboProduttore(0).BackColor = vbWhite
    
End Sub

Private Sub cboTipoApparato_GotFocus(Index As Integer)
        cboTipoApparato(0).BackColor = colArancione
End Sub

Private Sub cboTipoApparato_LostFocus(Index As Integer)
            
    If Len(cboTipoApparato(0)) > 30 Then
        MsgBox "Impossibile memorizzare pi� di 30 caratteri", vbInformation, "Informazione"
        cboTipoApparato(0).Text = ""
        cboTipoApparato(0).SetFocus
        Exit Sub
    End If
    
    
    If cboTipoApparato(0).Text <> "" Then
        Call GestisciNuovo("APPARATI_TIPO", cboTipoApparato(0))
    End If

    cboTipoApparato(0).BackColor = vbWhite
    
End Sub

Private Sub cmdChiudi_Click()
    Unload frmApparatiInput
End Sub

Private Sub cmdMemorizza_Click()
Dim v_Nomi() As Variant
Dim v_Val() As Variant
Dim numKey As Integer

    'decidere i campi obbligatori oltre al numero inventario

    If txtNumeroInventario.Text = "" Then
        MsgBox "Inserire il N� Inventario", vbInformation, "Informazione"
        Exit Sub
    End If
       
    If txtPeriodoAmmortamento = "" Then
        txtPeriodoAmmortamento = 0
    End If
       
    Call SuperUcase(Me)
        
    Set rsMemorizzaApparecchiature = New Recordset
        
    If ModificaApparato = True Then
         numKey = NumeroApparato
    Else
        numKey = GetNumero("APPARATI")
    End If
        
    v_Nomi = Array("KEY", "NUMERO_INVENTARIO", "NUMERO_APPARATO", "TIPO_APPARATO", "MODELLO", "MATRICOLA", "PRODUTTORE", "MANUTENTORE", "DATA_FABBRICAZIONE" _
                    , "DATA_COLLAUDO", "NOTE_COLLAUDO", "DATA_SCADENZA", "MODALITA_ACQUISIZIONE", "DATA_ACQUISIZIONE", "PERIODO_AMMORTAMENTO")
        
    v_Val = Array(numKey, txtNumeroInventario, txtNumeroApparato, cboTipoApparato(0).Text, cboModello(2).Text, txtMatricola, cboProduttore(0).Text, cboManutentore(1).Text, IIf(oDataFabbricazione(0).data = "", Null, oDataFabbricazione(0).data) _
                    , IIf(oDataCollaudo(3).data = "", Null, oDataCollaudo(3).data), txtNoteCollaudo, IIf(oDataScadenza(1).data = "", Null, oDataScadenza(1).data), cboModalitaAcquisizione(1).Text, IIf(oDataAcquisizione(2).data = "", Null, oDataAcquisizione(2).data), txtPeriodoAmmortamento)
            
    If ModificaApparato = True Then
        rsMemorizzaApparecchiature.Open "SELECT * FROM APPARATI WHERE KEY=" & NumeroApparato, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsMemorizzaApparecchiature.Update v_Nomi, v_Val
    Else
        rsMemorizzaApparecchiature.Open "APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsMemorizzaApparecchiature.AddNew v_Nomi, v_Val
    End If
            
    Set rsMemorizzaApparecchiature = Nothing
    
    Call Pulisci
    Call NumeroInventario
    
    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    ModificaApparato = False
End Sub

Private Sub Pulisci()
    txtNumeroApparato.Text = ""
    cboTipoApparato(0).Text = ""
    cboModello(2).Text = ""
    txtMatricola.Text = ""
    cboProduttore(0).Text = ""
    cboManutentore(1).Text = ""
    oDataFabbricazione(0).Pulisci
    oDataScadenza(1).Pulisci
    cboModalitaAcquisizione(1).Text = ""
    oDataAcquisizione(2).Pulisci
    oDataCollaudo(3).Pulisci
    txtNoteCollaudo.Text = ""
    txtPeriodoAmmortamento.Text = ""
    txtNumeroInventario.SetFocus
    txtNumeroInventario_GotFocus
    NumeroApparato = 0
    tTrova.keyReturn = 0
End Sub

Private Sub Form_Activate()
    Call RicaricaComboBox("APPARATI_TIPO", "NOME", cboTipoApparato(0))
    Call RicaricaComboBox("APPARATI_MODELLO", "NOME", cboModello(2))
    Call RicaricaComboBox("APPARATI_PRODUT_MANUTENT", "NOME", cboProduttore(0))
    Call RicaricaComboBox("APPARATI_PRODUT_MANUTENT", "NOME", cboManutentore(1))
    Call RicaricaComboBox("APPARATI_MOD_ACQ", "NOME", cboModalitaAcquisizione(1))
End Sub

Private Sub Form_Load()
    If tTrova.keyReturn = 0 Then
        Call NumeroInventario
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
        cboModello(2).Text = rsCercaApparato("MODELLO")
        txtMatricola.Text = rsCercaApparato("MATRICOLA")
        cboProduttore(0).Text = rsCercaApparato("PRODUTTORE")
        cboManutentore(1).Text = rsCercaApparato("MANUTENTORE")
        oDataFabbricazione(0).txtBox = rsCercaApparato("DATA_FABBRICAZIONE") & ""
        oDataScadenza(1).txtBox = rsCercaApparato("DATA_SCADENZA") & ""
        cboModalitaAcquisizione(1).Text = rsCercaApparato("MODALITA_ACQUISIZIONE")
        oDataAcquisizione(2).txtBox = rsCercaApparato("DATA_ACQUISIZIONE") & ""
        oDataCollaudo(3).txtBox = rsCercaApparato("DATA_COLLAUDO") & ""
        txtNoteCollaudo.Text = rsCercaApparato("NOTE_COLLAUDO")
        txtPeriodoAmmortamento.Text = rsCercaApparato("PERIODO_AMMORTAMENTO")
        
    Set rsCercaApparato = Nothing
    ModificaApparato = True
    
End Sub

Private Sub NumeroInventario()
    Set rsNumeroProgressivo = New Recordset
    
    rsNumeroProgressivo.Open "SELECT MAX(NUMERO_INVENTARIO) AS MASSIMO FROM APPARATI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsNumeroProgressivo("MASSIMO")) Then
        txtNumeroInventario = rsNumeroProgressivo("MASSIMO") + 1
    Else
        txtNumeroInventario = 1
    End If
    
    Set rsNumeroProgressivo = Nothing
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
