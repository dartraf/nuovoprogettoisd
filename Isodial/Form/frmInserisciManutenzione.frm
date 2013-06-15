VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmInserisciManutenzione 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraManutenzioneStraordinaria 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.ComboBox cboDettagliIntervento 
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
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   6615
      End
      Begin VB.TextBox txtTipoManutenzione 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "STRAORDINARIA"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtNumeroDocumneto 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboDescrizone 
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
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   6615
      End
      Begin DataTimeBox.uDataTimeBox oDataRichiestaManutenzione 
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   3
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataEffettivaManutenzione 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Manutenzione"
         BeginProperty Font 
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
         TabIndex        =   15
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "Riferimento N° Doc. di Lavoro"
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
         Index           =   3
         Left            =   5160
         TabIndex        =   13
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Effettiva Manut."
         BeginProperty Font 
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
         TabIndex        =   12
         Top             =   990
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Rchiesta Manut."
         BeginProperty Font 
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
         Left            =   5160
         TabIndex        =   10
         Top             =   990
         Width           =   2220
      End
      Begin VB.Label Label1 
         Caption         =   "Descrizione Manutenzione o Motivazione Richiesta"
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dettagli Intervento"
         BeginProperty Font 
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
         TabIndex        =   8
         Top             =   2280
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   9735
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
         Left            =   6720
         TabIndex        =   6
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
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmInserisciManutenzione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsManutenzione As Recordset
Dim rsCercaManutenzione As Recordset
Dim NumeroApparato As Integer
Dim ModificaApparato As Boolean

Private Sub cboDescrizone_GotFocus(Index As Integer)
    cboDescrizone(0).BackColor = colArancione
End Sub

Private Sub cboDescrizone_LostFocus(Index As Integer)
    If Len(cboDescrizone(0)) > 120 Then
        MsgBox "Impossibile memorizzare più di 120 caratteri", vbCritical, "Attenzione"
        cboDescrizone(0).Text = ""
        cboDescrizone(0).SetFocus
        Exit Sub
    End If
    
    If cboDescrizone(0).Text <> "" Then
        Call GestisciNuovo("DESCRIZIONE_MANUTENZIONE", cboDescrizone(0))
    End If

    cboDescrizone(0).BackColor = vbWhite
End Sub

Private Sub cboDettagliIntervento_GotFocus(Index As Integer)
    cboDettagliIntervento(1).BackColor = colArancione
End Sub

Private Sub cboDettagliIntervento_LostFocus(Index As Integer)
    If Len(cboDettagliIntervento(1)) > 120 Then
        MsgBox "Impossibile memorizzare più di 120 caratteri", vbCritical, "Attenzione"
        cboDettagliIntervento(1).Text = ""
        cboDettagliIntervento(1).SetFocus
        Exit Sub
    End If
    
    If cboDettagliIntervento(1).Text <> "" Then
        Call GestisciNuovo("DETTAGLIO_MANUTENZIONE", cboDettagliIntervento(1))
    End If

    cboDettagliIntervento(1).BackColor = vbWhite
End Sub

Private Sub cmdChiudi_Click()
    If KeyReturnManutenzione > 0 Then
        Unload frmInserisciManutenzione
    Else
        KeyReturnManutenzione = -2
        Unload frmInserisciManutenzione
    End If
End Sub

Private Sub cmdMemorizza_Click()
Dim v_Nomi() As Variant
Dim v_Val() As Variant
Dim numKey As Integer
       
    Call SuperUcase(Me)
        
    Set rsManutenzione = New Recordset
        
    If KeyReturnManutenzione = 0 Then
        numKey = GetNumero("MANUTENZIONE_APPARATI")
    Else
        numKey = KeyReturnManutenzione
    End If
         
    v_Nomi = Array("KEY", "CODICE_APPARATO", "TIPO_MANUTENZIONE", "DATA_RICHIESTA_MANUTENZIONE", "DATA_EFFETTIVA_MANUTENZIONE", "DESCRIZIONE_MANUTENZIONE", "DETTAGLI_INTERVENTO", "NUMERO_DOCUMENTO")
        
    v_Val = Array(numKey, KeyApparato, txtTipoManutenzione.Text, IIf(oDataRichiestaManutenzione(0).data = "", Null, oDataRichiestaManutenzione(0).data), IIf(oDataEffettivaManutenzione(1).data = "", Null, oDataEffettivaManutenzione(1).data), cboDescrizone(0).Text, cboDettagliIntervento(1).Text, txtNumeroDocumneto)
            
    If KeyReturnManutenzione > 0 Then
        rsManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsManutenzione.Update v_Nomi, v_Val
    Else
        rsManutenzione.Open "MANUTENZIONE_APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsManutenzione.AddNew v_Nomi, v_Val
    End If
            
    Set rsManutenzione = Nothing
                
    Call Pulisci
        
    If KeyReturnManutenzione > 0 Then
        Unload frmInserisciManutenzione
    Else
        KeyReturnManutenzione = 0
        Unload frmInserisciManutenzione
    End If
    
End Sub

Private Sub Pulisci()
    oDataRichiestaManutenzione(0).Pulisci
    oDataEffettivaManutenzione(1).Pulisci
    cboDescrizone(0).ListIndex = -1
    cboDettagliIntervento(1).ListIndex = -1
    txtNumeroDocumneto.Text = ""
End Sub

Private Sub Form_Activate()
    Call RicaricaComboBox("DESCRIZIONE_MANUTENZIONE", "NOME", cboDescrizone(0))
    Call RicaricaComboBox("DETTAGLIO_MANUTENZIONE", "NOME", cboDettagliIntervento(1))
End Sub

Private Sub Form_Load()
    Select Case tTabellaManutenzione
        Case tpMANUNTENZIONESTRAORDINARIA
            fraManutenzioneStraordinaria.Visible = True
            frmInserisciManutenzione.Caption = "Manutenzione Straordinaria"
            txtNumeroDocumneto_GotFocus
            If KeyReturnManutenzione > 0 Then
                Call CaricaManutenzioneStraordinaria
            End If
            
        Case tpMANUTENZIONEORDINARIA
            frmInserisciManutenzione.Caption = "Manutenzione Ordinaria"
    
    End Select
End Sub

Private Sub CaricaManutenzioneStraordinaria()
    Set rsCercaManutenzione = New Recordset
    
    rsCercaManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY =" & KeyReturnManutenzione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        txtTipoManutenzione.Text = rsCercaManutenzione("TIPO_MANUTENZIONE")
        oDataRichiestaManutenzione(0).txtBox = rsCercaManutenzione("DATA_RICHIESTA_MANUTENZIONE")
        oDataEffettivaManutenzione(1).txtBox = rsCercaManutenzione("DATA_EFFETTIVA_MANUTENZIONE")
        cboDescrizone(0).Text = rsCercaManutenzione("DESCRIZIONE_MANUTENZIONE")
        cboDettagliIntervento(1).Text = rsCercaManutenzione("DETTAGLI_INTERVENTO")
        txtNumeroDocumneto.Text = rsCercaManutenzione("NUMERO_DOCUMENTO")
        
    Set rsCercaManutenzione = Nothing
End Sub

Private Sub txtNumeroDocumneto_GotFocus()
    txtNumeroDocumneto.BackColor = colArancione
End Sub

Private Sub txtNumeroDocumneto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumeroDocumneto_LostFocus()
    txtNumeroDocumneto.BackColor = vbWhite
End Sub
