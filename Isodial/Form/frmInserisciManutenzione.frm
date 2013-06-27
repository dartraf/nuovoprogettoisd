VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmInserisciManutenzione 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraManutenzione 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.CheckBox chkSicurezza 
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
         Left            =   5880
         TabIndex        =   16
         Top             =   870
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkFunzionalità 
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
         Left            =   5880
         TabIndex        =   15
         Top             =   390
         Visible         =   0   'False
         Width           =   1575
      End
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
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   3975
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
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   1
         Top             =   2040
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
         Top             =   1440
         Width           =   6855
      End
      Begin DataTimeBox.uDataTimeBox oDataEffettivaManutenzione 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataScadenzaManutenzione 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oDataRichiestaManutenzione 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "N° Rif. Doc. di Lavoro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2355
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
         Top             =   870
         Width           =   2145
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
         Top             =   1320
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
         Left            =   3840
         TabIndex        =   8
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Scadenza Manut."
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
         TabIndex        =   14
         Top             =   390
         Visible         =   0   'False
         Width           =   2340
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
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Visible         =   0   'False
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   9975
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
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
         Left            =   6960
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
         Left            =   8640
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
    
    If tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA Then
         
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
                
    ElseIf tTabellaManutenzione = tpMANUTENZIONEORDINARIA Then
    
        If chkFunzionalità.Value = Checked And chkSicurezza.Value = Unchecked Then
            txtTipoManutenzione.Text = "ORD. FUNZ."
        ElseIf chkSicurezza.Value = Checked And chkFunzionalità.Value = Unchecked Then
            txtTipoManutenzione.Text = "ORD. SICUR."
        ElseIf chkFunzionalità.Value = Checked And chkSicurezza.Value = Checked Then
            txtTipoManutenzione.Text = "ORD. FUN. SIC."
        ElseIf chkFunzionalità.Value = Unchecked Or chkSicurezza.Value = Unchecked Then
            txtTipoManutenzione.Text = "ORDINARIA"
         End If
            
        v_Nomi = Array("KEY", "CODICE_APPARATO", "TIPO_MANUTENZIONE", "DATA_SCADENZA_MANUTENZIONE", "DATA_EFFETTIVA_MANUTENZIONE", "DESCRIZIONE_MANUTENZIONE", "DETTAGLI_INTERVENTO", "NUMERO_DOCUMENTO", "FUNZIONALITA", "SICUREZZA")
        
        v_Val = Array(numKey, KeyApparato, txtTipoManutenzione.Text, IIf(oDataScadenzaManutenzione(1).data = "", Null, oDataScadenzaManutenzione(1).data), IIf(oDataEffettivaManutenzione(1).data = "", Null, oDataEffettivaManutenzione(1).data), cboDescrizone(0).Text, cboDettagliIntervento(1).Text, txtNumeroDocumneto, IIf(chkFunzionalità.Value = Checked, True, False), IIf(chkSicurezza.Value = Checked, True, False))
            
        If KeyReturnManutenzione > 0 Then
            rsManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY=" & numKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsManutenzione.Update v_Nomi, v_Val
        Else
            rsManutenzione.Open "MANUTENZIONE_APPARATI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsManutenzione.AddNew v_Nomi, v_Val
        End If
            
        Set rsManutenzione = Nothing
    
    End If
                
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
            frmInserisciManutenzione.Caption = "Manutenzione Straordinaria"
            Label1(12).Visible = True
            oDataRichiestaManutenzione(0).Visible = True
            txtTipoManutenzione = "STRAORDINARIA"
            txtNumeroDocumneto_GotFocus
            If KeyReturnManutenzione > 0 Then
                Call CaricaManutenzione
            End If
            
        Case tpMANUTENZIONEORDINARIA
            frmInserisciManutenzione.Caption = "Manutenzione Ordinaria"
            chkFunzionalità.Visible = True
            chkSicurezza.Visible = True
            Label1(5).Visible = True
            oDataScadenzaManutenzione(1).Visible = True
            txtNumeroDocumneto_GotFocus
            If KeyReturnManutenzione > 0 Then
                Call CaricaManutenzione
            End If
            
    End Select
End Sub

Private Sub CaricaManutenzione()
    Set rsCercaManutenzione = New Recordset
    
    rsCercaManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY =" & KeyReturnManutenzione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        txtTipoManutenzione.Text = rsCercaManutenzione("TIPO_MANUTENZIONE")
        If tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA Then
            oDataRichiestaManutenzione(0).txtBox = rsCercaManutenzione("DATA_RICHIESTA_MANUTENZIONE") & ""
        ElseIf tTabellaManutenzione = tpMANUTENZIONEORDINARIA Then
            oDataScadenzaManutenzione(1).txtBox = rsCercaManutenzione("DATA_SCADENZA_MANUTENZIONE") & ""
            chkFunzionalità.Value = IIf(CBool(rsCercaManutenzione("FUNZIONALITA")), Checked, Unchecked)
            chkSicurezza.Value = IIf(CBool(rsCercaManutenzione("SICUREZZA")), Checked, Unchecked)
        End If
        oDataEffettivaManutenzione(1).txtBox = rsCercaManutenzione("DATA_EFFETTIVA_MANUTENZIONE") & ""
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
