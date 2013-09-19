VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProduttoreManutentore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scheda Produttore/Manutentore"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Scheda Produttore/Manutentore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtRagioneSociale 
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
         MaxLength       =   40
         TabIndex        =   1
         Top             =   480
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   2
         Top             =   960
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2400
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   8
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox txtPartitaIva 
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
         MaxLength       =   15
         TabIndex        =   9
         Top             =   3840
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   18
         TabIndex        =   10
         Top             =   4320
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2880
         Width           =   4335
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1920
         Width           =   495
      End
      Begin MSComDlg.CommonDialog cdlStampa 
         Left            =   6240
         Top             =   -120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Partita IVA"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   1110
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
         Index           =   38
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   675
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
         Index           =   41
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   945
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
         Index           =   40
         Left            =   120
         TabIndex        =   20
         Top             =   4320
         Width           =   1575
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
         Index           =   39
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   390
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
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   585
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
         Left            =   3840
         TabIndex        =   17
         Top             =   1920
         Width           =   555
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
         TabIndex        =   16
         Top             =   960
         Width           =   870
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
         Index           =   34
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ragione Sociale"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Frame fraAzioni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   7215
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
         Left            =   3960
         TabIndex        =   11
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
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmProduttoreManutentore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProduttoreManutentore As Recordset
Dim intProduttoreManutentoreKey As Integer
Dim modifica As Boolean

Private Function Completo() As Boolean
    Completo = False
    
    If txtRagioneSociale.Text = "" Then
        MsgBox "Inserire la REGIONE SOCIALE", vbInformation, "Informazione"
        txtRagioneSociale.SetFocus
        Exit Function
    End If
    
    Completo = True
End Function

Private Sub PulisciTutto()
    intProduttoreManutentoreKey = 0
    modifica = False
    Call PulisciForm(Me)
    txtRagioneSociale.SetFocus
End Sub

Private Sub cmdChiudi_Click()
    modifica = False
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim numKey As Integer
    
    If Completo Then
    
        Call SuperUcase(Me)
        
        Set rsProduttoreManutentore = New Recordset
        
        If modifica Then
            numKey = intProduttoreManutentoreKey
        Else
            numKey = GetNumero("PRODUTTORE_MANUTENTORE")
        End If

        v_Nomi = Array("KEY", "RAGIONE_SOCIALE", "INDIRIZZO", "CITTA", "CAP", "PROV", "TELEFONO" _
                    , "FAX", "EMAIL", "PARTITA_IVA", "CODICE_FISCALE")
        v_Val = Array(numKey, txtRagioneSociale, txtIndirizzo, txtCitta, txtCap, txtProv, txtTelefono _
                    , txtFax, txtEmail, txtPartitaIva, txtCodiceFiscale)
        
        If modifica Then
            rsProduttoreManutentore.Open "SELECT * FROM PRODUTTORE_MANUTENTORE WHERE KEY=" & intProduttoreManutentoreKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsProduttoreManutentore.Update v_Nomi, v_Val
        Else
            rsProduttoreManutentore.Open "PRODUTTORE_MANUTENTORE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsProduttoreManutentore.AddNew v_Nomi, v_Val
        End If
        
        Set rsProduttoreManutentore = Nothing
                        
        Call PulisciTutto
        
        cmdChiudi_Click

    End If
End Sub

Private Sub CaricaProduttoreManutentore()
    
        modifica = True
        
        Set rsProduttoreManutentore = New Recordset
        rsProduttoreManutentore.Open "SELECT * FROM PRODUTTORE_MANUTENTORE WHERE KEY=" & intProduttoreManutentoreKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        txtRagioneSociale = rsProduttoreManutentore("RAGIONE_SOCIALE") & ""
        txtIndirizzo = rsProduttoreManutentore("INDIRIZZO") & ""
        txtCitta = rsProduttoreManutentore("CITTA") & ""
        txtCap = rsProduttoreManutentore("CAP") & ""
        txtProv = rsProduttoreManutentore("PROV") & ""
        txtTelefono = rsProduttoreManutentore("TELEFONO") & ""
        txtFax = rsProduttoreManutentore("FAX") & ""
        txtEmail = rsProduttoreManutentore("EMAIL") & ""
        txtPartitaIva = rsProduttoreManutentore("PARTITA_IVA") & ""
        txtCodiceFiscale = rsProduttoreManutentore("CODICE_FISCALE") & ""
    
        Set rsProduttoreManutentore = Nothing
    
End Sub

Private Sub Form_Load()
    intProduttoreManutentoreKey = tTrova.keyReturn
    ' Se mi trovo in fase di modifica vado a caricare direttamente il record
    If intProduttoreManutentoreKey > 0 Then
        Call CaricaProduttoreManutentore
    End If
End Sub

Private Sub txtPartitaIva_KeyPress(KeyAscii As Integer)
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

Private Sub txtRagioneSociale_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
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

Private Sub txtFax_KeyPress(KeyAscii As Integer)
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

Private Sub txtStudio_KeyPress(KeyAscii As Integer)
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

Private Sub txtCitta_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
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

Private Sub txtCodiceFiscale_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtEmail_GotFocus()
    txtEmail.BackColor = colArancione
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtEmail_LostFocus()
    txtEmail.BackColor = vbWhite
End Sub

Private Sub txtFax_GotFocus()
    txtFax.BackColor = colArancione
End Sub

Private Sub txtFax_LostFocus()
    txtFax.BackColor = vbWhite
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

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtPartitaIva_GotFocus()
    txtPartitaIva.BackColor = colArancione
End Sub

Private Sub txtPartitaIva_LostFocus()
    txtPartitaIva.BackColor = vbWhite
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

Private Sub txtRagioneSociale_GotFocus()
    txtRagioneSociale.BackColor = colArancione
End Sub

Private Sub txtRagioneSociale_LostFocus()
    txtRagioneSociale.BackColor = vbWhite
End Sub

Private Sub txtTelefono_GotFocus()
    txtTelefono.BackColor = colArancione
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.BackColor = vbWhite
End Sub
