VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMediciBase 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scheda Medici di Base"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Scheda Medico di Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7935
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6735
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
         ItemData        =   "frmMediciBase.frx":0000
         Left            =   5840
         List            =   "frmMediciBase.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox txtRiceve 
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
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   5280
         Width           =   4215
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   1320
         Picture         =   "frmMediciBase.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Seleziona il medico"
         Top             =   360
         Width           =   450
      End
      Begin VB.ComboBox cboTipologia 
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
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   7440
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   2
         Top             =   960
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1920
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2880
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   11
         Top             =   4800
         Width           =   4215
      End
      Begin VB.TextBox txtStudio 
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
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   8
         Top             =   3360
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   9
         Top             =   3840
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   10
         Top             =   4320
         Width           =   4215
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   5
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtCodiceMedico 
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
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   13
         Top             =   6480
         Width           =   855
      End
      Begin VB.CheckBox chkPresenzaBarCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Barcode Cod.Fisc. su ricetta"
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
         Left            =   120
         TabIndex        =   14
         Top             =   6960
         Width           =   3975
      End
      Begin MSComDlg.CommonDialog cdlStampa 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Riceve"
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
         TabIndex        =   35
         Top             =   5280
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipologia medico"
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
         TabIndex        =   33
         Top             =   7470
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Studio"
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
         TabIndex        =   32
         Top             =   3360
         Width           =   675
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
         TabIndex        =   31
         Top             =   4800
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
         TabIndex        =   30
         Top             =   2880
         Width           =   945
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
         Index           =   40
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   945
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
         TabIndex        =   28
         Top             =   4320
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
         TabIndex        =   27
         Top             =   2400
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
         Left            =   5200
         TabIndex        =   26
         Top             =   2430
         Width           =   552
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
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Citt�"
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
         TabIndex        =   24
         Top             =   1440
         Width           =   480
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
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   630
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
         TabIndex        =   22
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codice Regionale"
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
         TabIndex        =   21
         Top             =   6480
         Width           =   1890
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
      Top             =   7800
      Width           =   6735
      Begin VB.CommandButton cmdElimina 
         Caption         =   "&Elimina"
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
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1215
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
         Height          =   495
         Left            =   1920
         TabIndex        =   18
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
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
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
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMediciBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmMediciBase.frm
'
' <b>Descrizione</b>: Scheda Medici di Base associata alla tab MEDICI_BASE
'
' @remarks
'
' @author
'
' @date 08/02/2011 20.59
Option Explicit

'' indica se si � in fase di modifica
Dim modifica As Boolean
'' rs della scheda
Dim rsMedico As Recordset
Dim intMediciBaseKey As Integer


Private Sub cmdElimina_Click()
    Dim blnElimina As Boolean
    Dim blnElimina2 As Boolean
    Dim blnEliminato As Boolean
    Dim rsDataset As Recordset
      
    If intMediciBaseKey = 0 Then
        Exit Sub
    Else
        blnElimina = IsPossibleDelete("PAZIENTI", "CODICE_MEDICO", intMediciBaseKey)
        blnElimina2 = IsPossibleDelete("RICETTE", "CODICE_MEDICO", intMediciBaseKey)
        
        If blnElimina And blnElimina2 Then
            If MsgBox("Sicuro di voler eliminare " & txtCognome & " " & txtNome & "?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                Set rsDataset = New Recordset
                rsDataset.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMediciBaseKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                If rsDataset.EOF And rsDataset.BOF Then
                    MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
                Else
                    rsDataset.Delete
                    blnEliminato = True
                End If
                Set rsDataset = Nothing
            End If
        Else
            MsgBox "Impossibile eliminare " & txtCognome & " " & txtNome & " perch� in relazione con altri dati del sistema", vbInformation, Me.Caption
        End If
    End If
            
    If blnEliminato Then
        Call PulisciTutto
        MsgBox "Eliminazione avvenuta con successo", vbInformation, Me.Caption
    End If
End Sub

'' Ricarica le cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    Call RicaricaComboBox("TIPOLOGIE_MEDICO", "NOME", cboTipologia)
    Call RicaricaComboBox("SIGLE_PROVINCIE", "NOME", cboProv)
    If cboTipologia = "" Then cboTipologia.ListIndex = 4
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    modifica = False
    Set rsMedico = New Recordset
    'rsMedico.Open "SELECT CODICE FROM (INTESTAZIONE_STAMPA I LEFT OUTER JOIN ASL A ON A.KEY=I.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Or structIntestazione.sCodiceSTS = CODICESTS_EM_IRPINA Then
    'If rsMedico("CODICE") = "201" Then
        Frame2.Height = 7935
        fraAzioni.Top = 7800
        frmMediciBase.Height = 9090
    Else
        Frame2.Height = 7335
        fraAzioni.Top = 7200
        frmMediciBase.Height = 8490
    End If
    Set rsMedico = Nothing
End Sub

'' Verifica prima di salvare che siano inserti i dati
Private Function Completo() As Boolean
    Completo = False
    If txtCognome.Text = "" Then
        MsgBox "Inserire il COGNOME del medico", vbCritical, "Attenzione"
        txtCognome.SetFocus
        Exit Function
    End If
    Completo = True
End Function

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    intMediciBaseKey = 0
    modifica = False
    chkPresenzaBarCode.Value = Unchecked
    Call PulisciForm(Me)
    txtCognome.SetFocus
End Sub

Private Sub cmdChiudi_Click()
    Unload frmMediciBase
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim numKey As Integer
    Dim codiceTimbro As String
    Dim pos As Integer
    
    If Completo Then
        Call SuperUcase(Me)
        Set rsMedico = New Recordset
        ' setta i valori
        If modifica Then
            numKey = intMediciBaseKey
        Else
            numKey = GetNumero("MEDICI_BASE")
        End If
        pos = InStr(txtCodiceMedico, "/")
        If pos <> 0 Then
            codiceTimbro = Mid(txtCodiceMedico, 1, pos - 1) & Mid(txtCodiceMedico, pos + 1, Len(txtCodiceMedico))
        Else
            codiceTimbro = txtCodiceMedico
        End If
        v_Nomi = Array("KEY", "COGNOME", "NOME", "COMUNE", "INDIRIZZO", "CAP", "PROV", "TELEFONO", "STUDIO" _
                    , "CELLULARE", "FAX", "EMAIL", "CODICE", "PRESENZA_BARCODE", "CODICE_TIPO_MEDICO", "RICEVE")
        v_Val = Array(numKey, txtCognome, txtNome, txtCitta, txtIndirizzo, txtCap, cboProv.Text, txtTelefono, txtStudio _
                    , txtCellulare, txtFax, txtEmail, codiceTimbro, IIf(chkPresenzaBarCode.Value = Checked, True, False), -1, txtRiceve)
        If cboTipologia.ListIndex <> -1 Then
            v_Val(14) = cboTipologia.ItemData(cboTipologia.ListIndex)
        End If
        If modifica Then
            rsMedico.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMediciBaseKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsMedico.Update v_Nomi, v_Val
        Else
            rsMedico.Open "MEDICI_BASE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsMedico.AddNew v_Nomi, v_Val
        End If
        Set rsMedico = Nothing
        Call PulisciTutto
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    End If
End Sub

Private Sub cmdStampa_Click()
    If intMediciBaseKey = 0 Then
        MsgBox "Selezionare il medico di base", vbInformation, "Attenzione"
    Else
    On Error GoTo gestione
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
'    Else
'        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
    Printer.FontSize = 16           'grandezza
    Printer.FontBold = True         'grassetto
    Printer.FontItalic = True       'corsivo
    Printer.Print
    Printer.Print , "                  SCHEDA MEDICO DI BASE"
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.FontUnderline = True    'sottolineato
    Printer.Print "                                                                                                                                                                                                        "
    Printer.FontUnderline = False
    Printer.Print
    Printer.Print
    Printer.Print "Cognome: ", txtCognome.Text
    Printer.Print
    Printer.Print
    Printer.Print "Nome: ", txtNome.Text
    Printer.Print
    Printer.Print
    Printer.Print "Citt�: ", txtCitta.Text
    Printer.Print
    Printer.Print
    Printer.Print "Indirizzo: ", txtIndirizzo.Text
    Printer.Print
    Printer.Print
    Printer.Print "C.A.P.: ", txtCap.Text, , "Prov.:", cboProv.Text
    Printer.Print
    Printer.Print
    Printer.Print "Telefono: ", txtTelefono.Text
    Printer.Print
    Printer.Print
    Printer.Print "Studio: ", txtStudio.Text
    Printer.Print
    Printer.Print
    Printer.Print "Cellulare: ", txtCellulare.Text
    Printer.Print
    Printer.Print
    Printer.Print "Fax: ", txtFax.Text
    Printer.Print
    Printer.Print
    Printer.Print "E-mail: ", txtEmail.Text
    Printer.Print
    Printer.Print
    Printer.Print "Codice Medico: ", txtCodiceMedico.Text
    Printer.Print
    Printer.Print
    Printer.Print "Riceve: ", txtRiceve.Text
    Printer.EndDoc
    End If
End Sub

Private Sub cmdTrova_Click()
    Call PulisciTutto
    tTrova.Tipo = tpMEDICOBASE
    tTrova.condizione = ""
    tTrova.condStato = ""
    Unload frmTrova
    frmTrova.Show 1
    intMediciBaseKey = tTrova.keyReturn
    Call CaricaMedico
End Sub

Private Sub cboTipologia_DropDown()
    Call SetComboWidth(cboTipologia, 280)
End Sub

'' Carica i dati del medico nel form
Private Sub CaricaMedico()
    If intMediciBaseKey = 0 Then Exit Sub
    modifica = True
    Set rsMedico = New Recordset
    rsMedico.Open "SELECT * FROM MEDICI_BASE WHERE KEY=" & intMediciBaseKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    txtCap = rsMedico("CAP") & ""
    txtCellulare = rsMedico("CELLULARE") & ""
    txtCitta = rsMedico("COMUNE") & ""
    txtCognome = rsMedico("COGNOME") & ""
    txtEmail = rsMedico("EMAIL") & ""
    txtFax = rsMedico("FAX") & ""
    txtIndirizzo = rsMedico("INDIRIZZO") & ""
    txtNome = rsMedico("NOME") & ""
    If rsMedico("PROV") = "" Then
        cboProv.ListIndex = -1
    Else
        cboProv.Text = rsMedico("PROV") & ""
    End If
    txtStudio = rsMedico("STUDIO") & ""
    txtTelefono = rsMedico("TELEFONO") & ""
    txtCodiceMedico = rsMedico("CODICE") & ""
    txtRiceve = rsMedico("RICEVE") & ""
    chkPresenzaBarCode.Value = IIf(CBool(rsMedico("PRESENZA_BARCODE")), Checked, Unchecked)
    cboTipologia.ListIndex = GetCboListIndex(rsMedico("CODICE_TIPO_MEDICO"), cboTipologia)
    Set rsMedico = Nothing
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

Private Sub txtCellulare_GotFocus()
    txtCellulare.BackColor = colArancione
End Sub

Private Sub txtCellulare_LostFocus()
    txtCellulare.BackColor = vbWhite
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

Private Sub txtCognome_GotFocus()
    txtCognome.BackColor = colArancione
End Sub

Private Sub txtcogNome_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCognome_LostFocus()
    txtCognome.BackColor = vbWhite
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

Private Sub txtCodiceMedico_GotFocus()
    txtCodiceMedico.BackColor = colArancione
End Sub

Private Sub txtCodiceMedico_KeyPress(KeyAscii As Integer)
    Call InvioTab(KeyAscii)
End Sub

Private Sub txtCodiceMedico_LostFocus()
    txtCodiceMedico.BackColor = vbWhite
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

Private Sub txtRiceve_GotFocus()
    txtRiceve.BackColor = colArancione
End Sub

Private Sub txtRiceve_LostFocus()
    txtRiceve.BackColor = vbWhite
End Sub

Private Sub txtStudio_GotFocus()
    txtStudio.BackColor = colArancione
End Sub

Private Sub txtStudio_LostFocus()
    txtStudio.BackColor = vbWhite
End Sub

Private Sub txtTelefono_GotFocus()
    txtTelefono.BackColor = colArancione
End Sub

Private Sub txtTelefono_LostFocus()
    txtTelefono.BackColor = vbWhite
End Sub

' insieme di sub per la gestione di valori solo numerici

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
