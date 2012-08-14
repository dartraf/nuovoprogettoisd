VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup Incrementale"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtNumero 
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° di backup"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblNumMax 
         AutoSize        =   -1  'True
         Caption         =   "N° max di backup: "
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
      Begin VB.CommandButton cmdEsci 
         Cancel          =   -1  'True
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
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmBackup.frm
'
' <b>Descrizione</b>: Pannello per il settaggio del num di backup incrementali da effettuare nella penna
'
' @remarks
'
' @author
'
' @date 01/02/2011 21.48
Option Explicit

'' rs della scheda
Dim rsImpostazioni As Recordset
'' numero massimo di backup possibile nella penna
Dim numMax As Integer
Const Megabyte = 1048576

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    Set rsImpostazioni = New Recordset
    rsImpostazioni.Open "IMPOSTAZIONI_BACKUP", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    txtNumero = rsImpostazioni("NUMERO")
    Set rsImpostazioni = Nothing
    Call CalcolaMax
End Sub

'' Calcola il massimo di backup possibili nella penna
Private Sub CalcolaMax()
    Dim dimVolume  As Double
    Dim dimPenna As Double
    Dim lettera As String
    
    ' verifica la presenza della penna
    If Not VerificaDiscoRimovibile(lettera) Then
        MsgBox "Impossibile continuare" & vbCrLf & "Unita' di backup mancante", vbCritical, "Apertura archivio"
        Unload Me
    End If
    dimPenna = CLng(GetDriveSize(lettera & ":"))
    dimVolume = FileLen(structApri.pathVolume & "\" & nomeVolume) / Megabyte
    numMax = Int(dimPenna / dimVolume)
    lblNumMax = lblNumMax & numMax
End Sub

'' Memorizza il num di backup
Private Sub Memorizza()
    Set rsImpostazioni = New Recordset
    rsImpostazioni.Open "IMPOSTAZIONI_BACKUP", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsImpostazioni("NUMERO") = txtNumero
    rsImpostazioni.Update
    Set rsImpostazioni = Nothing
End Sub

'' Chiude il pannello
Private Sub cmdEsci_Click()
    Unload Me
End Sub

'' Controlla che il num inserito sia < del massimo consentito
Private Sub Form_Unload(Cancel As Integer)
    If txtNumero > numMax Then
        MsgBox "Numero di backup superiore al limite massimo", vbCritical, "Attenzione"
        txtNumero = numMax
        Cancel = True
    Else
        Call Memorizza
        Cancel = False
    End If
End Sub

Private Sub txtNumero_GotFocus()
    txtNumero.BackColor = colArancione
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNumero_LostFocus()
    txtNumero.BackColor = vbWhite
End Sub

