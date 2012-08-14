VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAttendi 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   3240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblAttendi 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Attendere prego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3690
   End
   Begin VB.Label lblScritta 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "In attesa della connessione al server"
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
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4200
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAttendi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAttendi.frm
'
' <b>Descrizione</b>: Pannello Attendi richiamata dal client in fase di avvio e chiusura di Isodial
'
' @remarks
'
' @author
'
' @date 01/02/2011 21.41
Option Explicit

'' determina se il disco è condiviso
Dim condiviso As Boolean
'' determina se il client è connesso al server
Dim connesso As Boolean

'' Avvia la richiesta di connessione a condividi.exe
Private Sub Form_Load()
    On Error GoTo gestione
    If tRete = tpCONNETTI Then
        lblScritta = "In attesa della connessione al server"
    Else
        lblScritta = "In attesa della disconnessione al server"
    End If
    condiviso = False
    connesso = False
    wsk.Connect Mid(structApri.nomeServer, 3, Len(structApri.nomeServer)), 4000
    Timer2.Enabled = True
    Exit Sub
gestione:
    MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End
End Sub

'' Richiama la verifica di connessione
Private Sub Timer1_Timer()
    Static numPunti As Integer
    If numPunti <= 3 Then
        lblAttendi = lblAttendi & "."
    Else
        lblAttendi = Left(lblAttendi, Len(lblAttendi) - 3)
        numPunti = 0
        If condiviso Then
            If tRete = tpCONNETTI Then
                Call ProvaConnessione
            Else
                Unload Me
            End If
        End If
    End If
    numPunti = numPunti + 1
End Sub

'' Prova la connessione (per max 10 volte)
Private Sub ProvaConnessione()
    On Error GoTo gestione
    Static attesa As Integer
    attesa = attesa + 1
    If Dir(structApri.nomeServer & "\RISORSA\Centro.mdb") <> "" Then
        Unload Me
    End If
    Exit Sub
gestione:
    If attesa = 10 Then
        MsgBox "Impossibile aprire l'archivio" & vbCrLf & "Risorsa non condivisa", vbCritical, "Isodial"
        End
    End If
End Sub

'' Se il volume non è montato chiede a condividi.exe di montarlo e condividerlo
' In fase di chiusura chiede di essere scollegato
Private Sub Timer2_Timer()
    If connesso Then
        If tRete = tpCONNETTI Then
            wsk.SendData "condividi"
            condiviso = True
        Else
            wsk.SendData "chiudi"
            condiviso = True
        End If
        Timer2.Enabled = False
    Else
        MsgBox "Impossibile avviare Isodial" & vbCrLf & "Verificare la connessione al server", vbCritical, "Attenzione"
        End
    End If
End Sub

'' Imposta connesso a true
Private Sub wsk_Connect()
    connesso = True
End Sub

Private Sub wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Descrizione: " & Description, vbCritical, "Errore n# " & Number
    End
End Sub
