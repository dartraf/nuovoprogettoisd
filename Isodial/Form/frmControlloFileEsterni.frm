VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmControlloFileEsterni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controllo configurazione sistema"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame panTesto 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Label lblTesto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame panComandi 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   7575
      Begin VB.CommandButton cmdIndietro 
         Caption         =   "&Chiudi"
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
         Left            =   6120
         TabIndex        =   3
         ToolTipText     =   "Chiude il programma"
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog oCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmControlloFileEsterni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsporta_Click()
 '   On Error GoTo gestione
 '   oCommonDialog.DefaultExt = "*.txt"
 '   oCommonDialog.DialogTitle = "Salva file"
 '   oCommonDialog.Filter = "File di testo|*.txt"
 '   oCommonDialog.ShowSave
    
 '   If oCommonDialog.FileName <> "" Then
 '       Open oCommonDialog.FileName For Output As 1
 '       Print #1, lblTesto.Caption
 '       Close #1
 '   End If
    
'    Exit Sub
'gestione:
'    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdIndietro_Click()
    End
End Sub

Private Sub Form_Activate()
    Call TakeCloseOff(Me.hWnd)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        End
    End If
End Sub
