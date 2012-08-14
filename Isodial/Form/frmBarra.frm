VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBarra 
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   600
   End
   Begin VB.Frame fraAttendere 
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
      Left            =   0
      TabIndex        =   0
      Top             =   -70
      Width           =   4575
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblDescrizone 
         Caption         =   "Processo in corso..."
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lblAttendere 
         Caption         =   "ATTENDERE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1340
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Static intNumVolte As Integer
    
    If intNumVolte = 3 Then
        intNumVolte = 1
    Else
        intNumVolte = intNumVolte + 1
    End If
    
    lblDescrizone = "Processo in corso" & String(intNumVolte, ".")
End Sub
