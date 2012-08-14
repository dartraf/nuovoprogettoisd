VERSION 5.00
Begin VB.Form frmPannelloSceltaTurno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scelta turno"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPeriodo 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Sera"
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
         Index           =   5
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Pomeriggio"
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
         Index           =   4
         Left            =   2760
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Dispari Mattina"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Sera"
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
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Pomeriggio"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optSessione 
         Caption         =   "Pari Mattina"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   5295
      Begin VB.CommandButton cmdEsci 
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
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdAvanti 
         Cancel          =   -1  'True
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
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmPannelloSceltaTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Sessione As enumSessioni

Public Property Get GetSessione() As enumSessioni
    GetSessione = m_Sessione
End Property

Public Property Let LetSessione(ByVal intSessione As enumSessioni)
    m_Sessione = intSessione
End Property

Private Sub cmdAvanti_Click()
    Dim i As Integer
        
    For i = 0 To 5
        If optSessione(i).Value = True Then
            m_Sessione = i
            Exit For
        End If
    Next i
    
    Unload Me
End Sub

Private Sub cmdEsci_Click()
    m_Sessione = tpNoneSession
    Unload Me
End Sub

