VERSION 5.00
Begin VB.Form frmTurni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Turno Dialitico"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   109
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   240
         Picture         =   "frmTurni.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblCognome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   117
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6720
         TabIndex        =   116
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblEta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   11040
         TabIndex        =   115
         Top             =   360
         Width           =   615
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
         Index           =   0
         Left            =   840
         TabIndex        =   112
         Top             =   360
         Width           =   1005
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
         Index           =   1
         Left            =   5880
         TabIndex        =   111
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anni"
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
         Left            =   10320
         TabIndex        =   110
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   87
      Top             =   720
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   240
         Picture         =   "frmTurni.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° rene"
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
         Index           =   47
         Left            =   3120
         TabIndex        =   121
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lblNumeroRene 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3960
         TabIndex        =   120
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   10800
         TabIndex        =   119
         Top             =   375
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Index           =   26
         Left            =   10200
         TabIndex        =   118
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblTipoRene 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5760
         TabIndex        =   91
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblPostazione 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2040
         TabIndex        =   90
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monitor"
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
         Index           =   18
         Left            =   4920
         TabIndex        =   89
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Postazione"
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
         Left            =   840
         TabIndex        =   88
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   120
      TabIndex        =   45
      Top             =   1440
      Width           =   12015
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   40
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   41
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   38
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   39
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   36
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   37
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   34
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   35
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   32
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   33
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   30
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   31
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   28
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioSr 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   29
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   27
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   25
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   23
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   21
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   19
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   17
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   26
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   24
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   22
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   20
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   18
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   16
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioPm 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   14
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   13
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaAOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCercaDaOrarioAm 
         Caption         =   "->"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   108
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   107
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   106
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   105
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   104
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   103
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   102
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   101
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   100
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   99
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   98
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   97
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblAOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   96
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblDaOrarioSr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   95
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Turno  SER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   21
         Left            =   120
         TabIndex        =   94
         Top             =   3105
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dalle"
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
         Index           =   20
         Left            =   1110
         TabIndex        =   93
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alle"
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
         Index           =   19
         Left            =   1110
         TabIndex        =   92
         Top             =   3495
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   11880
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   86
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   85
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   84
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   83
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   82
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   81
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblAOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   80
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   79
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   78
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   77
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   76
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   75
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   74
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDaOrarioPm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   73
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   72
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   71
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   70
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   69
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   68
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   67
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblAOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   66
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   6
         Left            =   11040
         TabIndex        =   65
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   9600
         TabIndex        =   64
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   8280
         TabIndex        =   63
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   6840
         TabIndex        =   62
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   3840
         TabIndex        =   61
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   2400
         TabIndex        =   60
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDaOrarioAm 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   5280
         TabIndex        =   59
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   11880
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alle"
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
         Index           =   17
         Left            =   1080
         TabIndex        =   58
         Top             =   2340
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dalle"
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
         Index           =   16
         Left            =   1080
         TabIndex        =   57
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Turno  POM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   1950
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Turno  MAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   14
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alle"
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
         Left            =   1110
         TabIndex        =   54
         Top             =   1245
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dalle"
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
         Left            =   1110
         TabIndex        =   53
         Top             =   765
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domenica"
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
         Left            =   10680
         TabIndex        =   52
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sabato"
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
         Left            =   9300
         TabIndex        =   51
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Venerdì"
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
         Left            =   7920
         TabIndex        =   50
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Giovedì"
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
         Left            =   6480
         TabIndex        =   49
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mercoledì"
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
         Left            =   4920
         TabIndex        =   48
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Martedì"
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
         Left            =   3480
         TabIndex        =   47
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lunedì"
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
         Left            =   2040
         TabIndex        =   46
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   44
      Top             =   5160
      Width           =   12015
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
         Left            =   10560
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
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
         Height          =   495
         Left            =   8760
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTurni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modifica As Boolean
Dim stoPulendo As Boolean
Dim rsTurni As Recordset
Dim keyId As Integer
Dim codice_rene As Integer          'il codice del rene associato
Dim intPazientiKey As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Set rsTurni = Nothing
    If intPazientiKey = 0 Then
        cmdTrova_Click (0)
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    lblPostazione.BackColor = vbWhite
    lblNumeroRene.BackColor = vbWhite
    lblTipoRene.BackColor = vbWhite
    For i = 0 To 6
        lblAOrarioAm(i).BackColor = vbWhite
        lblAOrarioPm(i).BackColor = vbWhite
        lblAOrarioSr(i).BackColor = vbWhite
        lblDaOrarioAm(i).BackColor = vbWhite
        lblDaOrarioPm(i).BackColor = vbWhite
        lblDaOrarioSr(i).BackColor = vbWhite
    Next i
    modifica = False
    codice_rene = -1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Function completoOre() As Boolean
    Dim i As Integer
    completoOre = False
    For i = 0 To 6
        If lblAOrarioAm(i) <> "" Or lblDaOrarioAm(i) <> "" Then
            If lblAOrarioAm(i) <> "" And lblDaOrarioAm(i) <> "" Then
                If CDate(lblDaOrarioAm(i)) > CDate(lblAOrarioAm(i)) Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
        If lblAOrarioPm(i) <> "" Or lblDaOrarioPm(i) <> "" Then
            If lblAOrarioPm(i) <> "" And lblDaOrarioPm(i) <> "" Then
                If CDate(lblDaOrarioPm(i)) > CDate(lblAOrarioPm(i)) Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
        If lblAOrarioSr(i) <> "" Or lblDaOrarioSr(i) <> "" Then
            If lblAOrarioSr(i) <> "" And lblDaOrarioSr(i) <> "" Then
                If CDate(lblDaOrarioSr(i)) > CDate(lblAOrarioSr(i)) Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    Next i
    completoOre = True
End Function

Private Function turniUnici() As Boolean
    ' per il controllo del turno unico (mattina o pomeriggio o sera)
    Dim i As Integer
    turniUnici = False
    For i = 0 To 6
        If ((lblAOrarioAm(i) <> "" And lblDaOrarioAm(i) <> "") And (lblAOrarioPm(i) <> "" And lblDaOrarioPm(i) <> "")) Or _
           ((lblAOrarioAm(i) <> "" And lblDaOrarioAm(i) <> "") And (lblAOrarioSr(i) <> "" And lblDaOrarioSr(i) <> "")) Or _
           ((lblAOrarioSr(i) <> "" And lblDaOrarioSr(i) <> "") And (lblAOrarioPm(i) <> "" And lblDaOrarioPm(i) <> "")) Then
            Exit Function
        End If
    Next i
    turniUnici = True
End Function

Private Function Completo() As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionere il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If Not completoOre Then
        MsgBox "Impostazione ora errata", vbCritical, "Attenzione"
        Exit Function
    End If
    If Not turniUnici Then
        MsgBox "Impossibile definire più turni nella stessa giornata", vbCritical, "Attenzione"
        Exit Function
    End If
    ' nn deve essere obbligatorio
    'If codice_rene = -1 Then
    '    MsgBox "Impostare il rene per il paziente", vbCritical, "Attenzione"
    '    Exit Function
    'End If
    Completo = True
End Function

Private Sub PulisciTutto()
    Dim i As Integer
    stoPulendo = True
    modifica = False
    codice_rene = -1
    intPazientiKey = 0
    For i = 0 To 6
        lblAOrarioAm(i) = ""
        lblAOrarioPm(i) = ""
        lblAOrarioSr(i) = ""
        lblDaOrarioAm(i) = ""
        lblDaOrarioPm(i) = ""
        lblDaOrarioSr(i) = ""
    Next i
    lblPostazione = ""
    lblNumeroRene = ""
    lblTipoRene = ""
    lblTipo = ""
    Call PulisciForm(Me)
    stoPulendo = False
    cmdTrova(0).SetFocus
    cmdMemorizza.Enabled = False
End Sub

Private Sub cmdCercaAOrarioAm_Click(Index As Integer)
    tOrario = tpMAT
    frmOrario.Show 1
    If laOra <> "" Then lblAOrarioAm(Index) = laOra
End Sub

Private Sub cmdCercaAOrarioPm_Click(Index As Integer)
    tOrario = tpPOM
    frmOrario.Show 1
    If laOra <> "" Then lblAOrarioPm(Index) = laOra
End Sub

Private Sub cmdCercaAOrarioSr_Click(Index As Integer)
    tOrario = tpSER
    frmOrario.Show 1
    If laOra <> "" Then lblAOrarioSr(Index) = laOra
End Sub

Private Sub cmdCercaDaOrarioAm_Click(Index As Integer)
    tOrario = tpMAT
    frmOrario.Show 1
    If laOra <> "" Then lblDaOrarioAm(Index) = laOra
End Sub

Private Sub cmdCercaDaOrarioPm_Click(Index As Integer)
    tOrario = tpPOM
    frmOrario.Show 1
    If laOra <> "" Then lblDaOrarioPm(Index) = laOra
End Sub

Private Sub cmdCercaDaOrarioSr_Click(Index As Integer)
    tOrario = tpSER
    frmOrario.Show 1
    If laOra <> "" Then lblDaOrarioSr(Index) = laOra
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Nomi(1 To 45) As Variant
    Dim v_Val(1 To 45) As Variant
    Dim i As Integer
    If Completo Then
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "CODICE_PAZIENTE"
        For i = 1 To 7
            v_Nomi(2 + i) = "AM_INIZIO" & i
            v_Nomi(2 + 7 + i) = "AM_FINE" & i
            v_Nomi(2 + 14 + i) = "PM_INIZIO" & i
            v_Nomi(2 + 21 + i) = "PM_FINE" & i
            v_Nomi(2 + 28 + i) = "SR_INIZIO" & i
            v_Nomi(2 + 35 + i) = "SR_FINE" & i
        Next i
        v_Nomi(45) = "CODICE_RENE"
        v_Val(1) = IIf(modifica, keyId, GetNumero("TURNI"))
        v_Val(2) = intPazientiKey
        For i = 1 To 7
            v_Val(2 + i) = lblDaOrarioAm(i - 1)
            v_Val(2 + 7 + i) = lblAOrarioAm(i - 1)
            v_Val(2 + 14 + i) = lblDaOrarioPm(i - 1)
            v_Val(2 + 21 + i) = lblAOrarioPm(i - 1)
            v_Val(2 + 28 + i) = lblDaOrarioSr(i - 1)
            v_Val(2 + 35 + i) = lblAOrarioSr(i - 1)
        Next i
        v_Val(45) = IIf(codice_rene = 0, -1, codice_rene)
        Set rsTurni = New Recordset
        If modifica Then
            rsTurni.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsTurni.Update v_Nomi, v_Val
        Else
            rsTurni.Open "TURNI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsTurni.AddNew v_Nomi, v_Val
            rsTurni.Update
        End If
        Set rsTurni = Nothing
        Call PulisciTutto
        MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
        cmdMemorizza.Enabled = False
        cmdTrova_Click (0)
    End If
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    Dim i As Integer
    If intPazientiKey = 0 Then
        Exit Sub
    Else
        cmdMemorizza.Enabled = True
    End If
    ' carica i dati del paziente
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognome = rsDataset("COGNOME")
    lblNome = rsDataset("NOME")
    Dim somma As Integer
    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
    Set rsDataset = Nothing
    ' cerca i riferimenti al paziente indifferentemente dalla data
    ' infatti la data è un campo informativo (passivo)
    Set rsTurni = New Recordset
    rsTurni.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsTurni.EOF And rsTurni.BOF Then
        ' nn ha trovato nessun record
        modifica = False
    Else
        keyId = rsTurni("KEY")
        modifica = True
        For i = 1 To 7
            lblAOrarioAm(i - 1) = rsTurni("AM_FINE" & i) & ""
            lblAOrarioPm(i - 1) = rsTurni("PM_FINE" & i) & ""
            lblAOrarioSr(i - 1) = rsTurni("SR_FINE" & i) & ""
            lblDaOrarioAm(i - 1) = rsTurni("AM_INIZIO" & i) & ""
            lblDaOrarioPm(i - 1) = rsTurni("PM_INIZIO" & i) & ""
            lblDaOrarioSr(i - 1) = rsTurni("SR_INIZIO" & i) & ""
        Next i
        codice_rene = rsTurni("CODICE_RENE")
        Call CaricaRene
    End If
    Set rsTurni = Nothing
End Sub

Private Sub CaricaRene()
    Dim rsDataset As New Recordset
    If codice_rene = -1 Then Exit Sub
    rsDataset.Open "SELECT * FROM APPARATI WHERE KEY=" & codice_rene, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
    Else
        lblPostazione = rsDataset("POSTAZIONE")
        lblNumeroRene = rsDataset("NUMERO_APPARATO") & ""
        lblTipoRene = rsDataset("MODELLO")
        lblTipo = Choose(rsDataset("TIPO") + 1, "NEG", "HCV POS", "HBV POS")
    End If
    Set rsDataset = Nothing
End Sub

Private Sub lblAOrarioAm_Click(Index As Integer)
    laOra = ""
    lblAOrarioAm(Index) = ""
End Sub

Private Sub lblAOrarioPm_Click(Index As Integer)
    laOra = ""
    lblAOrarioPm(Index) = ""
End Sub

Private Sub lblAOrarioSr_Click(Index As Integer)
    laOra = ""
    lblAOrarioSr(Index) = ""
End Sub

Private Sub lblDaOrarioAm_Click(Index As Integer)
    laOra = ""
    lblDaOrarioAm(Index) = ""
End Sub

Private Sub lblDaOrarioPm_Click(Index As Integer)
    laOra = ""
    lblDaOrarioPm(Index) = ""
End Sub

Private Sub lblDaOrarioSr_Click(Index As Integer)
    laOra = ""
    lblDaOrarioSr(Index) = ""
End Sub

Private Function CreaCondizione() As String
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "SELECT    DATA_INIZIO, DATA_FINE, PAZIENTI.KEY " & _
             "FROM      (ANAMNESI_NEFROLOGICHE " & _
             "          INNER JOIN PAZIENTI ON PAZIENTI.KEY=ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE) " & _
             "WHERE     STATO=4"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        If Not IsNull(rsDataset("DATA_INIZIO")) Then
            If CDate(rsDataset("DATA_INIZIO")) <= date Then
                If rsDataset("DATA_FINE") <> "" Then
                    If CDate(rsDataset("DATA_FINE")) >= date Then
                        CreaCondizione = CreaCondizione & rsDataset("KEY") & ","
                    End If
                Else
                    CreaCondizione = CreaCondizione & rsDataset("KEY") & ","
                End If
            End If
        End If
        rsDataset.MoveNext
    Loop
    If CreaCondizione <> "" Then
        ' elimina la , finale e aggiunge le parentesi
        CreaCondizione = Left(CreaCondizione, Len(CreaCondizione) - 1)
        CreaCondizione = " KEY IN (" & CreaCondizione & ")"
    Else
        ' non deve trovare nessun paziente (key=-1 piezzo)
        CreaCondizione = " KEY IN (-1)"
    End If
    CreaCondizione = "( " & CreaCondizione & " OR STATO=0)"
    
    Set rsDataset = Nothing
End Function

Private Sub cmdTrova_Click(Index As Integer)
    If Index = 0 Then
        ' pulisce per evitare problemi
        Call PulisciTutto
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = CreaCondizione
        tTrova.condStato = "(-1)"
        frmTrova.Show 1
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
    Else
        frmVisualizzaReni.Show 1
        If tReni.postazione <> Str(-1) Then
            codice_rene = tReni.key
            lblPostazione = tReni.postazione
            lblNumeroRene = tReni.numero_apparato
            lblTipoRene = tReni.monitor
            lblTipo = tReni.Tipo
        End If
    End If
    
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
    
End Sub
