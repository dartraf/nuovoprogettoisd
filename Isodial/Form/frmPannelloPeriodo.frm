VERSION 5.00
Begin VB.Form frmPannelloPeriodo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selezione Turno"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2280
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   4
         ToolTipText     =   "Cerca data"
         Top             =   1320
         Width           =   360
      End
      Begin VB.OptionButton optTempo 
         Caption         =   "&Sera"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optTempo 
         Caption         =   "&Pomeriggio"
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optTempo 
         Caption         =   "&Mattina"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2160
         Picture         =   "frmPannelloPeriodo.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1365
         Width           =   510
      End
      Begin VB.Label lblData 
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
         Left            =   840
         TabIndex        =   9
         Top             =   1365
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
      Begin VB.CommandButton cmdConferma 
         Cancel          =   -1  'True
         Caption         =   "&Conferma"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Annulla"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmPannelloPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private periodo As Integer               ' 1 matt 2 pome 3 annulla
Private data As Date
Private SenzaData As Boolean


Public Property Get GetSenzaData() As Boolean
    GetSenzaData = SenzaData
End Property

Public Property Let LetSenzaData(ByVal VSenzaData As Boolean)
    SenzaData = VSenzaData
End Property

Public Property Get GetPeriodo() As Integer
    GetPeriodo = periodo
End Property

Public Property Let LetPeriodo(ByVal Vperiodo As Integer)
    periodo = Vperiodo
End Property

Public Property Get getData() As Date
    getData = data
End Property

Public Property Let LetData(ByVal vdata As Date)
    data = vdata
End Property


Private Sub lblData_Click()
    lblData = date
End Sub

Private Sub picData_Click()
    frmCalendario.Show 1
    If laData <> "" Then lblData = laData
End Sub

Private Sub picData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData.Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub cmdAnnulla_Click()
    periodo = -1     ' annulla
    data = date
    Unload Me
End Sub

Private Sub cmdConferma_Click()
    If optTempo(0).Value Then
        periodo = 1 ' matt
    ElseIf optTempo(1).Value Then
        periodo = 2  ' pom
    Else
        periodo = 3 ' sera
    End If
    data = lblData
    If data > date Then
        MsgBox "Impossibile selezionare una data successiva a quella odierna", vbCritical, "Attenzione"
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    If Int(Hour(Now)) < 13 Then
        optTempo(0).Value = True
    ElseIf Int(Hour(Now)) > 12 And Int(Hour(Now)) < 18 Then
        optTempo(1).Value = True
    Else
        optTempo(2).Value = True
    End If
    periodo = -1     ' annulla
    lblData = date
    picData.Picture = LoadResPicture("cal1", 0)
    lblData.BackColor = vbWhite
    If SenzaData Then
        Frame1.Height = 1335
        Frame2.Top = 1200
        Me.Height = 2490
        picData.Visible = False
    End If
End Sub

