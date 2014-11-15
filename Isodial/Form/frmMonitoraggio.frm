VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitoraggio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monitoraggio"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   0
         Left            =   360
         Picture         =   "frmMonitoraggio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Seleziona il paziente"
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
         Left            =   2160
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   360
         Width           =   615
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
         Left            =   10440
         TabIndex        =   41
         Top             =   360
         Width           =   465
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
         Left            =   6000
         TabIndex        =   40
         Top             =   360
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
         Index           =   0
         Left            =   1080
         TabIndex        =   39
         Top             =   360
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab tabSchede 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   850
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Attuazione protocollo di vaccinazione epatite"
      TabPicture(0)   =   "frmMonitoraggio.frx":0459
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblData(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboEsito(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNote(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "picElenca(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "picData(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Attuazione protocollo di monitoraggio accessi vascolari"
      TabPicture(1)   =   "frmMonitoraggio.frx":0475
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblData(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(7)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "picData(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "picElenca(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboEsito(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtNote(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Valutazione psico-sociale"
      TabPicture(2)   =   "frmMonitoraggio.frx":0491
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(15)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(16)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblData(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(14)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(17)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(18)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(19)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblNomePsicologo"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblCognomePsicologo"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboEsito(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtNote(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "picData(1)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "picElenca(1)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmdTrova(1)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Collegamenti funzionali tra nefrologo e medici di base"
      TabPicture(3)   =   "frmMonitoraggio.frx":04AD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(42)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(41)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(40)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label1(35)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label1(33)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label1(32)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label1(30)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label1(31)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblData(3)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label1(34)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblCognomeMedico"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "lblNomeMedico"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "lblIndirizzo"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "lblTelefono"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "lblStudio"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "lblCellulare"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label1(8)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "lblDataSchedaPaziente(4)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "picData(2)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "picData(3)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "cmdTrova(2)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "txtNote(3)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).ControlCount=   22
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   4
         Left            =   2400
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   62
         ToolTipText     =   "Cerca data"
         Top             =   1065
         Width           =   360
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   2
         Left            =   2880
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   61
         ToolTipText     =   "Elenca date"
         Top             =   1065
         Width           =   360
      End
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   3
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   3240
         Width           =   11535
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   2
         Left            =   -73560
         Picture         =   "frmMonitoraggio.frx":04C9
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Seleziona il medico di base"
         Top             =   960
         Width           =   450
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   1
         Left            =   -73440
         Picture         =   "frmMonitoraggio.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Seleziona lo psicologo"
         Top             =   1560
         Width           =   450
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   3
         Left            =   -64680
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   25
         ToolTipText     =   "Cerca data"
         Top             =   2490
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   -71160
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   24
         ToolTipText     =   "Cerca data"
         Top             =   2490
         Width           =   360
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   -71520
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   16
         ToolTipText     =   "Elenca date"
         Top             =   1035
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   -72000
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   15
         ToolTipText     =   "Cerca data"
         Top             =   1035
         Width           =   360
      End
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Index           =   2
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2970
         Width           =   11535
      End
      Begin VB.ComboBox cboEsito 
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
         Index           =   2
         ItemData        =   "frmMonitoraggio.frx":0D7B
         Left            =   -73440
         List            =   "frmMonitoraggio.frx":0D85
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2520
         Width           =   11535
      End
      Begin VB.ComboBox cboEsito 
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
         ItemData        =   "frmMonitoraggio.frx":0D9D
         Left            =   960
         List            =   "frmMonitoraggio.frx":0DA7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtNote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Index           =   1
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2520
         Width           =   11535
      End
      Begin VB.ComboBox cboEsito 
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
         ItemData        =   "frmMonitoraggio.frx":0DBF
         Left            =   -74040
         List            =   "frmMonitoraggio.frx":0DD8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   3615
      End
      Begin VB.PictureBox picElenca 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   -72120
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   2
         ToolTipText     =   "Elenca date"
         Top             =   1065
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   -72600
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   1
         ToolTipText     =   "Cerca data"
         Top             =   1065
         Width           =   360
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
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   65
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblDataSchedaPaziente 
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
         Left            =   -72480
         TabIndex        =   64
         Top             =   2520
         Width           =   1215
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
         Index           =   9
         Left            =   240
         TabIndex        =   63
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note sui collegamenti funzionali tra nefrologo e medici di base"
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
         Left            =   -74760
         TabIndex        =   60
         Top             =   3000
         Width           =   6465
      End
      Begin VB.Label lblCellulare 
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
         Left            =   -67320
         TabIndex        =   52
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label lblStudio 
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
         Left            =   -67320
         TabIndex        =   51
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label lblTelefono 
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
         Left            =   -72840
         TabIndex        =   50
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label lblIndirizzo 
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
         Left            =   -72840
         TabIndex        =   49
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label lblNomeMedico 
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
         Left            =   -67320
         TabIndex        =   48
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblCognomeMedico 
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
         Left            =   -72840
         TabIndex        =   47
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblCognomePsicologo 
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
         Left            =   -71520
         TabIndex        =   46
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label lblNomePsicologo 
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
         Left            =   -66600
         TabIndex        =   45
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Riunione periodica del"
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
         Left            =   -68520
         TabIndex        =   35
         Top             =   2540
         Width           =   2370
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
         Height          =   285
         Index           =   3
         Left            =   -66000
         TabIndex        =   34
         Top             =   2535
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Scheda paziente del"
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
         Index           =   31
         Left            =   -74760
         TabIndex        =   33
         Top             =   2540
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dati medico di base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   -74760
         TabIndex        =   32
         Top             =   600
         Width           =   2100
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
         Left            =   -74760
         TabIndex        =   31
         Top             =   1080
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
         Index           =   33
         Left            =   -68520
         TabIndex        =   30
         Top             =   1080
         Width           =   630
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   1560
         Width           =   870
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
         Left            =   -68520
         TabIndex        =   28
         Top             =   2040
         Width           =   945
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   2040
         Width           =   945
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
         Left            =   -68520
         TabIndex        =   26
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   2310
         Width           =   540
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
         Index           =   18
         Left            =   -72600
         TabIndex        =   22
         Top             =   1680
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
         Index           =   17
         Left            =   -67320
         TabIndex        =   21
         Top             =   1680
         Width           =   630
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   20
         Top             =   1080
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
         Index           =   1
         Left            =   -73440
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note sulla valutazione psico-sociale"
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
         Left            =   -74760
         TabIndex        =   18
         Top             =   2760
         Width           =   3780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Psicologo"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   17
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note sulla vaccinazione"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
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
         TabIndex        =   11
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note sul monitoraggio dell'accesso vascolare"
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
         Left            =   -74760
         TabIndex        =   10
         Top             =   2280
         Width           =   4785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Esito"
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
         Left            =   -74760
         TabIndex        =   9
         Top             =   1695
         Width           =   540
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
         Index           =   7
         Left            =   -74760
         TabIndex        =   8
         Top             =   1080
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
         Index           =   0
         Left            =   -74040
         TabIndex        =   7
         Top             =   1110
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   54
      Top             =   5040
      Width           =   12015
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Left            =   5040
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   7200
         TabIndex        =   57
         Top             =   240
         Width           =   1455
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
         Left            =   8880
         TabIndex        =   56
         Top             =   240
         Width           =   1575
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
         Left            =   10680
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   10800
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMonitoraggio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmMonitoraggio.frm
'
' <b>Descrizione</b>: Scheda Monitoraggi associata alle tab MON_...
'
' @remarks
'
' @author
'
' @date 21/02/2011 19.48
Option Explicit

'' rs della scheda
Dim rsMonitoraggi As Recordset
'' indica per ogni scheda se si è in modifica e il relativo key
Private Type structModifica
    v_modifica(1 To 4) As Boolean
    v_numKey(1 To 4) As Integer
End Type
Dim modifica As structModifica
Dim intPazientiKey As Integer
Dim intMedicoKey(1) As Integer

Dim rsDialisi As Recordset

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click (0)
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    For i = 0 To 3
        lblData(i).BackColor = vbWhite
    Next i
    For i = 0 To 4
        picData(i).Picture = LoadResPicture("cal1", 0)
    Next i
    lblDataSchedaPaziente(4).BackColor = vbWhite
    picElenca(0).Picture = LoadResPicture("elenca1", 0)
    picElenca(1).Picture = LoadResPicture("elenca1", 0)
    picElenca(2).Picture = LoadResPicture("elenca1", 0)
    tabSchede.Tab = 0
    Call EliminaScansioniSospese("SCAN_PSICO_SOCIALE")
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    intPazientiKey = 0
    intMedicoKey(0) = 0
    intMedicoKey(1) = 0
    lblCognomeMedico = ""
    lblNomeMedico = ""
    lblIndirizzo = ""
    lblStudio = ""
    lblTelefono = ""
    lblCellulare = ""
    For i = 0 To 3
        lblData(i) = ""
    Next i
    lblDataSchedaPaziente(4) = ""
    For i = 1 To 4
        modifica.v_modifica(i) = False
        modifica.v_numKey(i) = 0
    Next i
    Call PulisciForm(Me)
    Call EliminaScansioniSospese("SCAN_PSICO_SOCIALE")
End Sub

'' Carica la scheda nel form
'
' @param index indice della scheda da caricare
Private Sub CaricaScheda(Index As Integer)
    Dim data As Date
    Dim nomeTabella As String
    Select Case Index
        Case 0
            nomeTabella = "MON_ACCESSI"
        Case 1
            nomeTabella = "MON_VALUTAZIONI"
        Case 2
            nomeTabella = "MON_VACC_EPATITE"
    End Select
    If intPazientiKey <> 0 And lblData(Index) <> "" Then
        Set rsMonitoraggi = New Recordset
        data = DateValue(Month(lblData(Index)) & "/" & Day(lblData(Index)) & "/" & Year(lblData(Index)))
        rsMonitoraggi.Open "SELECT * FROM " & nomeTabella & " WHERE (CODICE_PAZIENTE=" & intPazientiKey & ") AND (DATA=#" & data & "#)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsMonitoraggi.EOF And rsMonitoraggi.BOF) Then
            Select Case Index
                Case 0
                    cboEsito(1).ListIndex = rsMonitoraggi("ESITO")
                    txtNote(1) = rsMonitoraggi("NOTE") & ""
                    modifica.v_modifica(2) = True
                    modifica.v_numKey(2) = rsMonitoraggi("KEY")
                Case 1
                    cboEsito(2).ListIndex = rsMonitoraggi("ESITO")
                    txtNote(2) = rsMonitoraggi("NOTE") & ""
                    intMedicoKey(0) = rsMonitoraggi("CODICE_PSICOLOGO")
                    Call CaricaMedico(0)
                    modifica.v_modifica(3) = True
                    modifica.v_numKey(3) = rsMonitoraggi("KEY")
                Case 2
                    lblData(2).Caption = rsMonitoraggi("DATA")
                    cboEsito(0).ListIndex = rsMonitoraggi("ESITO")
                    txtNote(0) = rsMonitoraggi("NOTE") & ""
                    modifica.v_modifica(1) = True
                    modifica.v_numKey(1) = rsMonitoraggi("KEY")
            End Select
        Else
            modifica.v_modifica(1 + Index) = False
            modifica.v_numKey(1 + Index) = 0
        End If
        Set rsMonitoraggi = Nothing
    End If
End Sub

'' Pulisca la scheda ma non il paziente
Private Sub Pulisci(Index As Integer)
    Select Case Index
        Case 0
            cboEsito(1).ListIndex = -1
            txtNote(1) = ""
            modifica.v_modifica(2) = False
            modifica.v_numKey(2) = 0
        Case 1
            cboEsito(2).ListIndex = -1
            txtNote(2) = ""
            intMedicoKey(0) = 0
            lblCognomePsicologo = ""
            lblNomePsicologo = ""
            modifica.v_modifica(3) = False
            modifica.v_numKey(3) = 0
        Case 2
            cboEsito(0).ListIndex = -1
            txtNote(0) = ""
            modifica.v_modifica(1) = False
            modifica.v_numKey(1) = 0
    End Select
End Sub

'' Salva le date nel db solo per l'ultima scheda
Private Sub GestisciDate()
    With rsMonitoraggi
        .Fields("DATA") = Null
        .Fields("DATA_RIUNIONE") = Null
        If lblDataSchedaPaziente(4) <> "" Then
            .Fields("DATA") = lblDataSchedaPaziente(4)
        End If
        If lblData(3) <> "" Then
            .Fields("DATA_RIUNIONE") = lblData(3)
        End If
    End With
End Sub

Private Sub cmdChiudi_Click()
    Call PulisciTutto
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim numScheda As Byte
    Dim nomeTabella As String
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim numKey As Integer
    Dim nomeFile As String
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If
    
    If tabSchede.Tab = 0 Then
        If lblData(2).Caption = "" Then
            MsgBox "Selezionare la data", vbInformation, "Attenzione"
            Exit Sub
        End If
    End If
    
    If tabSchede.Tab = 1 Then
        If lblData(0).Caption = "" Then
            MsgBox "Selezionare la data", vbInformation, "Attenzione"
            Exit Sub
        End If
    End If
    
    If tabSchede.Tab = 2 Then
        If lblData(1).Caption = "" Then
            MsgBox "Selezionare la data", vbInformation, "Attenzione"
            Exit Sub
        End If
    End If
    
    numScheda = tabSchede.Tab + 1       '(1 la prima scheda)
    If numScheda = 2 Or numScheda = 3 Then
        If lblData(numScheda - 2) = "" Then
            Exit Sub
        End If
    End If
    
    If numScheda <> 4 Then
        If cboEsito(numScheda - 1).ListIndex = -1 Then
            MsgBox "Selezionare un esito", vbInformation, "Attenzione"
            Exit Sub
        End If
    End If
    
    Select Case numScheda
        Case 1: nomeTabella = "MON_VACC_EPATITE"
        Case 2: nomeTabella = "MON_ACCESSI"
        Case 3: nomeTabella = "MON_VALUTAZIONI"
        Case 4: nomeTabella = "MON_COLLEGAMENTI"
    End Select
    
    If modifica.v_modifica(numScheda) Then
        numKey = modifica.v_numKey(numScheda)
    Else
        numKey = GetNumero(nomeTabella)
    End If
    Select Case numScheda
        Case 1
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "ESITO", "NOTE")
            v_Val = Array(numKey, intPazientiKey, lblData(2), cboEsito(0).ListIndex, CStr(txtNote(0) & ""))
        Case 2
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "ESITO", "NOTE")
            v_Val = Array(numKey, intPazientiKey, lblData(0), cboEsito(1).ListIndex, txtNote(1) & "")
        Case 3
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "DATA", "NOTE", "CODICE_PSICOLOGO", "ESITO")
            v_Val = Array(numKey, intPazientiKey, lblData(1), txtNote(2) & "", intMedicoKey(0), cboEsito(2).ListIndex)
        Case 4
            v_Nomi = Array("KEY", "CODICE_PAZIENTE", "CODICE_MEDICO", "NOTE_COLLEGAMENTI")
            v_Val = Array(numKey, intPazientiKey, intMedicoKey(1), txtNote(3))
    End Select
    
    Set rsMonitoraggi = New Recordset
    If modifica.v_modifica(numScheda) Then
        rsMonitoraggi.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & modifica.v_numKey(numScheda), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If numScheda = 4 Then
            Call GestisciDate
        End If
        rsMonitoraggi.Update v_Nomi, v_Val
    Else
        rsMonitoraggi.Open nomeTabella, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsMonitoraggi.AddNew v_Nomi, v_Val
        If numScheda = 4 Then
            Call GestisciDate
        End If
        rsMonitoraggi.Update
        rsMonitoraggi.Close
        If numScheda = 3 Then
            ' controlla eventuali scansioni memorizzate in sospeso
            rsMonitoraggi.Open "SELECT * FROM SCAN_PSICO_SOCIALE WHERE CODICE_SCHEDA=0", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            Do While Not rsMonitoraggi.EOF
                rsMonitoraggi("CODICE_SCHEDA") = numKey
                nomeFile = rsMonitoraggi("NOME_FILE")
                rsMonitoraggi("NOME_FILE") = M_PS & numKey & " " & Replace(lblData(1), "/", "-") & Right(nomeFile, 2)
                rsMonitoraggi.Update
                rsMonitoraggi.MoveNext
                If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".jpg" As structApri.pathDB & "\" & M_PS & numKey & " " & Replace(lblData(1), "/", "-") & Right(nomeFile, 2) & ".jpg"
                ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".pdf" As structApri.pathDB & "\" & M_PS & numKey & " " & Replace(lblData(1), "/", "-") & Right(nomeFile, 2) & ".pdf"
                End If
            Loop
            rsMonitoraggi.Close
        End If
    End If
    Set rsMonitoraggi = Nothing
    
    Call PulisciTutto
    MsgBox "Salvataggio effettuato" & vbCrLf & "Scheda: " & tabSchede.TabCaption(numScheda - 1), vbInformation, "Salvataggio"
    cmdTrova_Click (0)
End Sub

Private Sub cmdTrova_Click(Index As Integer)
    Select Case Index
        Case 0
            ' pulisce per evitare problemi
            Call PulisciTutto
            tTrova.Tipo = tpPAZIENTE
        Case 1
            tTrova.Tipo = tpPSICOLOGI
        Case 2
            tTrova.Tipo = tpMEDICOBASE
    End Select
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    Select Case Index
        Case 0
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        Case Else
            intMedicoKey(Index - 1) = tTrova.keyReturn
            Call CaricaMedico(Index - 1)
    End Select
End Sub

Private Sub cmdGestioneReferti_Click()
    Unload frmGestioneDocumentiEsterni
    Load frmGestioneDocumentiEsterni
    frmGestioneDocumentiEsterni.LetCodicePaziente = intPazientiKey
    If modifica.v_modifica(3) Then
        frmGestioneDocumentiEsterni.letcodiceRecord = modifica.v_numKey(3)
        frmGestioneDocumentiEsterni.LetNomeFile = M_PS & modifica.v_numKey(3) & " " & Replace(lblData(1), "/", "-")
    Else
        frmGestioneDocumentiEsterni.letcodiceRecord = 0
        frmGestioneDocumentiEsterni.LetNomeFile = M_PS & 0 & " " & Replace(lblData(1), "/", "-")
    End If
    tDocumentiEsterni = tpSCANMONITORAGGIO
    frmGestioneDocumentiEsterni.Show 1
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If
      
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDialisi("COGNOME") & " " & rsDialisi("NOME")
    structIntestazione.sDataPaziente = rsDialisi("DATA_NASCITA")
    codiceId = rsDialisi("CODICE_ID")
    Set rsDialisi = Nothing

    strSql = "SHAPE APPEND  NEW adVarChar(15)  AS ESITO_PROTOCOLLO_VACCINAZIONE, " & _
                    "       NEW adVarChar(15)  AS DATA_PROTOCOLLO_VACCINAZIONE, " & _
                    "       NEW adLongVarChar  AS NOTE_VACCINAZIONE, "
    strSql = strSql & _
                    "       NEW adVarChar(15)  AS DATA_PROTOCOLLO_MONITORAGGIO, " & _
                    "       NEW adVarChar(40)  AS ESITO_PROTOCOLLO_MONITORAGGIO, " & _
                    "       NEW adLongVarChar  AS NOTE_PROTOCOLLO_MONITORAGGIO, "
    strSql = strSql & _
                    "       NEW adVarChar(15)  AS DATA_VALUTAZIONE_PSICOSOCIALE, " & _
                    "       NEW adVarChar(15)  AS ESITO_VALUTAZIONE_PSICOSOCIALE, " & _
                    "       NEW adVarChar(30)  AS PSICOLOGO, " & _
                    "       NEW adLongVarChar  AS NOTE_VALUTAZIONE_PSICOSOCIALE, "
    strSql = strSql & _
                    "       NEW adVarChar(20)  AS COGNOME_MEDICO, " & _
                    "       NEW adVarChar(20)  AS NOME_MEDICO, " & _
                    "       NEW adVarChar(30)  AS INDIRIZZO_MEDICO, " & _
                    "       NEW adVarChar(30)  AS STUDIO_MEDICO, " & _
                    "       NEW adVarChar(15)  AS TELEFONO_MEDICO, " & _
                    "       NEW adVarChar(15)  AS CELLULARE_MEDICO, " & _
                    "       NEW adVarChar(15)  AS DATA_SCHEDA_PAZIENTE, " & _
                    "       NEW adVarChar(15)  AS RIUNIONE_PERIODICA, " & _
                    "       NEW adLongVarChar  AS NOTE_COLLEGAMENTI "
                 
                 
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Set rsDialisi = New Recordset
        
    With rsMain
        .AddNew
                            
                            'Attuazione protocollo di vaccinazione epatite
                            
        .Fields("DATA_PROTOCOLLO_VACCINAZIONE") = lblData(2)
        .Fields("ESITO_PROTOCOLLO_VACCINAZIONE") = cboEsito(0)
        .Fields("NOTE_VACCINAZIONE") = txtNote(0)
            
                            'Attuazione protocollo di monitoraggio acc. vascolari
        
        .Fields("DATA_PROTOCOLLO_MONITORAGGIO") = lblData(0)
        .Fields("ESITO_PROTOCOLLO_MONITORAGGIO") = cboEsito(1)
        .Fields("NOTE_PROTOCOLLO_MONITORAGGIO") = txtNote(1)
        
                            'Valutazione psico-sociale
        
        .Fields("DATA_VALUTAZIONE_PSICOSOCIALE") = lblData(1)
        .Fields("ESITO_VALUTAZIONE_PSICOSOCIALE") = cboEsito(2)
        .Fields("PSICOLOGO") = lblCognomePsicologo & " " & lblNomePsicologo
        .Fields("NOTE_VALUTAZIONE_PSICOSOCIALE") = txtNote(2)
        
                            'Collegamenti funzionali tra nefrologo e medici di base
        
        .Fields("COGNOME_MEDICO") = lblCognomeMedico
        .Fields("NOME_MEDICO") = lblNomeMedico
        .Fields("INDIRIZZO_MEDICO") = lblIndirizzo
        .Fields("STUDIO_MEDICO") = lblStudio
        .Fields("TELEFONO_MEDICO") = lblTelefono
        .Fields("CELLULARE_MEDICO") = lblCellulare
        .Fields("DATA_SCHEDA_PAZIENTE") = lblDataSchedaPaziente(4)
        .Fields("RIUNIONE_PERIODICA") = lblData(3)
        .Fields("NOTE_COLLEGAMENTI") = txtNote(3)
        
    End With
                               
    Set rptMonitoraggio.DataSource = rsMain
    rptMonitoraggio.TopMargin = 0
    rptMonitoraggio.BottomMargin = 0
    rptMonitoraggio.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptMonitoraggio.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptMonitoraggio.PrintReport True, rptRangeAllPages
    Set rsDialisi = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EliminaScansioniSospese("SCAN_PSICO_SOCIALE")
End Sub

'' Carica i dati del medico
Private Sub CaricaMedico(Index As Integer)
    Dim rsDataset As Recordset
    If intMedicoKey(Index) = 0 Then Exit Sub
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & IIf(Index = 0, "PSICOLOGI", "MEDICI_BASE") & " WHERE KEY=" & intMedicoKey(Index), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Index = 0 Then
        lblCognomePsicologo = rsDataset("COGNOME") & ""
        lblNomePsicologo = rsDataset("NOME") & ""
    Else
        lblCognomeMedico = rsDataset("COGNOME") & ""
        lblNomeMedico = rsDataset("NOME") & ""
        lblIndirizzo = rsDataset("INDIRIZZO") & ""
        lblStudio = rsDataset("STUDIO") & ""
        lblTelefono = rsDataset("TELEFONO") & ""
        lblCellulare = rsDataset("CELLULARE") & ""
    End If
    Set rsDataset = Nothing
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        Exit Sub
    Else
        cmdMemorizza.Enabled = True
    End If
    
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
    rsDataset.Close

    rsDataset.Open "SELECT * FROM MON_COLLEGAMENTI WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        intMedicoKey(1) = rsDataset("CODICE_MEDICO")
        Call CaricaMedico(1)
        lblDataSchedaPaziente(4) = rsDataset("DATA") & ""
        lblData(3) = rsDataset("DATA_RIUNIONE") & ""
        txtNote(3) = rsDataset("NOTE_COLLEGAMENTI") & ""
        modifica.v_modifica(4) = True
        modifica.v_numKey(4) = rsDataset("KEY")
    Else
        modifica.v_modifica(4) = False
        modifica.v_numKey(4) = 0
    End If
    Set rsDataset = Nothing
End Sub

Private Sub lblData_Change(Index As Integer)
    Call Pulisci(Index)

    If Index = 0 Or Index = 1 Or Index = 2 Then
        If lblData(Index) <> "" Then
            Call CaricaScheda(Index)
        End If
    End If
    
End Sub

Private Sub lblData_Click(Index As Integer)
    lblData(Index) = ""
    laData = ""
End Sub

Private Sub lblDataSchedaPaziente_Click(Index As Integer)
    lblDataSchedaPaziente(4).Caption = ""
    laData = ""
End Sub

Private Sub picData_Click(Index As Integer)
    frmCalendario.Show 1
    If Index = 4 Then
        Index = 2
    End If
    If tabSchede.Tab = 3 Then
        If laData <> "" Then lblDataSchedaPaziente(4) = laData
    End If
    If laData <> "" Then lblData(Index) = laData
End Sub

Private Sub picData_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal1", 0)
End Sub

Private Sub picElenca_Click(Index As Integer)
    ' sfrutto il fatto che sono consecutivi
    tElenca.Tipo = tpMON_ACC_VASCOLARE + Index
    tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey
    frmElencaDate.Show 1
    If laData <> "" Then lblData(Index) = laData
End Sub

Private Sub picElenca_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picElenca(Index).Picture = LoadResPicture("elenca2", 0)
End Sub

Private Sub picElenca_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picElenca(Index).Picture = LoadResPicture("elenca1", 0)
End Sub

Private Sub tabSchede_Click(PreviousTab As Integer)
    If tabSchede.Tab = 2 Then
        cmdGestioneReferti.Visible = True
    Else
        cmdGestioneReferti.Visible = False
    End If
End Sub

Private Sub txtNote_GotFocus(Index As Integer)
    txtNote(Index).BackColor = colArancione
End Sub

Private Sub txtNote_LostFocus(Index As Integer)
    txtNote(Index).BackColor = vbWhite
End Sub

