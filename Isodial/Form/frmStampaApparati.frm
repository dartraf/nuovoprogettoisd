VERSION 5.00
Begin VB.Form frmStampaApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Apparati "
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStampa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CheckBox chkIncludiApparatiRottamati 
         Caption         =   "Includi apparati rottamati"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optApparatiRottamati 
         Caption         =   "Elenco apparati rottamati"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton optTipoApparato 
         Caption         =   "Elenco per tipo apparato"
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
         TabIndex        =   6
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton optProduttore 
         Caption         =   "Elenco per produttore"
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
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optInventario 
         Caption         =   "Elenco per n° d'inventario"
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
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5415
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
         Left            =   4080
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
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmStampaApparati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAvanti_Click()

    If optInventario.Value = unchekd And optProduttore.Value = unchekd And optTipoApparato.Value = unchekd And optApparatiRottamati.Value = Unchecked Then
        MsgBox "Selezionare il tipo di elenco da stampare", vbInformation, "Informazione"
        Exit Sub
    End If
    
End Sub

Private Sub cmdEsci_Click()
    Unload frmStampaApparati
End Sub
