VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su ISODIAL"
   ClientHeight    =   4344
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5736
   ClipControls    =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2992.094
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3825
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEGUICI SU:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblFacebook 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facebook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2280
      MouseIcon       =   "frmInfo.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Seguici su FACEBOOK"
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Image Image4 
      Height          =   204
      Left            =   4800
      Picture         =   "frmInfo.frx":06DC
      Top             =   1800
      Width           =   192
   End
   Begin VB.Image Image3 
      Height          =   192
      Left            =   4800
      Picture         =   "frmInfo.frx":09D5
      Top             =   2880
      Width           =   192
   End
   Begin VB.Image Image2 
      Height          =   192
      Left            =   4800
      Picture         =   "frmInfo.frx":0F5F
      Top             =   720
      Width           =   192
   End
   Begin VB.Label lblposta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail: info@isodial.it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1680
      MouseIcon       =   "frmInfo.frx":14E9
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Indirizzo di POSTA ELETTRONICA"
      Top             =   2880
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POSTA ELETTRONICA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WEB:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblsito 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.isodial.it "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   288
      Left            =   2160
      MouseIcon       =   "frmInfo.frx":163B
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Visita il sito!"
      Top             =   720
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   240
      Picture         =   "frmInfo.frx":178D
      Top             =   240
      Width           =   384
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.426
      X2              =   5309.473
      Y1              =   2484.457
      Y2              =   2484.457
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblsito.FontUnderline = False
    lblposta.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblsito.FontUnderline = False
    lblposta.FontUnderline = False
End Sub

Private Function Link(ByVal URL As String) As Long
   Link = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub lblFacebook_Click()
    Call Link("www.facebook.com/groups/160805834018168/")
End Sub

Private Sub lblFacebook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblFacebook.FontUnderline = True
End Sub

Private Sub lblposta_Click()
    Call Link("mailto:info@isodial.it")
End Sub

Private Sub lblposta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblposta.FontUnderline = True
End Sub

Private Sub lblsito_Click()
    Call Link("http://www.isodial.it ")
End Sub

Private Sub lblsito_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblsito.FontUnderline = True
End Sub
