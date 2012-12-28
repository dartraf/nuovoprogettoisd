VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDisconnetti 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Disconnessione"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSpegni 
      BackColor       =   &H00808080&
      Caption         =   "Spegni il pc al termine del backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   3855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":0682
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":1386
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":1A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisconnetti.frx":208A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAnnulla 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3600
      MouseIcon       =   "frmDisconnetti.frx":270C
      MousePointer    =   99  'Custom
      Picture         =   "frmDisconnetti.frx":2A16
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   7
      ToolTipText     =   "Annulla l' operazione"
      Top             =   960
      Width           =   360
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5655
      TabIndex        =   5
      Top             =   2040
      Width           =   5655
      Begin VB.Line Line5 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   4760
         Y1              =   700
         Y2              =   700
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5655
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   960
         Picture         =   "frmDisconnetti.frx":2BA0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   4560
         Y1              =   10
         Y2              =   10
      End
   End
   Begin VB.PictureBox picDisconnetti 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   600
      MouseIcon       =   "frmDisconnetti.frx":2EAA
      MousePointer    =   99  'Custom
      Picture         =   "frmDisconnetti.frx":31B4
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   1
      ToolTipText     =   "Consente l'accesso ad un altro utente"
      Top             =   960
      Width           =   360
   End
   Begin VB.PictureBox picChiudi 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2160
      MouseIcon       =   "frmDisconnetti.frx":333E
      MousePointer    =   99  'Custom
      Picture         =   "frmDisconnetti.frx":3648
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   0
      ToolTipText     =   "Termina l'esecuzione del programma"
      Top             =   960
      Width           =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   720
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   10
      X2              =   10
      Y1              =   480
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Annulla"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label lblChiudi 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "      Chiudi    Effettua Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1725
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambia  Utente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDisconnetti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    Dim Y As Integer
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 2
    Me.ScaleMode = 3
    Me.ScaleHeight = (256 * 2)
    For i = 0 To 255
        Me.Line (0, Y)-(Me.Width, Y + 2), RGB(i, i, i), BF
        Y = Y + 2
    Next i
    Call Text3D(Picture3, "IsoDial", "Times New Roman", 26, 1800, 80, 100, 60, 181, 247)   'modificare tutti e tre i numeri
    If Not structApri.server Then
        lblChiudi = "        Chiudi    "
        chkSpegni.Visible = False
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nRet As Long
    ReleaseCapture
    nRet = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub picAnnulla_Click()
    tDisconnetti = tpDANNULLA
    Unload Me
End Sub

Private Sub picAnnulla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picAnnulla.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub picAnnulla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picAnnulla.Picture = ImageList1.ListImages(1).Picture
End Sub

Private Sub picChiudi_Click()
    tDisconnetti = tpDCHIUDICONBACKUP
    '/ prevent error if the menu is not subclassed
    On Error Resume Next
    '/ release object
    Call objMenuEx.Uninstall(Me.hWnd, ImageList1, MenuEvents)
    Set MenuEvents = Nothing
    Set objMenuEx = Nothing
    If structApri.server Then
        ' effettua la copia di backup
        tPeriferica = tpBACKUP
        If (chkSpegni.Value = Checked) Then
            SpegniPc = True
        Else
            SpegniPc = False
        End If
        frmTrasferisciFile.Show 1
    End If
    Unload Me
End Sub

Private Sub picChiudi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picChiudi.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub picChiudi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picChiudi.Picture = ImageList1.ListImages(5).Picture
End Sub

Private Sub picDisconnetti_Click()
    tDisconnetti = tpDLOGIN
    ' cancella le info sul dottore
    tAccesso.cognome = ""
    tAccesso.nome = ""
    tAccesso.pass = ""
    tAccesso.Tipo = tpAMASTER
    frmMain.staBar.Panels(4).Text = ""
    ' mostra frmLogin
    Unload Me
    frmLogin.Show 1
End Sub

Private Sub picDisconnetti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDisconnetti.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub picDisconnetti_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDisconnetti.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nRet As Long
    ReleaseCapture
    nRet = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nRet As Long
    ReleaseCapture
    nRet = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

