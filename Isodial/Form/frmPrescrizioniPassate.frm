VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrescrizioniPassate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elenco ricette"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cboAnno 
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
         ItemData        =   "frmPrescrizioniPassate.frx":0000
         Left            =   4800
         List            =   "frmPrescrizioniPassate.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboMese 
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mese"
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
         TabIndex        =   3
         Top             =   260
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anno"
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
         Left            =   4080
         TabIndex        =   2
         Top             =   260
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "| Data ricetta          | Data prenotazione     | Numero ricetta        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6135
      Begin VB.CommandButton cmdCarica 
         Caption         =   "C&arica"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdChiudi 
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
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPrescrizioniPassate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPrescrizioniPassate.frm
'
' <b>Descrizione</b>: Pannello che mostra le vecchie ricette di un paziente
'
' @remarks
'
' @author
'
' @date 07/08/2011 11.54
Option Explicit

'' rs della scheda
Dim rsDataset As Recordset

Private keyReturn As Integer
Private keyPaziente As Integer
Dim stoCaricando As Boolean

Public Property Get getkeyReturn() As Integer
    getkeyReturn = keyReturn
End Property

Public Property Let LetkeyPaziente(ByVal vkeyPaziente As Integer)
    keyPaziente = vkeyPaziente
End Property

Private Sub cboMese_Click()
    Call CaricaFlx
End Sub

Private Sub Form_Activate()
    Call CaricaFlx
End Sub

Private Sub Form_Load()
    Dim i   As Integer
        
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    For i = 1 To 12
        cboMese.AddItem UCase(MonthName(i))
    Next i
    cboMese.ListIndex = Month(Now) - 1

    With flxGriglia
        .ColAlignment(0) = vbLeftJustify
        .Row = 0
        .ColWidth(0) = 0
        For i = 1 To 3
            .Col = i
            .CellFontBold = True
        Next i
    End With
    
    laData = ""
End Sub

'' Permette il funzionamento della rotellina del mouse nella flx
'Public Sub MouseWheel(flx As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'    Dim NewValue As Long
'    Dim Lstep As Single

'    On Error Resume Next
'    With flx
'        Lstep = .Height / .RowHeight(0)
'        Lstep = Int(Lstep)
'        If Lstep < 10 Then
'            Lstep = 10
'        End If
'        If Rotation > 0 Then
'            NewValue = .TopRow - Int(Lstep / 3)
'            If NewValue < 1 Then
'                NewValue = 1
'            End If
'        Else
'            NewValue = .TopRow + Int(Lstep / 3)
'            If NewValue > .Rows - 1 Then
'                NewValue = .Rows - 1
'            End If
'        End If
'        .TopRow = NewValue
'    End With
'End Sub
'---------------------------------------


Private Sub CaricaFlx()
    With flxGriglia
        .Rows = 1
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM RICETTE WHERE (NOT FLAG=3 AND CODICE_PAZIENTE=" & keyPaziente & " AND YEAR([DATA_RICETTA])=" & cboAnno.Text & " AND MONTH([DATA_RICETTA])=" & cboMese.ListIndex + 1 & ") ORDER BY DATA_RICETTA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDataset.EOF
            .Rows = flxGriglia.Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("DATA_RICETTA")
            .TextMatrix(.Rows - 1, 2) = rsDataset("DATA_PRENOTAZIONE")
            .TextMatrix(.Rows - 1, 3) = rsDataset("NUMERO_RICETTA")
            rsDataset.MoveNext
        Loop
        Set rsDataset = Nothing
    End With
    keyReturn = 0
End Sub

Private Sub cmdCarica_Click()
    If flxGriglia.Row <> 0 Then
        keyReturn = flxGriglia.TextMatrix(flxGriglia.Row, 0)
    Else
        keyReturn = 0
    End If
    Unload Me
End Sub

Private Sub cmdChiudi_Click()
    keyReturn = 0
    Unload Me
End Sub

'Private Sub flxGriglia_GotFocus()
'    Call WheelHook(Me, flxGriglia)
'End Sub

'Private Sub flxGriglia_LostFocus()
'    Call WheelUnHook
'End Sub
'-------------------------------

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdChiudi_Click
    End If
End Sub

Private Sub flxGriglia_Click()
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) = True Then
        cmdCarica_Click
    End If
End Sub

Private Sub cboAnno_Click()
    If stoCaricando Then Exit Sub
    Call CaricaFlx
End Sub

