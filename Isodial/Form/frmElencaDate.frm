VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmElencaDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selezionare Data"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   Begin WheelCatch.WheelCatcher WheelCatcher1 
      Height          =   480
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSFlexGridLib.MSFlexGrid flxGriglia 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   16776960
      ForeColorSel    =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   " | Elenco date   "
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
Attribute VB_Name = "frmElencaDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmElencaDate.frm
'
' <b>Descrizione</b>: Elenca le date caricate da una tabella
'
' @remarks
'
' @author
'
' @date 05/02/2011 16.38
Option Explicit

'' rs della scheda
Dim rsElenca As Recordset
Dim nomeTabella As String

'' Setta le impostazioni iniziali e carica i dati nel form
Private Sub Form_Load()
    'Dim nomeTabella As String
    Dim PuntoX As Integer
    Dim PuntoY As Integer
    
    Call PosizioneCursore(PuntoX, PuntoY)
    Me.Top = PuntoY
    Me.Left = PuntoX
    If Me.Left + Me.Width > frmMain.Width Then
        Me.Left = frmMain.Width - Me.Width - 300
    End If
    If Me.Top + Me.Height > frmMain.Height Then
        Me.Top = frmMain.Height - Me.Height - 300
    End If
        
    Select Case tElenca.Tipo
        Case tpREGISTRAZIONESAMI
            nomeTabella = "ANAMNESI_ESAMI"
        Case tpACCESSO
            nomeTabella = "ACCESSI_VASCOLARI_TAB"
        Case tpDIARIO
            nomeTabella = "DIARI_CLINICI"
        Case tpSCHEDEDIALITICHE
            nomeTabella = "SCHEDE_DIALISI"
        Case tpESAMISTRUMENTALI
            nomeTabella = "ESAMI_STRUMENTALI"
        Case tpCOLTURE
            nomeTabella = "COLTURE"
        Case tpMON_ACC_VASCOLARE
            nomeTabella = "MON_ACCESSI"
        Case tpMON_TRAT_ACQUE
            nomeTabella = "MON_TRAT_ACQUE"
        Case tpMON_VACC_EPATITE
            nomeTabella = "MON_VACC_EPATITE"
        Case tpMON_VAL_PSICO
            nomeTabella = "MON_VALUTAZIONI"
        Case tpESPORTAESAMI
            nomeTabella = "ANAMNESI_ESAMI"
            Me.Top = (Screen.Height - Me.Height) / 2 - 500
            Me.Left = (Screen.Width - Me.Width) / 2
    End Select
    flxGriglia.Rows = 1
    
            If nomeTabella = "MON_TRAT_ACQUE" Then
        Set rsElenca = New Recordset
        rsElenca.Open "SELECT DISTINCT DATA, KEY FROM " & nomeTabella & " " & tElenca.condizione & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsElenca.EOF
        flxGriglia.Rows = flxGriglia.Rows + 1
        flxGriglia.TextMatrix(flxGriglia.Rows - 1, 0) = rsElenca("KEY")
        flxGriglia.TextMatrix(flxGriglia.Rows - 1, 1) = rsElenca("DATA")
        rsElenca.MoveNext
        Loop
    Set rsElenca = Nothing
    
    Else
    
    Set rsElenca = New Recordset
    rsElenca.Open "SELECT DISTINCT DATA FROM " & nomeTabella & " " & tElenca.condizione & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsElenca.EOF
        flxGriglia.Rows = flxGriglia.Rows + 1
        flxGriglia.TextMatrix(flxGriglia.Rows - 1, 1) = rsElenca("DATA")
        rsElenca.MoveNext
    Loop
    Set rsElenca = Nothing
    
    End If
    
    With flxGriglia
        .ColAlignment(1) = vbLeftJustify
        .ColWidth(0) = 0
        .Row = 0
        .Col = 1
        .CellFontBold = True
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
'----------------------------------------

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub flxGriglia_Click()
    If VerificaClickFlx(flxGriglia) = False Then
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) Then
        If nomeTabella = "MON_TRAT_ACQUE" Then
            tTrova.keyReturn = flxGriglia.TextMatrix(flxGriglia.Row, 0)
            laData = flxGriglia.TextMatrix(flxGriglia.Row, 1)
        Else
            laData = flxGriglia.TextMatrix(flxGriglia.Row, 1)
        End If
        If Not IsDate(laData) Then
            laData = ""
        End If
        Unload Me
    End If
End Sub

'Private Sub flxGriglia_GotFocus()
    'Call WheelHook(Me, flxGriglia)
'End Sub

'Private Sub flxGriglia_LostFocus()
    'Call WheelUnHook
'End Sub
'-------------------------------

Private Sub WheelCatcher1_WheelRotation(Rotation As Long, X As Long, Y As Long, CtrlHwnd As Long)
' se NON è stata selezionata una riga esce e NON attiva lo scroll
'    If flxGriglia.Row = 0 Then
'       Exit Sub
'    End If

    Select Case CtrlHwnd

        Case flxGriglia.hWnd
            If flxGriglia.TopRow - Rotation > 0 Then
               flxGriglia.TopRow = flxGriglia.TopRow - Rotation
            End If
    
        End Select
End Sub

