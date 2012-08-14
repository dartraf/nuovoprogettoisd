VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStorico 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Storico dei"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxGriglia 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      FormatString    =   "Data            | Peso secco   "
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
Attribute VB_Name = "frmStorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDataset As Recordset
Dim nomeTabella As String



' questo form è strettamente legato al form frmAnamnesiDialitica
' soprattutto nei metodi di ordinamento

Private Sub Form_Load()
    Dim i As Integer
    Dim PuntoX As Integer
    Dim PuntoY As Integer
    flxGriglia.Row = 0
    For i = 0 To 1
        flxGriglia.Col = i
        flxGriglia.CellFontBold = True
    Next i
    Call PosizioneCursore(PuntoX, PuntoY)
    Me.Top = PuntoY
    Me.Left = PuntoX
    If Me.Left + Me.Width > frmMain.Width Then
        Me.Left = frmMain.Width - Me.Width - 300
    End If
    If Me.Top + Me.Height > frmMain.Height Then
        Me.Top = frmMain.Height - Me.Height - 300
    End If
    flxGriglia.ColAlignment(1) = vbLeftJustify
    Select Case tStorico.Tipo
        Case tpsFILTRO, tpsLINEE
            Me.Caption = Me.Caption & " tipi di " & IIf(tStorico.Tipo = tpsFILTRO, "filtro", "linee")
            flxGriglia.TextMatrix(0, 1) = "Tipo di " & IIf(tStorico.Tipo = tpsFILTRO, "filtro", "linee")
            flxGriglia.Width = 4800
            Me.Width = 5100
            flxGriglia.ColWidth(1) = 3500
            nomeTabella = "STORICO_DIALISI_" & IIf(tStorico.Tipo = tpsFILTRO, "FILTRO", "LINEE")
        Case tpsPESO
            Me.Caption = "Storico del peso secco"
            nomeTabella = "STORICO_DIALISI_PESO"
    End Select
    Call CaricaStorico
End Sub

Private Sub CaricaStorico()
    Dim strFrom As String
    flxGriglia.Rows = 1
    If tStorico.Tipo = tpsFILTRO Then
        strFrom = "( " & nomeTabella & " T INNER JOIN FILTRI F ON F.KEY=T.TIPO_FILTRO )"
    ElseIf tStorico.Tipo = tpsLINEE Then
        strFrom = "( " & nomeTabella & " T INNER JOIN LINEE F ON F.KEY=T.TIPO_LINEE )"
    Else
        strFrom = nomeTabella & " T "
    End If
    Set rsDataset = New Recordset
    ' mi fa visualizzare in base al key
    rsDataset.Open "SELECT * FROM " & strFrom & " " & tStorico.condizione & " ORDER BY T.KEY DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("DATA")
            Select Case tStorico.Tipo
                Case tpsPESO
                    .TextMatrix(.Rows - 1, 1) = rsDataset("PESO")
                Case tpsLINEE, tpsFILTRO
                    .TextMatrix(.Rows - 1, 1) = rsDataset("NOME")
            End Select
        End With
        rsDataset.MoveNext
    Loop
    Set rsDataset = Nothing
End Sub

Private Sub flxGriglia_Click()
    Call ColoraFlx(flxGriglia, 1)
End Sub

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
