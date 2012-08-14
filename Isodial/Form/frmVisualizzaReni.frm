VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVisualizzaReni 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Reni"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraListaMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "| Postazione    | N° Rene   | Monitor                                         | Matricola           |Tipo             "
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
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8175
      Begin VB.CommandButton cmdConferma 
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
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Annulla"
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
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmVisualizzaReni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReni As Recordset

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'---------------------------------

Private Sub Form_Activate()
    Me.ZOrder
End Sub

Private Sub Form_Load()

    Dim i As Integer
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
    Call CaricaFlx
    With flxGriglia
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColWidth(0) = 0
        .Row = 0
        For i = 0 To 5
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
End Sub

Private Sub CaricaFlx()
    flxGriglia.Rows = 1
    Set rsReni = New Recordset
    rsReni.Open "RENI ORDER BY POSTAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsReni.BOF And rsReni.EOF) Then
        Do While Not rsReni.EOF
            If IsNull(rsReni("DATA_ROTTAMAZIONE")) Or rsReni("DATA_ROTTAMAZIONE") > date Or CBool(rsReni("SOSTITUITO")) = False Then
                With flxGriglia
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsReni("KEY")
                    .TextMatrix(.Rows - 1, 1) = rsReni("POSTAZIONE")
                    .TextMatrix(.Rows - 1, 2) = rsReni("NUMERO_RENE") & ""
                    .TextMatrix(.Rows - 1, 3) = rsReni("TIPO_RENE")
                    .TextMatrix(.Rows - 1, 4) = rsReni("MATRICOLA")
                    If rsReni("TIPO") = 0 Then
                        .TextMatrix(.Rows - 1, 5) = "NEG"
                    ElseIf rsReni("TIPO") = 1 Then
                        .TextMatrix(.Rows - 1, 5) = "HCV POS"
                    Else
                        .TextMatrix(.Rows - 1, 5) = "HBV POS"
                    End If
                End With
            End If
            rsReni.MoveNext
        Loop
        Set rsReni = Nothing
        flxGriglia.Row = 0
    End If
End Sub

Private Sub cmdAnnulla_Click()
    tReni.postazione = -1
    Unload Me
End Sub

Private Sub cmdConferma_Click()
    With flxGriglia
        If .Row <> 0 Then
            tReni.key = .TextMatrix(.Row, 0)
            tReni.postazione = .TextMatrix(.Row, 1)
            tReni.numero_rene = .TextMatrix(.Row, 2)
            tReni.monitor = .TextMatrix(.Row, 3)
            tReni.Tipo = .TextMatrix(.Row, 5)
            Unload Me
        End If
    End With
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
    cmdConferma_Click
End Sub

