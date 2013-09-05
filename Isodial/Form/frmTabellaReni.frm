VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTabellaReni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parco Reni"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraListaMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8025
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   15
         FormatString    =   "| N° rene   | Monitor                                             | Matricola         | Tipo            |Dt.Rottam.   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTabellaReni.frx":0000
      End
   End
   Begin VB.Frame fraAzioni 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   8025
      Begin VB.CommandButton cmdSostituisci 
         Caption         =   "&Sostituisci"
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
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   1230
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
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTabellaReni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDataset As Recordset
Dim rsTabella As Recordset
Dim vRow As Integer
Dim vCol As Integer

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    
    'Me.Top = intTop
    Me.Top = 3900
    Me.Left = intLeft

      
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 0 To 5
            .Col = i
            .CellFontBold = True
        Next i
       .MousePointer = vbArrow
    End With
    Call CaricaFlx
End Sub

Private Sub cmdSostituisci_Click()
            
    If flxGriglia.Row <> 0 Then
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_RENE=" & cod_rene, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
 
 ' sostituisce il numero del rene
        Do While Not rsDataset.EOF
           rsDataset("CODICE_RENE") = flxGriglia.TextMatrix(vRow, 0)
           rsDataset.MoveNext
        Loop
           rsDataset.Close
 ' flagga i reni come sostituiti
           rsDataset.Open "SELECT * FROM APPARATI WHERE KEY=" & cod_rene, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        Do While Not rsDataset.EOF
                rsDataset("SOSTITUITO") = True
                rsDataset.MoveNext
            Loop
           rsDataset.Close
             
           flxGriglia.Row = 0
           sostituito = True
           Unload frmTabellaReni

       Else
        MsgBox "Selezionare il rene in sostituzione", vbCritical, "Attenzione"
       End If
   End Sub


Private Sub CaricaFlx()
    Dim strSql As String
    Dim data As Date
    
    flxGriglia.Rows = 1
    data = DateValue(Month(dt_rott_rene) & "/" & Day(dt_rott_rene) & "/" & Year(dt_rott_rene))
    Set rsTabella = New Recordset
    strSql = "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE>#" & data & "# AND SOSTITUITO=FALSE OR DATA_ROTTAMAZIONE IS NULL ORDER BY NUMERO_APPARATO"

    rsTabella.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsTabella.EOF
       With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsTabella("KEY")
            .TextMatrix(.Rows - 1, 1) = rsTabella("NUMERO_APPARATO") & ""
            .TextMatrix(.Rows - 1, 2) = rsTabella("MODELLO") & ""
            .TextMatrix(.Rows - 1, 3) = rsTabella("MATRICOLA") & ""
            If rsTabella("TIPO") = 0 Then
               .TextMatrix(.Rows - 1, 4) = "NEG"
            ElseIf rsTabella("TIPO") = 1 Then
               .TextMatrix(.Rows - 1, 4) = "HCV POS"
            Else
               .TextMatrix(.Rows - 1, 4) = "HBV POS"
            End If
            .TextMatrix(.Rows - 1, 5) = rsTabella("DATA_ROTTAMAZIONE") & ""
        End With
        rsTabella.MoveNext
     Loop
    rsTabella.Close
    flxGriglia.Row = 0
    Set rsTabella = Nothing
End Sub

Private Sub cmdAnnulla_Click()
 flxGriglia.Row = 0
 Unload frmTabellaReni
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        vRow = flxGriglia.Row
        vCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub



