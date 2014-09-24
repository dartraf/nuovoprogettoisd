VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmVisualizzaFattureElettroniche 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualizza Fatture Elettroniche"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   99
         FormatString    =   "| Tabella                                                                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmVisualizzaFattureElettroniche.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   7215
      Begin VB.CommandButton RigeneraFE 
         Caption         =   "Rigenera XML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton VediFE 
         Caption         =   "Visualizza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
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
         Height          =   510
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmVisualizzaFattureElettroniche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVisualizza As Recordset
Dim vRow As Integer             ' riga selezionata
Dim vCol As Integer             ' colonna selezionata
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx
Dim strSql As String

Private Sub Form_Load()
    Dim i As Integer
    
    strSql = "SELECT * FROM FE ORDER BY DATA_INVIO DESC"

    Set objAnnulla = New CAnnulla
    flxGriglia.Rows = 1
    
    With flxGriglia
        .Cols = 6
        .ColWidth(1) = .ColWidth(1) * 0.25
        .ColWidth(2) = .ColWidth(2) * 1.6
        .ColWidth(3) = .ColWidth(3) * 1.3
        .ColWidth(4) = .ColWidth(4) * 1.1
        .ColWidth(5) = .ColWidth(5) * 1.9
                                     
        .TextMatrix(0, 1) = "N°Fattura"
        .TextMatrix(0, 2) = "Tipo Documento"
        .TextMatrix(0, 3) = "N°Progr.Invio"
        .TextMatrix(0, 4) = "Data Invio"
        .TextMatrix(0, 5) = "Nome File"
    End With
    
    Call CaricaFlx
    
    With flxGriglia
        .ColWidth(0) = 0
         .Row = 0
         .MousePointer = flexCustom
         For i = 0 To flxGriglia.Cols - 1
            .Col = i
            .ColAlignment(i) = vbLeftJustify
            .CellFontBold = True
         Next i
     End With
     
End Sub

Private Sub CaricaFlx()
    flxGriglia.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    
    Set rsVisualizza = New Recordset
    rsVisualizza.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not (rsVisualizza.EOF And rsVisualizza.BOF) Then
        Do While Not rsVisualizza.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsVisualizza("KEY")
                .TextMatrix(.Rows - 1, 1) = rsVisualizza("N_FATTURA") '& ""
                .TextMatrix(.Rows - 1, 2) = rsVisualizza("TIPO_DOC") '& ""
                .TextMatrix(.Rows - 1, 3) = rsVisualizza("PROGR_INVIO") '& ""
                .TextMatrix(.Rows - 1, 4) = rsVisualizza("DATA_INVIO") '& ""
                .TextMatrix(.Rows - 1, 5) = rsVisualizza("NOME_FILE") '& ""
                rsVisualizza.MoveNext
            End With
        Loop
    End If
    Set rsVisualizza = Nothing
    flxGriglia.Row = 0
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

Private Sub flxGriglia_dblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    
    ' Seleziono la key della FE e la passo con la variabile, altrimenti da errore
    
    tTrova.keyGestioneApparato = KeyFE
    MantieniKeyReturn = tTrova.KeyFE
    'Visualizza nel browser la fattura dal file XML della cartella FE
    'SHOW_SHOWNORMAL = 1
    'SHOW_SHOWMAXIMIZED = 3
    ret = ShellExecute(Me.hWnd, "open", structApri.pathExe & "\FE\" & NameExtXML, vbNullString, vbNullString, 1)
    If ret < 32 Then MsgBox "Si è verificato un errore aprendo il browser di default", vbCritical, "ATTENZIONE!!!"
    tTrova.KeyFE = 0    'per evitare di ricaricare la FE
End Sub

Private Sub cmdChiudi_Click()
    Unload frmVisualizzaFattureElettroniche
End Sub

Private Sub WheelCatcher1_WheelRotation(Rotation As Long, X As Long, Y As Long, CtrlHwnd As Long)
On Error GoTo gestione

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
' Evita crash in caso di griglia non completa
gestione:

End Sub
