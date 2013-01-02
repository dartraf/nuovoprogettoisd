VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMostraElaborazioni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualizza giorni dialisi"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10695
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
         ItemData        =   "frmMostraElaborazioni.frx":0000
         Left            =   4200
         List            =   "frmMostraElaborazioni.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   12
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
         ItemData        =   "frmMostraElaborazioni.frx":0004
         Left            =   960
         List            =   "frmMostraElaborazioni.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pazienti in elenco:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label lblVoci 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   9600
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anno:"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   250
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mese:"
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
         TabIndex        =   7
         Top             =   255
         Width           =   645
      End
   End
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
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   3
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmMostraElaborazioni.frx":009A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraAzioni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   10695
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
         TabIndex        =   2
         Top             =   240
         Width           =   1455
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
         Left            =   9000
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTotaleDialisi 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3030
         TabIndex        =   11
         Top             =   260
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Totale Dialisi:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmMostraElaborazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmMostraElaborazioni.frm
'
' <b>Descrizione</b>: Pannello che mostra i giorni in dialisi dei pazienti per un dato mese e anno
'
' @remarks
'
' @author
'
' @date 21/02/2011 19.54
Option Explicit

Dim stoCaricando As Boolean

Private Sub Form_Load()
    Dim i As Integer
    With flxGriglia
        .Row = 0
        For i = 0 To 2
            .Col = i
            .CellFontBold = True
        Next i
        .Col = 0
        .ColAlignment(2) = vbLeftJustify
        '.ColAlignment(1) = vbRightJustify
    End With
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    flxGriglia.Col = 1
    flxGriglia.Row = 0
    flxGriglia.CellAlignment = flexAlignRightTop
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'------------------------------


'' Calcola la stringa elenco dei giorni di dialisi di un determinato paziente
'
' @param evStr condizione per la query
' @param totale numero totale di giorni
' @return stringa elenco dei giorni dialisi

Private Function giorniDialisi(evStr As String, ByRef totale As Integer) As String
    Dim rsDialisi As Recordset
    Set rsDialisi = New Recordset
    Dim v_giorni() As Integer
    Dim i As Integer
    
    ' resetta le var
    giorniDialisi = ""
    ReDim v_giorni(0)
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI " & evStr & " AND ERRATA=FALSE AND Month([DATA])=" & cboMese.ListIndex + 1 & " AND Year([DATA])=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDialisi.EOF
        ReDim Preserve v_giorni(UBound(v_giorni) + 1)
        v_giorni(UBound(v_giorni)) = Day(rsDialisi("DATA"))
        rsDialisi.MoveNext
    Loop
    Call BubbleSort(v_giorni)
    For i = 1 To UBound(v_giorni)
        giorniDialisi = giorniDialisi & IIf(Len(CStr(v_giorni(i))) = 1, Space(1), "") & v_giorni(i) & " - "
    Next i
    ' elimina il - finale
    If giorniDialisi <> "" Then
        giorniDialisi = Left(giorniDialisi, Len(giorniDialisi) - 3)
    End If
    totale = UBound(v_giorni)
    Set rsDialisi = Nothing
End Function

'' Riempe la flx con la lista dei giorni dialisi
Private Sub CaricaFatture()
    Dim i As Integer
    Dim totaleDialisi As Integer
    Dim totale As Integer
    Dim giorni As String
    Dim rsPazienti As Recordset
    
    flxGriglia.Rows = 1
    Set rsPazienti = New Recordset
    rsPazienti.Open "SELECT * FROM PAZIENTI ORDER BY COGNOME, NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsPazienti.EOF And rsPazienti.BOF Then
        MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
        Exit Sub
    Else
        With flxGriglia
            .Col = 0
            Do While Not rsPazienti.EOF
                totale = 0
                giorni = giorniDialisi("WHERE CODICE_PAZIENTE=" & rsPazienti("KEY"), totale)
                If giorni <> "" Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsPazienti("COGNOME") & "  " & rsPazienti("NOME")
                    .TextMatrix(.Rows - 1, 2) = giorni
                    .TextMatrix(.Rows - 1, 1) = IIf(totale = 0, "", totale) & "  "
                    .Row = .Rows - 1
                    .CellBackColor = RGB(231, 255, 255)
                End If
                rsPazienti.MoveNext
            Loop
        End With
        Set rsPazienti = Nothing
        
        'Somma le dialisi nella FlexGrid
        For i = 1 To flxGriglia.Rows - 1
            totaleDialisi = totaleDialisi + flxGriglia.TextMatrix(i, 1)
        Next i
            
        lblTotaleDialisi = totaleDialisi
        lblVoci = flxGriglia.Rows - 1
        
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim i As Integer
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(50) AS PAZIENTE, " & _
                "       NEW adInteger AS TOTALE, " & _
                "       NEW adLongVarChar AS DIALISI "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        
    With rsMain
        For i = 1 To flxGriglia.Rows - 1
            .AddNew
            .Fields("PAZIENTE") = " " & flxGriglia.TextMatrix(i, 0)
            .Fields("TOTALE") = flxGriglia.TextMatrix(i, 1)
            .Fields("DIALISI") = " " & flxGriglia.TextMatrix(i, 2)
        Next i
    End With
    
    
    Set rptMostraFatture.DataSource = rsMain
    rptMostraFatture.LeftMargin = rptMostraFatture.LeftMargin / 3
    rptMostraFatture.RightMargin = rptMostraFatture.RightMargin / 3
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblMese").Caption = cboMese.Text
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblPazienti").Caption = flxGriglia.Rows - 1
    rptMostraFatture.Sections("Section5").Controls.Item("lblDialisi").Caption = lblTotaleDialisi.Caption
    rptMostraFatture.PrintReport True, rptRangeAllPages
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cboMese_Click()
    If stoCaricando Then Exit Sub
    Call CaricaFatture
End Sub

Private Sub cboAnno_Click()
    If stoCaricando Then Exit Sub
    Call CaricaFatture
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, 2, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, 2)
    End If
End Sub

