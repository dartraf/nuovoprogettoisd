VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVisualizzaDialisiMensiliPazienti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualizza dialisi mensili pazienti"
   ClientHeight    =   5856
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10824
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5856
   ScaleWidth      =   10824
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10695
      Begin VB.ComboBox cboMeseFianle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":0000
         Left            =   2280
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboAnno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":0096
         Left            =   5160
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboMeseInziale 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":009A
         Left            =   120
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":00C2
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
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         TabIndex        =   10
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label lblVoci 
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
         Height          =   360
         Left            =   9600
         TabIndex        =   8
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4440
         TabIndex        =   7
         Top             =   252
         Width           =   600
      End
   End
   Begin VB.Frame fraListaMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      Begin MSFlexGridLib.MSFlexGrid flxGriglia1 
         Height          =   1212
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10452
         _ExtentX        =   18436
         _ExtentY        =   2138
         _Version        =   393216
         Cols            =   25
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmVisualizzaDialisiMensiliPazienti.frx":0130
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1932
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   10452
         _ExtentX        =   18436
         _ExtentY        =   3408
         _Version        =   393216
         Cols            =   3
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "Paziente                                                 | inizio| dine ! totale"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Size            =   9.6
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
            Size            =   9.6
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
            Size            =   9.6
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
   End
End
Attribute VB_Name = "frmVisualizzaDialisiMensiliPazienti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim stoCaricando As Boolean





Private Sub cboMeseFianle_Click()

    If stoCaricando Then Exit Sub
    Call CaricaFatture
End Sub


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

Private Function giorniDialisi(evStr As String, ByRef totale As Integer) As String
    Dim rsDialisi As Recordset
    Set rsDialisi = New Recordset
    Dim v_giorni() As Integer
    Dim i As Integer
    
    
    ' resetta le var
    giorniDialisi = ""
    ReDim v_giorni(0)
    rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI " & evStr & " AND ERRATA=FALSE AND Month([DATA])>=" & cboMeseInziale.ListIndex + 1 & " AND Month([DATA])<=" & cboMeseFianle.ListIndex + 1 & " AND  Year([DATA])=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

 '   rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI " & evStr & " AND ERRATA=FALSE AND Month([DATA])=" & cboMeseInziale.ListIndex + 1 & " AND " & cboMeseFianle.ListIndex + 1 & "  AND Year([DATA])=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
                                                 '   "          CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 1) & " AND " & _
                                                 '   "          DATA BETWEEN #" & data_min & "# AND #" & data_max & "# " & _
                                                 '     "ORDER BY  DATA DESC"
'    Month([data]) = " & flxGriglia.Col & " And " &" _

    Do While Not rsDialisi.EOF
        ReDim Preserve v_giorni(UBound(v_giorni) + 1)
        v_giorni(UBound(v_giorni)) = Day(rsDialisi("DATA"))
        rsDialisi.MoveNext
    Loop
    Call BubbleSort(v_giorni)
    For i = 1 To UBound(v_giorni)
        giorniDialisi = giorniDialisi & IIf(Len(CStr(v_giorni(i))) = 1, Space(1), "") & v_giorni(i) & " - "
    Next i
  '  ' elimina il - finale
  '  If giorniDialisi <> "" Then
  '      giorniDialisi = Left(giorniDialisi, Len(giorniDialisi) - 3)
  '  End If
    totale = UBound(v_giorni)
    Set rsDialisi = Nothing

    
End Function

'' Riempe la flx con la lista dei giorni dialisi
Private Sub CaricaFatture()
    Dim i As Integer
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
                   ' .TextMatrix(.Rows - 1, 2) = giorni
                    .TextMatrix(.Rows - 1, 1) = IIf(totale = 0, "", totale) & "  "
                    .Row = .Rows - 1
                    .CellBackColor = RGB(231, 255, 255)
                End If
                rsPazienti.MoveNext
            Loop
        End With
        Set rsPazienti = Nothing
        

       ' lblVoci = flxGriglia.Rows - 1
        
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
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblPazienti").Caption = flxGriglia.Rows - 1
    rptMostraFatture.PrintReport True, rptRangeAllPages
End Sub

Private Sub cmdChiudi_Click()
    Unload frmVisualizzaDialisiMensiliPazienti
End Sub

'Private Sub cboMese_Click()
'    If stoCaricando Then Exit Sub
'    Call CaricaFatture
'End Sub

'Private Sub cboAnno_Click()
'    If stoCaricando Then Exit Sub
'    Call CaricaFatture
'End Sub

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

