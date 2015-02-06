VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVisualizzaDialisiMensiliPazienti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualizza dialisi mensili pazienti"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   732
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9612
      Begin VB.ComboBox cboMeseFianle 
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
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":0000
         Left            =   3840
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   260
         Width           =   1572
      End
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
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":0096
         Left            =   6240
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   260
         Width           =   855
      End
      Begin VB.ComboBox cboMeseInziale 
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
         ItemData        =   "frmVisualizzaDialisiMensiliPazienti.frx":009A
         Left            =   1200
         List            =   "frmVisualizzaDialisiMensiliPazienti.frx":00C2
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   260
         Width           =   1572
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Al mese"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dal mese"
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
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1005
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
         Left            =   5520
         TabIndex        =   6
         Top             =   255
         Width           =   600
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
      Height          =   3852
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9612
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3372
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9372
         _ExtentX        =   16536
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   14
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmVisualizzaDialisiMensiliPazienti.frx":0130
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
      TabIndex        =   5
      Top             =   4320
      Width           =   9612
      Begin VB.CommandButton cmdElabora 
         Caption         =   "&Elabora"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
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
         Left            =   6600
         TabIndex        =   1
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
         Height          =   492
         Left            =   8160
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblVoci 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4220
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Totale Pazienti:"
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
         Left            =   2540
         TabIndex        =   15
         Top             =   350
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Totale Dialisi:"
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
         Left            =   200
         TabIndex        =   13
         Top             =   350
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotaleDialisi 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   1730
         TabIndex        =   12
         Top             =   360
         Width           =   500
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
Dim GiorniMese() As Integer

Private Sub cboMeseFianle_Click()
    If cboMeseInziale.ListIndex > cboMeseFianle.ListIndex Then
        MsgBox "Il mese selezionato deve essere SUPERIORE" & vbCrLf & "rispetto a quello precedente", vbInformation, "Informazione"
        cboMeseFianle.SetFocus
        Exit Sub
    Else
        If stoCaricando Then Exit Sub
        Call CaricaFatture
    End If
End Sub

Private Sub cboMeseInziale_Click()
    If cboMeseInziale.ListIndex > cboMeseFianle.ListIndex Then
        MsgBox "Il mese selezionato deve essere INFERIORE" & vbCrLf & "rispetto a quello precedente", vbInformation, "Informazione"
        cboMeseInziale.SetFocus
        Exit Sub
    Else
        If stoCaricando Then Exit Sub
        Call CaricaFatture
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    With flxGriglia
        .Row = 0
        For i = 0 To 13
            .Col = i
            .CellFontBold = True
        Next i
    End With
    stoCaricando = True
    cboMeseFianle.ListIndex = 11
    cboMeseInziale.ListIndex = 0
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    flxGriglia.Col = 1
    flxGriglia.Row = 0
    flxGriglia.CellAlignment = flexAlignLeftTop
'    Call CaricaFatture
End Sub

Private Function giorniDialisi(evStr As String) As String
    Dim rsDialisi As Recordset
    Set rsDialisi = New Recordset
    Dim v_giorni() As Integer
    Dim i As Integer
    Dim h As Integer
    ReDim GiorniMese(12)

    ' resetta la var
    giorniDialisi = ""
    
    For h = cboMeseInziale.ListIndex + 1 To cboMeseFianle.ListIndex + 1
        ReDim v_giorni(0)
        rsDialisi.Open "SELECT * FROM SCHEDE_DIALISI " & evStr & " AND ERRATA=FALSE AND Month([DATA])=" & h & " AND  Year([DATA])=" & cboAnno.Text, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

        Do While Not rsDialisi.EOF
            ReDim Preserve v_giorni(UBound(v_giorni) + 1)
            v_giorni(UBound(v_giorni)) = Day(rsDialisi("DATA"))
            rsDialisi.MoveNext
        Loop
    
        Call BubbleSort(v_giorni)
    
        For i = 1 To UBound(v_giorni)
            giorniDialisi = giorniDialisi & IIf(Len(CStr(v_giorni(i))) = 1, Space(1), "") & v_giorni(i) & " - "
        Next i

        GiorniMese(h) = UBound(v_giorni)
        rsDialisi.Close
    Next h
    Set rsDialisi = Nothing
    End Function

Private Sub CaricaFatture()
Dim i As Integer
Dim totale As Integer
Dim giorni As String
Dim rsPazienti As Recordset
Dim totaleDialisi As Integer
Dim h As Integer
    
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
                giorni = giorniDialisi("WHERE CODICE_PAZIENTE=" & rsPazienti("KEY"))
                If giorni <> "" Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsPazienti("COGNOME") & "  " & rsPazienti("NOME")
                    For h = cboMeseInziale.ListIndex + 1 To cboMeseFianle.ListIndex + 1
                        .TextMatrix(.Rows - 1, h) = GiorniMese(h)
                        'calcola il totale delle dialisi del paziente
                        totale = totale + GiorniMese(h)
                    Next h
                    .TextMatrix(.Rows - 1, 13) = IIf(totale = 0, "", totale) & "  "
                    .Row = .Rows - 1
                    .CellBackColor = RGB(231, 255, 255)
                End If
                rsPazienti.MoveNext
            Loop
        End With
        Set rsPazienti = Nothing
        
        
    'Somma le dialisi nella FlexGrid
    For i = 1 To flxGriglia.Rows - 1
        totaleDialisi = totaleDialisi + flxGriglia.TextMatrix(i, 13)
    Next i
            
    lblTotaleDialisi.Caption = totaleDialisi
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
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    rptMostraFatture.Sections("Intestazione").Controls.Item("lblPazienti").Caption = flxGriglia.Rows - 1
    rptMostraFatture.PrintReport True, rptRangeAllPages
End Sub

Private Sub cmdChiudi_Click()
    Unload frmVisualizzaDialisiMensiliPazienti
End Sub


