VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDialisiDaFare 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schede Dialitiche Giornaliere - Compilazione"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Label lblAttenzione 
         AutoSize        =   -1  'True
         Caption         =   "ATTENZIONE!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   2790
      End
      Begin VB.Label lblTesto 
         AutoSize        =   -1  'True
         Caption         =   "Non risulta registrata la seduta dialitica dei pazienti in elenco"
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
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   6345
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid flxTrova 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "| Turno | Cognome                            |  Nome                                 |  Data di nascita          "
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
      TabIndex        =   4
      Top             =   4200
      Width           =   7935
      Begin VB.CommandButton cmdIndietro 
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
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAvanti 
         Caption         =   "C&ompila"
         Enabled         =   0   'False
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
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblVoci 
         AutoSize        =   -1  'True
         Caption         =   "Pazienti in elenco:  "
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2025
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "frmDialisiDaFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmDialisiDaFare.frm
'
' <b>Descrizione</b>: Pannello Dialisi da Fare mostra i pazienti che ancora non hanno una scheda dialitica giornaliera del giorno scelto
'
' @remarks
'
' @author
'
' @date 04/02/2011 19.47
Option Explicit

'' turno scelto
Private turno As Integer

Public Property Get getTurno() As Integer
    getTurno = turno
End Property

Public Property Let LetTurno(ByVal vturno As Integer)
    turno = vturno
End Property

Private Sub Form_Activate()
    Call Cerca
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Integer
    With flxTrova
        .ColWidth(0) = 0
        .ColAlignment(4) = vbLeftJustify
        .Row = 0
        For i = 1 To 4
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
    tTrova.keyReturn = -1
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
'------------------------------------------

'' Caricare solo i pazienti che hanno il turno dialitico nella data scelta
Private Sub Cerca()

    Dim rsAppo As New Recordset
    Dim rsPazientiTurni As Recordset
    Dim rsDialisi As Recordset
    
    Dim giorno As Integer       ' 1 lun 2 mart 3 merc ..
    Dim trovato As Boolean
    Dim num As Integer
    
    Dim strTurno As String
    Dim tipostrTurno As String
    Dim strSql As String
    
    Select Case turno
        Case 1
            tipostrTurno = "AM_INIZIO"
        Case 2
            tipostrTurno = "PM_INIZIO"
        Case Else
            tipostrTurno = "SR_INIZIO"
    End Select
    
    ' pulisce la flx azzerando le righe
    flxTrova.Rows = 1
    num = 0
    giorno = Weekday(laData, vbMonday)
    
    strSql = "SELECT    PAZIENTI.KEY, PAZIENTI.COGNOME, PAZIENTI.NOME, PAZIENTI.DATA_NASCITA, PAZIENTI.STATO, TURNI.AM_INIZIO" & giorno & ", TURNI.PM_INIZIO" & giorno & ", TURNI.SR_INIZIO" & giorno & " " & _
             "FROM      ((PAZIENTI " & _
             "          INNER JOIN TURNI ON PAZIENTI.KEY = TURNI.CODICE_PAZIENTE) " & _
             "          INNER JOIN RENI ON RENI.KEY=TURNI.CODICE_RENE) " & _
             "WHERE     ( (PAZIENTI.STATO=0 OR PAZIENTI.STATO=4) AND " & _
             "          TURNI." & tipostrTurno & giorno & "<>"""" )"
                         
    Set rsPazientiTurni = New Recordset
    Set rsDialisi = New Recordset
    rsPazientiTurni.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsAppo.Open "ANAMNESI_NEFROLOGICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    rsDialisi.Open "SCHEDE_DIALISI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    rsPazientiTurni.Sort = ("COGNOME")
    Do While Not rsPazientiTurni.EOF
        ' effettua il controllo sulla data fine nnn in query perche il campo nn è obbligatorio
        rsAppo.Filter = ("CODICE_PAZIENTE=" & rsPazientiTurni("KEY"))
        ' se nn esiste nn  puo effettuare la dialisi (cazzi suoi)
        trovato = False
        If Not (rsAppo.BOF And rsAppo.EOF) Then
            If rsAppo("DATA_INIZIO") <> "" Then
                If CDate(rsAppo("DATA_INIZIO")) <= laData Then
                    If rsAppo("DATA_FINE") <> "" Then
                        If CDate(rsAppo("DATA_FINE")) >= laData Then
                            trovato = True
                        End If
                    Else
                        trovato = True
                    End If
                End If
            End If
        End If
        If trovato Then
            rsDialisi.Filter = ("CODICE_PAZIENTE=" & rsPazientiTurni("KEY") & " AND DATA=#" & laData & "#")
            If rsDialisi.EOF And rsDialisi.BOF Then
                With flxTrova
                    ' nn ha trovato dialisi quindi lo aggiunge
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsPazientiTurni("KEY")
                    If turno = 1 Then
                        strTurno = "MAT"
                    ElseIf turno = 2 Then
                        strTurno = "POM"
                    Else
                        strTurno = "SER"
                    End If
                    .TextMatrix(.Rows - 1, 1) = strTurno
                    .TextMatrix(.Rows - 1, 2) = rsPazientiTurni("COGNOME")
                    .TextMatrix(.Rows - 1, 3) = rsPazientiTurni("NOME") & ""
                    .TextMatrix(.Rows - 1, 4) = rsPazientiTurni("DATA_NASCITA") & ""
                    num = num + 1
                End With
            End If
        End If
        rsPazientiTurni.MoveNext
    Loop
    If num = 0 Then
        Unload Me
    Else
        Beep
    End If
    lblVoci = "Pazienti in elenco: " & num
    Set rsDialisi = Nothing
    Set rsAppo = Nothing
    Set rsPazientiTurni = Nothing
End Sub

Private Sub cmdAvanti_Click()
    tTrova.keyReturn = flxTrova.TextMatrix(flxTrova.Row, 0)
    Unload Me
End Sub

Private Sub cmdIndietro_Click()
    tTrova.keyReturn = -1
    Unload Me
End Sub

Private Sub Timer1_Timer()
    If lblAttenzione.ForeColor = vbRed Then
        lblAttenzione.ForeColor = vbBlack
    Else
        lblAttenzione.ForeColor = vbRed
    End If
End Sub

'Private Sub flxTrova_GotFocus()
'    Call WheelHook(Me, flxTrova)
'End Sub

'Private Sub flxTrova_LostFocus()
'    Call WheelUnHook
'End Sub
'--------------------------------

Private Sub flxTrova_Click()
    On Error GoTo gestione
    If VerificaClickFlx(flxTrova) = False Then
        ' discolora
        Call ColoraFlx(flxTrova, flxTrova.Cols - 1, True)
        ' annulla le row e col
        flxTrova.Row = 0
        flxTrova.Col = 0
        cmdAvanti.Enabled = False
        Exit Sub
    Else
        Call ColoraFlx(flxTrova, flxTrova.Cols - 1)
        cmdAvanti.Enabled = True
    End If
    Exit Sub
gestione:
    MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
End Sub

Private Sub flxTrova_DblClick()
    If VerificaClickFlx(flxTrova) = False Then Exit Sub
    cmdAvanti_Click
End Sub

Private Sub flxTrova_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdIndietro_Click
    End If
End Sub

