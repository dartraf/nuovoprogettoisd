VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrapianti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pazienti candidati ai trapianti"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid flxTrova 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmTrapianti.frx":0000
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
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   10695
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Index           =   4
         Left            =   8880
         TabIndex        =   31
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Index           =   3
         Left            =   8880
         TabIndex        =   30
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Index           =   2
         Left            =   8880
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Index           =   1
         Left            =   8880
         TabIndex        =   28
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdGestioneReferti 
         Caption         =   "&Gestione Referti"
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
         Index           =   0
         Left            =   8880
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboCentroTrapianti 
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
         Index           =   1
         ItemData        =   "frmTrapianti.frx":00AB
         Left            =   1800
         List            =   "frmTrapianti.frx":00AD
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cboCentroTrapianti 
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
         Index           =   0
         ItemData        =   "frmTrapianti.frx":00AF
         Left            =   1800
         List            =   "frmTrapianti.frx":00B1
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cboCentroTrapianti 
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
         Index           =   2
         ItemData        =   "frmTrapianti.frx":00B3
         Left            =   1800
         List            =   "frmTrapianti.frx":00B5
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.ComboBox cboCentroTrapianti 
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
         Index           =   3
         ItemData        =   "frmTrapianti.frx":00B7
         Left            =   1800
         List            =   "frmTrapianti.frx":00B9
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox cboCentroTrapianti 
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
         Index           =   4
         ItemData        =   "frmTrapianti.frx":00BB
         Left            =   1800
         List            =   "frmTrapianti.frx":00BD
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtNoteTrapianti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   5400
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtNoteTrapianti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   5400
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtNoteTrapianti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   5400
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtNoteTrapianti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   5400
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtNoteTrapianti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   5400
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro trapianti "
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
         Index           =   20
         Left            =   120
         TabIndex        =   24
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro trapianti "
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
         Index           =   21
         Left            =   120
         TabIndex        =   23
         Top             =   885
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro trapianti "
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
         Index           =   22
         Left            =   120
         TabIndex        =   22
         Top             =   1485
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro trapianti "
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
         Index           =   23
         Left            =   120
         TabIndex        =   21
         Top             =   2085
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro trapianti "
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
         Index           =   24
         Left            =   120
         TabIndex        =   20
         Top             =   2685
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   25
         Left            =   4800
         TabIndex        =   19
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   26
         Left            =   4800
         TabIndex        =   18
         Top             =   885
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   27
         Left            =   4800
         TabIndex        =   17
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   28
         Left            =   4800
         TabIndex        =   16
         Top             =   2085
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note"
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
         Index           =   29
         Left            =   4800
         TabIndex        =   15
         Top             =   2685
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   5280
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
         Left            =   6240
         TabIndex        =   26
         Top             =   240
         Width           =   1215
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
         Left            =   9240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
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
         Left            =   7680
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmTrapianti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTrapianti As Recordset
Dim modifica As Boolean
Dim keyId As Integer

Private Sub Form_Activate()
    Dim i As Integer
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    For i = 0 To 4
        Call RicaricaComboBox("CENTRI_TRAPIANTO", "NOME", cboCentroTrapianti(i))
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxTrova
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 4
            .ColAlignment(i) = vbLeftJustify
            .Col = i
            .CellFontBold = True
        Next i
    End With
    modifica = False
    keyId = 0
    Call CaricaFlx
    Call EliminaScansioniSospese("SCAN_TRAPIANTI")
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxTrova.hWnd Then
'        If flxTrova.TopRow - Rotation > 0 Then
'            If flxTrova.TopRow - Rotation < flxTrova.Rows Then
'                flxTrova.TopRow = flxTrova.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'-----------------------------------

Private Sub Form_Unload(Cancel As Integer)
    Call EliminaScansioniSospese("SCAN_TRAPIANTI")
End Sub

Private Sub Pulisci()
    Dim i As Integer
    modifica = False
    keyId = 0
    For i = 0 To 4
        cboCentroTrapianti(i).ListIndex = -1
        cboCentroTrapianti(i).Text = ""
        txtNoteTrapianti(i) = ""
    Next i
    Call EliminaScansioniSospese("SCAN_TRAPIANTI")
End Sub

Private Sub CaricaScheda()
    Dim i As Integer
    
    If flxTrova.Row = 0 Then Exit Sub
    Call Pulisci
    
    Set rsTrapianti = New Recordset
    rsTrapianti.Open "SELECT * FROM MON_PAZ_TRAPIANTO WHERE CODICE_PAZIENTE=" & flxTrova.TextMatrix(flxTrova.Row, 0), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTrapianti.EOF And rsTrapianti.BOF) Then
        For i = 1 To 5
            cboCentroTrapianti(i - 1).ListIndex = GetCboListIndex(rsTrapianti("CODICE_CENTRO" & i), cboCentroTrapianti(i - 1))
            txtNoteTrapianti(i - 1) = rsTrapianti("NOTE" & i) & ""
        Next i
        modifica = True
        keyId = rsTrapianti("KEY")
    Else
        modifica = False
        keyId = 0
    End If
    Set rsTrapianti = Nothing
End Sub

Private Sub PulisciTutto()
    ' discolora
    Call ColoraFlx(flxTrova, flxTrova.Cols - 1, True)
    ' annulla le row e col
    flxTrova.Row = 0
    flxTrova.Col = 0
    Call Pulisci
End Sub

Private Sub CaricaFlx()
    flxTrova.Rows = 1
    Set rsTrapianti = New Recordset
    rsTrapianti.Open "SELECT STATO, PAZIENTI.COGNOME, PAZIENTI.NOME, PAZIENTI.KEY, PAZIENTI.DATA_NASCITA, PAZIENTI.CODICE_FISCALE, ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE, ANAMNESI_NEFROLOGICHE.ATTESA_TRAPIANTO, ANAMNESI_NEFROLOGICHE.SEDE1" & _
                     " FROM ANAMNESI_NEFROLOGICHE, PAZIENTI" & _
                     " WHERE (((PAZIENTI.KEY)=[CODICE_PAZIENTE]) AND ((ANAMNESI_NEFROLOGICHE.ATTESA_TRAPIANTO)=True)) AND STATO=0", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    With flxTrova
        ' pulisce la flx azzerando le righe
        Do While Not rsTrapianti.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsTrapianti("KEY")
            .TextMatrix(.Rows - 1, 1) = rsTrapianti("COGNOME") & ""
            .TextMatrix(.Rows - 1, 2) = rsTrapianti("NOME") & ""
            .TextMatrix(.Rows - 1, 3) = rsTrapianti("DATA_NASCITA")
            .TextMatrix(.Rows - 1, 4) = rsTrapianti("CODICE_FISCALE")
            rsTrapianti.MoveNext
        Loop
    End With
    Set rsTrapianti = Nothing
End Sub

Private Sub cmdGestioneReferti_Click(Index As Integer)
    Dim keyCentro As Integer
    
    If flxTrova.Row = 0 Then
        MsgBox "E' necessario selezionare prima il paziente", vbInformation, "Informazione"
        Exit Sub
    End If
    If cboCentroTrapianti(Index).Text = "" Then
        MsgBox "E' necessario selezionare prima il centro trapianti", vbInformation, "Informazione"
        Exit Sub
    End If

    Unload frmGestioneDocumentiEsterni
    Load frmGestioneDocumentiEsterni
    frmGestioneDocumentiEsterni.LetCodicePaziente = flxTrova.TextMatrix(flxTrova.Row, 0)
    Call GestisciNuovo("CENTRI_TRAPIANTO", cboCentroTrapianti(Index))
    keyCentro = cboCentroTrapianti(Index).ItemData(cboCentroTrapianti(Index).ListIndex)
    frmGestioneDocumentiEsterni.LetcodiceCentro = keyCentro
    If modifica And Salvato(Index) Then
        frmGestioneDocumentiEsterni.letcodiceRecord = keyId
        frmGestioneDocumentiEsterni.LetNomeFile = M_TR & keyId & " " & keyCentro & " " & Replace(date, "/", "-")
    Else
        frmGestioneDocumentiEsterni.letcodiceRecord = 0
        frmGestioneDocumentiEsterni.LetNomeFile = M_TR & 0 & " " & keyCentro & " " & Replace(date, "/", "-")
    End If
    tDocumentiEsterni = tpSCANTRAPIANTI
    frmGestioneDocumentiEsterni.Show 1
End Sub

Private Function Salvato(Index As Integer) As Boolean
    ' verifica se il cbo(index) e stato salvato oppure appena inserito
    Set rsTrapianti = New Recordset
    rsTrapianti.Open "SELECT * FROM MON_PAZ_TRAPIANTO WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsTrapianti.EOF And rsTrapianti.BOF) Then
        If rsTrapianti("CODICE_CENTRO" & Index + 1) <> -1 Then
            Salvato = True
        Else
            Salvato = False
        End If
    Else
        Salvato = False
    End If
    Set rsTrapianti = Nothing
End Function

Private Sub cmdStampa_Click()
    Dim strSqlStampa As String
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    
    strSqlStampa = "SHAPE APPEND  NEW adVarChar(30) AS COGNOME, " & _
                "           NEW adVarChar(30) AS NOME, " & _
                "           NEW adVarChar(15) AS DATA_NASCITA, " & _
                "           NEW adVarChar(20) AS CODICE_FISCALE, " & _
                "           NEW adLongVarChar AS CENTRO_TRAPIANTI, " & _
                "           NEW adLongVarChar AS NOTE "
                 
         
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
        
    If flxTrova.Rows <= 1 Then
       MsgBox "Inserire i valore", vbInformation, "Attenzione"
       Exit Sub
    End If
        
    Set rsTrapianti = New Recordset
    strSql = "SELECT        PAZIENTI.NOME, PAZIENTI.COGNOME, DATA_NASCITA, CODICE_FISCALE, CODICE_CENTRO1, " & _
                "           CENTRI_TRAPIANTO1.NOME AS CENTRI_TRAPIANTO1NOME, CENTRI_TRAPIANTO2.NOME AS CENTRI_TRAPIANTO2NOME, " & _
                "           CENTRI_TRAPIANTO3.NOME AS CENTRI_TRAPIANTO3NOME, CENTRI_TRAPIANTO4.NOME AS CENTRI_TRAPIANTO4NOME, CENTRI_TRAPIANTO5.NOME AS CENTRI_TRAPIANTO5NOME, " & _
                "           MON_PAZ_TRAPIANTO.NOTE1 AS NOTE1, MON_PAZ_TRAPIANTO.NOTE2 AS NOTE2, " & _
                "           MON_PAZ_TRAPIANTO.NOTE3 AS NOTE3, MON_PAZ_TRAPIANTO.NOTE4 AS NOTE4, MON_PAZ_TRAPIANTO.NOTE5 AS NOTE5 " & _
                " FROM      (((((((ANAMNESI_NEFROLOGICHE " & _
                "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE) " & _
                "           LEFT OUTER JOIN MON_PAZ_TRAPIANTO ON PAZIENTI.KEY=MON_PAZ_TRAPIANTO.CODICE_PAZIENTE) " & _
                "           LEFT OUTER JOIN CENTRI_TRAPIANTO CENTRI_TRAPIANTO1 ON MON_PAZ_TRAPIANTO.CODICE_CENTRO1=CENTRI_TRAPIANTO1.KEY) " & _
                "           LEFT OUTER JOIN CENTRI_TRAPIANTO CENTRI_TRAPIANTO2 ON MON_PAZ_TRAPIANTO.CODICE_CENTRO2=CENTRI_TRAPIANTO2.KEY) " & _
                "           LEFT OUTER JOIN CENTRI_TRAPIANTO CENTRI_TRAPIANTO3 ON MON_PAZ_TRAPIANTO.CODICE_CENTRO3=CENTRI_TRAPIANTO3.KEY) " & _
                "           LEFT OUTER JOIN CENTRI_TRAPIANTO CENTRI_TRAPIANTO4 ON MON_PAZ_TRAPIANTO.CODICE_CENTRO4=CENTRI_TRAPIANTO4.KEY) " & _
                "           LEFT OUTER JOIN CENTRI_TRAPIANTO CENTRI_TRAPIANTO5 ON MON_PAZ_TRAPIANTO.CODICE_CENTRO5=CENTRI_TRAPIANTO5.KEY) " & _
                " WHERE     (ATTESA_TRAPIANTO=True AND " & _
                "           STATO=0)"
    rsTrapianti.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsTrapianti.EOF
        With rsMain
            .AddNew
            .Fields("COGNOME") = rsTrapianti("COGNOME")
            .Fields("NOME") = rsTrapianti("NOME")
            .Fields("DATA_NASCITA") = rsTrapianti("DATA_NASCITA")
            .Fields("CODICE_FISCALE") = rsTrapianti("CODICE_FISCALE")
            If IsNull(rsTrapianti("CODICE_CENTRO1")) Then
                .Fields("CENTRO_TRAPIANTI") = ""
                .Fields("NOTE") = ""
            Else
                .Fields("CENTRO_TRAPIANTI") = IIf(rsTrapianti("CENTRI_TRAPIANTO1NOME") = Null, "", rsTrapianti("CENTRI_TRAPIANTO1NOME") & vbCrLf & vbCrLf) & IIf(rsTrapianti("CENTRI_TRAPIANTO2NOME") = Null, "", rsTrapianti("CENTRI_TRAPIANTO2NOME") & vbCrLf & vbCrLf) & IIf(rsTrapianti("CENTRI_TRAPIANTO3NOME") = Null, "", rsTrapianti("CENTRI_TRAPIANTO3NOME") & vbCrLf & vbCrLf) & IIf(rsTrapianti("CENTRI_TRAPIANTO4NOME") = Null, "", rsTrapianti("CENTRI_TRAPIANTO4NOME") & vbCrLf & vbCrLf) & IIf(rsTrapianti("CENTRI_TRAPIANTO5NOME") = Null, "", rsTrapianti("CENTRI_TRAPIANTO5NOME") & vbCrLf & vbCrLf)
                .Fields("NOTE") = rsTrapianti("NOTE1") & vbCrLf & vbCrLf & rsTrapianti("NOTE2") & vbCrLf & vbCrLf & rsTrapianti("NOTE3") & vbCrLf & vbCrLf & rsTrapianti("NOTE4") & vbCrLf & vbCrLf & rsTrapianti("NOTE5")
            End If
            
        End With
        rsTrapianti.MoveNext
    Loop
    rsTrapianti.Close
    Set rsTrapianti = Nothing
        
    Set rptTrapianto.DataSource = rsMain
    rptTrapianto.TopMargin = 0
    rptTrapianto.BottomMargin = 0
    rptTrapianto.PrintReport True, rptRangeAllPages
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim i As Integer
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim numKey As Integer
    Dim nomeFile As String
    
    If flxTrova.Row = 0 Then Exit Sub
    
    If modifica Then
        numKey = keyId
    Else
        numKey = GetNumero("MON_PAZ_TRAPIANTO")
    End If
    ' salva i centri nuovi che sono stati inseriti
    For i = 0 To 4
        If cboCentroTrapianti(i).Text <> "" Then
            Call GestisciNuovo("CENTRI_TRAPIANTO", cboCentroTrapianti(i))
        End If
    Next i
    v_Nomi = Array("KEY", "CODICE_PAZIENTE", "NOTE1", "NOTE2", "NOTE3", "NOTE4", _
                    "NOTE5", "CODICE_CENTRO1", "CODICE_CENTRO2", "CODICE_CENTRO3", "CODICE_CENTRO4", "CODICE_CENTRO5")
    v_Val = Array(numKey, flxTrova.TextMatrix(flxTrova.Row, 0), txtNoteTrapianti(0) & "", txtNoteTrapianti(1) & "", txtNoteTrapianti(2) & "", txtNoteTrapianti(3) & "", _
                    txtNoteTrapianti(4) & "", -1, -1, -1, -1, -1)
    For i = 0 To 4
        If cboCentroTrapianti(i).ListIndex <> -1 Then
            v_Val(7 + i) = cboCentroTrapianti(i).ItemData(cboCentroTrapianti(i).ListIndex)
        End If
    Next
                    
    Set rsTrapianti = New Recordset
    If modifica Then
        rsTrapianti.Open "SELECT * FROM MON_PAZ_TRAPIANTO WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsTrapianti.Update v_Nomi, v_Val
        rsTrapianti.Close
    Else
        rsTrapianti.Open "MON_PAZ_TRAPIANTO", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsTrapianti.AddNew v_Nomi, v_Val
        rsTrapianti.Update
        rsTrapianti.Close
    End If
        
    ' controlla eventuali scansioni memorizzate in sospeso
    rsTrapianti.Open "SELECT * FROM SCAN_TRAPIANTI WHERE CODICE_SCHEDA=0", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    Do While Not rsTrapianti.EOF
        rsTrapianti("CODICE_SCHEDA") = numKey
        nomeFile = rsTrapianti("NOME_FILE")
        rsTrapianti("NOME_FILE") = M_TR & numKey & " " & rsTrapianti("CODICE_CENTRO") & " " & Replace(date, "/", "-") & Right(nomeFile, 2)
        rsTrapianti.Update
        If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
            Name structApri.pathDB & "\" & nomeFile & ".jpg" As structApri.pathDB & "\" & M_TR & numKey & " " & rsTrapianti("CODICE_CENTRO") & " " & Replace(date, "/", "-") & Right(nomeFile, 2) & ".jpg"
        ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
            Name structApri.pathDB & "\" & nomeFile & ".pdf" As structApri.pathDB & "\" & M_TR & numKey & " " & rsTrapianti("CODICE_CENTRO") & " " & Replace(date, "/", "-") & Right(nomeFile, 2) & ".pdf"
        End If
        rsTrapianti.MoveNext
    Loop
    rsTrapianti.Close
    Set rsTrapianti = Nothing
    
    Call PulisciTutto
    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
End Sub

Private Sub flxTrova_Click()
    flxTrova.SetFocus
    If VerificaClickFlx(flxTrova) = False Then
        ' discolora
        Call ColoraFlx(flxTrova, flxTrova.Cols - 1, True)
        ' annulla le row e col
        flxTrova.Row = 0
        flxTrova.Col = 0
        Call Pulisci
        Exit Sub
    Else
        Call ColoraFlx(flxTrova, flxTrova.Cols - 1)
        Call CaricaScheda
    End If
End Sub

Private Sub cboCentroTrapianti_Click(Index As Integer)
    ' non puo selezionare un centro gia esistente
    Dim i As Integer
    If cboCentroTrapianti(Index).ListIndex = -1 Then Exit Sub
    For i = 0 To 4
        If cboCentroTrapianti(Index).Text = cboCentroTrapianti(i) And Index <> i Then
            MsgBox "Non è possibile selezionare un centro medico presente nella scheda", vbCritical, "Attenzione"
            cboCentroTrapianti(Index).ListIndex = -1
            Exit Sub
        End If
    Next i
End Sub

Private Sub cboCentroTrapianti_KeyPress(Index As Integer, KeyAscii As Integer)
    If Len(cboCentroTrapianti(Index).Text) >= 30 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtNoteTrapianti_GotFocus(Index As Integer)
    txtNoteTrapianti(Index).BackColor = colArancione
End Sub

Private Sub txtNoteTrapianti_LostFocus(Index As Integer)
    txtNoteTrapianti(Index).BackColor = vbWhite
End Sub

