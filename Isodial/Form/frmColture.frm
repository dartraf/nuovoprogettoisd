VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmColture 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colture positive dell'acqua e bagno dialisi"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   Icon            =   "frmColture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtAppo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   3960
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         FormatString    =   $"frmColture.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmColture.frx":00CF
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
      TabIndex        =   3
      Top             =   3360
      Width           =   12015
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
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Annulla digitazione"
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
         Left            =   8040
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Inserisci"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1335
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
         Left            =   10680
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmColture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmColture.frm
'
' <b>Descrizione</b>: Scheda Colture associata alla tabella COLTURE
'
' @remarks
'
' @author
'
' @date 04/02/2011 19.43
Option Explicit

Dim vRow As Integer
Dim vCol As Integer
'' rs della scheda
Dim rsColture As Recordset
'' obj per la lista CAnnulla
Dim objAnnulla As CAnnulla
'' rs per la tracciatura
Dim rsDisco As Recordset
Const pos As String = "  POS  "
Const NEG As String = "  NEG  "

Private Sub cmdStampa_Click()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim i As Integer
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(15) AS DATA, " & _
                "       NEW adVarChar(30) AS COLTURE_ACQUA, " & _
                "       NEW adVarChar(10)  AS ESITO_COLTURE_ACQUA, " & _
                "       NEW adVarChar(30) AS COLTURE_BAGNO, " & _
                "       NEW adVarChar(10)  AS ESITO_COLTURE_BAGNO "
                
                
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
                   
    With rsMain
        For i = 1 To flxGriglia.Rows - 1
            .AddNew
            .Fields("DATA") = flxGriglia.TextMatrix(i, 1)
            .Fields("COLTURE_ACQUA") = flxGriglia.TextMatrix(i, 2)
            .Fields("ESITO_COLTURE_ACQUA") = flxGriglia.TextMatrix(i, 3)
            .Fields("COLTURE_BAGNO") = flxGriglia.TextMatrix(i, 4)
            .Fields("ESITO_COLTURE_BAGNO") = flxGriglia.TextMatrix(i, 5)
        Next i
    End With
        
    Set rptColture.DataSource = rsMain
    rptColture.TopMargin = 0
    rptColture.BottomMargin = 0
    rptColture.PrintReport True, rptRangeAllPages
    
End Sub

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft

    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 5
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
    End With
    Set objAnnulla = New CAnnulla
    Call ApriRsDisconnesso
    Call CaricaScheda
    
    If flxGriglia.Rows <= 1 Then
        cmdStampa.Enabled = False
    End If
    
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "COLTURE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Set rsDisco = New ADODB.Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsDisco.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsDisco.CursorLocation = adUseClient
    rsDisco.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

'' Confronta i campi per rilevare le eventuali modifiche
' e le salva nella relativa tabella delle modifiche
'
' @param rs rs che contiene lo stato del record che si è memorizzato
Private Sub Confronta(rs As Recordset)
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim v_modifiche() As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim nome_campi As String
    Dim valori As String
    Dim trovato As Boolean
    
    ReDim v_modifiche(0)
    ' filtra per la presenza di piu record
    rsDisco.Filter = "(KEY=" & rs("KEY") & ")"
    For i = 0 To rsDisco.Fields.count - 1
        trovato = False
        If IsNull(rsDisco(i)) Or IsNull(rs(i)) Then
            If Not (IsNull(rsDisco(i)) And IsNull(rs(i))) Then
                trovato = True
            End If
        Else
            If rsDisco(i) <> rs(i) Then
                trovato = True
            End If
        End If
        If trovato Then
            ReDim Preserve v_modifiche(UBound(v_modifiche) + 1)
            v_modifiche(UBound(v_modifiche)) = i
        End If
    Next i
    If UBound(v_modifiche) <> 0 Then
        For i = 1 To UBound(v_modifiche)
            nome_campi = nome_campi & rsDisco.Fields((v_modifiche(i))).Name & "&-&"
            valori = valori & IIf(IsNull(rsDisco.Fields((v_modifiche(i)))) Or rsDisco.Fields((v_modifiche(i))) = "", "NULL", rsDisco.Fields((v_modifiche(i)))) & "&-&"
            ' aggiorna il rsDisco
            rsDisco(v_modifiche(i)) = rs(v_modifiche(i))
        Next i
        nome_campi = Left(nome_campi, Len(nome_campi) - 3)
        valori = Left(valori, Len(valori) - 3)
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, rs("KEY"), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_COLTURE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Salva le modifica di un campo di un record
Private Sub SalvaModifiche()
    Dim valore As Variant
    Dim nome As Variant
    Select Case vCol
        Case 1
            nome = "DATA"
        Case 2
            nome = "COLTURA_ACQUA"
        Case 3
            nome = "ESITO_ACQUA"
        Case 4
            nome = "COLTURA_BAGNO"
        Case 5
            nome = "ESITO_BAGNO"
    End Select
    If vCol = 2 Or vCol = 4 Then
        valore = flxGriglia.TextMatrix(vRow, vCol)
    Else
        If flxGriglia.TextMatrix(vRow, vCol) = "" Then
            valore = 0
        ElseIf flxGriglia.TextMatrix(vRow, vCol) = pos Then
            valore = 1
        Else
            valore = 2
        End If
    End If
    Set rsColture = New Recordset
    rsColture.Open "SELECT * FROM COLTURE WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    rsColture.Update nome, valore
    If TRACCIATO Then
        Call Confronta(rsColture)
    End If
    Set rsColture = Nothing
End Sub

'' Carica i dati della tabella nel form e nell rsDisco
Private Sub CaricaScheda()
    Dim i As Integer
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    vRow = 0
    vCol = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    Set rsColture = New Recordset
    rsColture.Open "COLTURE" & " ORDER BY DATA DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If rsColture.EOF And rsColture.BOF Then
    Else
        Do While Not rsDisco.EOF
            rsDisco.Delete
            rsDisco.MoveNext
        Loop
        Do While Not rsColture.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Rows - 1, 0) = rsColture("KEY")
                .TextMatrix(.Rows - 1, 1) = rsColture("DATA")
                .TextMatrix(.Rows - 1, 2) = rsColture("COLTURA_ACQUA") & ""
                .Col = 3
                .CellForeColor = IIf(rsColture("ESITO_ACQUA") = 1, vbRed, vbBlack)
                .TextMatrix(.Rows - 1, 3) = Choose(rsColture("ESITO_ACQUA") + 1, "", pos, NEG)
                .TextMatrix(.Rows - 1, 4) = rsColture("COLTURA_BAGNO") & ""
                .Col = 5
                .CellForeColor = IIf(rsColture("ESITO_BAGNO") = 1, vbRed, vbBlack)
                .TextMatrix(.Rows - 1, 5) = Choose(rsColture("ESITO_BAGNO") + 1, "", pos, NEG)
            End With
            
            ' aggiorna i dati nel rsDisco
            rsDisco.AddNew
            For i = 0 To rsDisco.Fields.count - 1
                rsDisco.Fields(i) = rsColture.Fields(i)
            Next i
            rsDisco.Update
        
            rsColture.MoveNext
        Loop
    End If
    Set rsColture = Nothing
    flxGriglia.Row = 0
End Sub

'' Inserisce un nuovo record
Private Sub cmdInserisci_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim num As Integer
    Unload frmInput
    tInput.Tipo = tpICOLTURE
    frmInput.Show 1
    If Not (tInput.v_valori(1) = "") Then
        If Not (tInput.v_valori(2) = "" And CInt(tInput.v_valori(3)) = 0 And tInput.v_valori(4) = "" And CInt(tInput.v_valori(5)) = 0) Then
            v_Nomi = Array("KEY", "DATA", "COLTURA_ACQUA", "ESITO_ACQUA", "COLTURA_BAGNO", "ESITO_BAGNO")
            num = GetNumero("COLTURE")
            v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(4), tInput.v_valori(5))
            Set rsColture = New Recordset
            rsColture.Open "COLTURE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsColture.AddNew v_Nomi, v_Val
            rsColture.Update
            Set rsColture = Nothing
            
            ' aggiorna i dati nel rsDisco
            rsDisco.AddNew v_Nomi, v_Val
            rsDisco.Update
                    
            ' aggiorna la flx
            Call CaricaScheda
            
            ' si posiziona sul record e lo seleziona
            flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            If flxGriglia.Row > 11 Then
                flxGriglia.TopRow = flxGriglia.Row
            End If
            
'            MsgBox "Inserimento effettuato", vbInformation, "Inserimento"
            cmdStampa.Enabled = True
            
        End If
    End If
End Sub

Private Sub cmdAnnulla_Click()
    Dim Dato As String
    Dim Col As Integer
    Dim RowKey As Integer
    Dim i As Integer
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    RowKey = objAnnulla.Row
    ' cerca la riga con il key memorizzato in rowkey
    With flxGriglia
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = RowKey Then
                ' annulla
                .TextMatrix(i, Col) = Dato
                objAnnulla.Remove
                ' modifica anche il db
                vRow = i
                vCol = Col
                Call SalvaModifiche
                If objAnnulla.Vuoto = True Then
                    cmdAnnulla.Enabled = False
                End If
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub flxGriglia_Click()
    vCol = flxGriglia.Col
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        vRow = flxGriglia.Row
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        If .Col = 1 Then
            frmCalendario.Show 1
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
            Call SalvaModifiche
            ' cambia colonna per evitave di ricaricare il calendario
            .Col = 0
        ElseIf .Col = 3 Or .Col = 5 Then
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            ' esito puo esser solo pos o neg
            If .TextMatrix(.Row, .Col) = pos Then
                '.Row = vRow
                '.CellForeColor = vbBlack
                .TextMatrix(.Row, .Col) = NEG
            Else
                '.Row = vRow
                '.CellForeColor = vbRed
                .TextMatrix(.Row, .Col) = pos
            End If
            Call SalvaModifiche
        Else
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        End If
    End With
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    If UCase(flxGriglia.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
        Call SalvaModifiche
    End If
    txtAppo.Visible = False
End Sub

