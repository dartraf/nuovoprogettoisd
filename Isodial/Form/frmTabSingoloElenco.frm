VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTabSingoloElenco 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGriglia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   9495
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
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   7200
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   16776960
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         MousePointer    =   99
         FormatString    =   "| Tabella                                                                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTabSingoloElenco.frx":0000
      End
   End
   Begin VB.Frame fraPulsanti 
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
      TabIndex        =   7
      Top             =   3120
      Width           =   9495
      Begin VB.Frame fraPulsantiInterno 
         BorderStyle     =   0  'None
         Height          =   640
         Left            =   3480
         TabIndex        =   8
         Top             =   120
         Width           =   5895
         Begin VB.CommandButton cmdElimina 
            Caption         =   "&Elimina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            TabIndex        =   1
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnnulla 
            Caption         =   "&Annulla"
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
            Height          =   480
            Left            =   3120
            TabIndex        =   3
            Top             =   120
            Width           =   1215
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
            Height          =   480
            Left            =   1680
            TabIndex        =   2
            Top             =   120
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
            Height          =   480
            Left            =   4560
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmTabSingoloElenco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intTipoTabSingolo As enumTipoTabSingolo
Dim strNomeTabella As String
Dim strNomeElemento As String
Dim objAnnulla As CAnnulla
Dim intRow As Integer
Dim intCol As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
    Dim i As Integer
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    Me.Height = GetSetting(appName, "Forms", Me.Name & intTipoTabSingolo & ".Height", 4440)
    Me.Width = GetSetting(appName, "Forms", Me.Name & intTipoTabSingolo & ".Width", 9825)
    
    
    Set objAnnulla = New CAnnulla

    With flxGriglia
        .Rows = 1
    
        .Cols = 2
        Select Case intTipoTabSingolo
            Case enumTipoTabSingolo.AGO
                strNomeTabella = "AGO"
                strNomeElemento = "L'ago"
                Me.Caption = "Tabella: Aghi"
                .TextMatrix(0, 1) = "Ago"
            Case enumTipoTabSingolo.ANTICOAGULANTI
                strNomeTabella = "ANTICOAGULANTI"
                strNomeElemento = "L'anticoagulante"
                Me.Caption = "Tabella: Anticoagulanti"
                .TextMatrix(0, 1) = "Anticoagulante"
            Case enumTipoTabSingolo.filtro
                strNomeTabella = "FILTRI"
                strNomeElemento = "Il filtro"
                Me.Caption = "Tabella: Filtri"
                .TextMatrix(0, 1) = "Filtro"
            Case enumTipoTabSingolo.LINEE
                strNomeTabella = "LINEE"
                strNomeElemento = "La linea"
                Me.Caption = "Tabella: Linee"
                .TextMatrix(0, 1) = "Linea"
            Case enumTipoTabSingolo.Medicinali
                strNomeTabella = "MEDICINALI"
                strNomeElemento = "Il farmaco"
                Me.Caption = "Tabella: Farmaci"
                .TextMatrix(0, 1) = "Farmaco"
            Case enumTipoTabSingolo.ORGANO
                strNomeTabella = "ORGANI"
                strNomeElemento = "L'organo"
                Me.Caption = "Tabella: Organi"
                .TextMatrix(0, 1) = "Organo"
            Case enumTipoTabSingolo.TITOLIDIARIO
                strNomeTabella = "TITOLI_DIARIO"
                strNomeElemento = "Il titolo di diario clinico"
                Me.Caption = "Tabella: Titoli di diario clinico"
                .TextMatrix(0, 1) = "Titolo"
        End Select
    
        .ColWidth(0) = 0
        .Row = 0
        .MousePointer = flexCustom
        For i = 0 To flxGriglia.Cols - 1
            .Col = i
            .ColAlignment(i) = vbLeftJustify
            .CellFontBold = True
        Next i
    End With
    
    Call CaricaFlx
End Sub

Private Sub Form_Resize()
    If Me.Width <= 9825 Then Me.Width = 9825
    If Me.Height <= 2580 Then Me.Height = 2580

    fraGriglia.Height = Me.Height - fraPulsanti.Height - 360
    fraGriglia.Width = Me.Width - 340
    flxGriglia.Height = fraGriglia.Height - 360
    flxGriglia.Width = fraGriglia.Width - 240
    
    fraPulsanti.Width = Me.Width - 340
    fraPulsanti.Top = fraGriglia.Top + fraGriglia.Height - 150
    
    fraPulsantiInterno.Left = fraPulsanti.Width - fraPulsantiInterno.Width - 50
    
    Call AutoResizeGrid(flxGriglia, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(appName, "Forms", Me.Name & intTipoTabSingolo & ".Width", Me.Width)
    Call SaveSetting(appName, "Forms", Me.Name & intTipoTabSingolo & ".Height", Me.Height)
    Set objAnnulla = Nothing
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
'-----------------------------------

Private Sub CaricaFlx()
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    flxGriglia.Rows = 1
    flxGriglia.Redraw = False
    
    strSql = "SELECT * FROM " & strNomeTabella & " ORDER BY NOME"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("NOME") & ""
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    Call AutoResizeGrid(flxGriglia, Me, 1)
    
    flxGriglia.Row = 0
    flxGriglia.Redraw = True
    Set rsDataset = Nothing
End Sub

Private Sub SalvaModifiche()
    Dim rsDataset As New Recordset
    Dim intKey As Integer
    Dim intTipoInfermiere As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant

    With flxGriglia
        intKey = .TextMatrix(intRow, 0)
        v_Nomi = Array("KEY", "NOME")
        v_Val = Array(intKey, .TextMatrix(intRow, 1))
        
        rsDataset.Open "SELECT * FROM " & strNomeTabella & " WHERE KEY=" & intKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsDataset.Update v_Nomi, v_Val
        rsDataset.Close
    End With
    
    Set rsDataset = Nothing
End Sub

Private Function VerificaDuplicato() As Boolean
    Dim rsDataset As New Recordset
    Dim strSql As String
    Dim strNome As String
    
    strNome = txtAppo.Text
    
    strSql = "Select    count(Key) as Totale " & _
            "From " & strNomeTabella & " " & _
            "Where      Nome  like '" & Apostrophe(strNome) & "'"
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly
    
    If rsDataset("Totale") <> 0 Then
        VerificaDuplicato = True
    Else
        VerificaDuplicato = False
    End If
    
    rsDataset.Close
    Set rsDataset = Nothing
End Function

Private Sub cmdElimina_Click()
    Dim blnElimina As Boolean
    Dim intKey As Integer
    Dim strNome As String
    Dim cmCommand As Command
   
    With flxGriglia
        If .Row = 0 Then
            MsgBox "Selezionare " & LCase(strNomeElemento) & " da eliminare", vbCritical, "Attenzione"
        Else
            intKey = .TextMatrix(.Row, 0)
            strNome = .TextMatrix(.Row, 1)
            blnElimina = False

            Select Case intTipoTabSingolo
                Case AGO
                    blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "AGO1", intKey)
                    If blnElimina Then
                        blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "AGO2", intKey)
                    End If
                Case ANTICOAGULANTI
                    blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "ANTICOAGULANTE1", intKey)
                    If blnElimina Then
                        blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "ANTICOAGULANTE2", intKey)
                    End If
                Case filtro
                    blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "TIPO_FILTRO", intKey)
                Case LINEE
                    blnElimina = IsPossibleDelete("ANAMNESI_DIALITICHE", "TIPO_LINEE", intKey)
                Case TITOLIDIARIO
                    blnElimina = IsPossibleDelete("DIARI_CLINICI", "CODICE_TITOLO", intKey)
                Case ORGANO
                    blnElimina = IsPossibleDelete("ESAMI_STRUMENTALI", "CODICE_ORGANO", intKey)
                Case Medicinali
                    blnElimina = IsPossibleDelete("TERAPIE_DIALITICHE", "CODICE_MEDICINALE", intKey)
                    If blnElimina Then
                        blnElimina = IsPossibleDelete("TERAPIE_DOMICILIARI", "CODICE_MEDICINALE", intKey)
                    End If
            End Select
                
            If blnElimina Then
                If MsgBox("Sicuro di voler eliminare " & strNome & "?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                    Set cmCommand = New Command
                    cmCommand.ActiveConnection = cnPrinc
                    cmCommand.CommandType = adCmdText
                    
                    If intTipoTabSingolo = ORGANO Then
                        cmCommand.CommandText = "Delete From Esami Where Codice_Organo=" & intKey
                        cmCommand.Execute
                    End If
                    
                    cmCommand.CommandText = "Delete From " & strNomeTabella & " Where Key=" & intKey
                    cmCommand.Execute
                    
                    Set cmCommand = Nothing
                    
                    ' rimuove dalla flx
                    If .Rows = 2 Then
                        .Rows = 1
                    Else
                        .RemoveItem (.Row)
                    End If
                    intRow = 0
                    .Row = 0
                    MsgBox "Eliminazione avvenuta con successo", vbInformation, Me.Caption
                End If
            Else
                MsgBox "Impossibile eliminare " & strNome & " perchè in relazione con altri dati del sistema", vbInformation, Me.Caption
            End If
        End If
    End With
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
                intRow = i
                Call SalvaModifiche
                If objAnnulla.Vuoto = True Then
                    cmdAnnulla.Enabled = False
                End If
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub cmdInserisci_Click()
    Dim lfrmTabSingoloInput As New frmTabSingoloInput
    lfrmTabSingoloInput.intTipoTabSingolo = intTipoTabSingolo
    lfrmTabSingoloInput.Show 1
    If lfrmTabSingoloInput.blnRefresh Then
        Call CaricaFlx
        
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, 0, lfrmTabSingoloInput.intIDInserito)
        intRow = flxGriglia.Row
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        If flxGriglia.Row > Int(flxGriglia.Height / flxGriglia.CellHeight) - 3 Then
            flxGriglia.TopRow = flxGriglia.Row
        End If
    End If
    Unload lfrmTabSingoloInput
    Set lfrmTabSingoloInput = Nothing
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub flxGriglia_Click()
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        intRow = flxGriglia.Row
        intCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intNumeroRigheVisibili As Integer
    
    If flxGriglia.Rows = 1 Then Exit Sub
    
    intNumeroRigheVisibili = Int(flxGriglia.Height / flxGriglia.CellHeight) - 3
    
    If flxGriglia.Row = flxGriglia.Rows - 1 Then
        i = 1
    Else
        i = flxGriglia.Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia.TextMatrix(i, 1), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia.Row = i
            If i >= intNumeroRigheVisibili Or flxGriglia.TopRow > intNumeroRigheVisibili Then
                flxGriglia.TopRow = i
                Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            End If
            Exit Do
        End If
        If i = flxGriglia.Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxGriglia.Row
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub flxGriglia_DblClick()
    ' fase di modifica
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        txtAppo.Left = .colPos(intCol) + .Left + 45
        txtAppo.Top = .rowPos(intRow) + .Top + 45
        txtAppo.Width = .ColWidth(intCol)
        txtAppo.Text = .TextMatrix(intRow, intCol)
        txtAppo.Visible = True
        txtAppo.SetFocus
    End With
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
    txtAppo.MaxLength = 50
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    txtAppo.Visible = False
    
    If (flxGriglia.TextMatrix(intRow, intCol)) <> (txtAppo.Text) Then
        
        If txtAppo = "" Then
            MsgBox "Impossibile memorizzare dati vuoti", vbCritical, "Attenzione"
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            flxGriglia.TopRow = intRow
        Else
            If VerificaDuplicato() Then
                MsgBox strNomeElemento & " è gia presente in archivio.", vbExclamation, Me.Caption
            Else
                Call objAnnulla.Add(flxGriglia.TextMatrix(intRow, intCol), intCol, Int(flxGriglia.TextMatrix(intRow, 0)))
                cmdAnnulla.Enabled = True
                flxGriglia.TextMatrix(intRow, intCol) = txtAppo.Text
                Call SalvaModifiche
            End If
        End If
    End If
End Sub

