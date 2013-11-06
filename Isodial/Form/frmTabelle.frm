VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmTabelle 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4140
   ClientLeft      =   630
   ClientTop       =   1515
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboAppo 
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
      Left            =   3600
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
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
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
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
         TabIndex        =   7
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         MouseIcon       =   "frmTabelle.frx":0000
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
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   9495
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
         Height          =   600
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
         Height          =   600
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
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
         Height          =   600
         Left            =   4680
         TabIndex        =   1
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
         Height          =   600
         Left            =   8160
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraListaSec 
      Caption         =   "Organi/Apparati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   9495
      Begin VB.ListBox lstOrgani 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   9255
      End
   End
End
Attribute VB_Name = "frmTabelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim keyOrgano As Integer              ' codice organo
Dim rsTabelle As Recordset
Dim nomeTabella As String
Dim vRow As Integer             ' riga selezionata
Dim vCol As Integer             ' colonna selezionata
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx
Dim lettera As String * 1

Const ESENTE As String = "ESENTE"           ' true
Const NONESENTE As String = "NON ESENTE"    ' false

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'---------------------------

Private Sub Form_Activate()
    Dim nomeTabella As String
    
    If Not RidisponiForms(Me) Then Exit Sub
    
    If tTabelle = tpCOMUNI Or tTabelle = tpasl Or tTabelle = tpDISTRETTI Then
        ' deve ricaricare la cbo
        If tTabelle = tpCOMUNI Or tTabelle = tpasl Then
            nomeTabella = "REGIONI"
        Else
            nomeTabella = "ASL"
        End If
        
        Call RicaricaComboBox(nomeTabella, "NOME", cboAppo)
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    Set objAnnulla = New CAnnulla
    flxGriglia.Rows = 1
    
    If tTabelle >= tpRegioni And tTabelle <= tpEDTA Then
        ' tipo con due informazioni
        flxGriglia.Cols = 3
        Select Case tTabelle
            Case tpRegioni
                nomeTabella = "REGIONI"
                Me.Caption = Me.Caption & "Regioni"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) * 2 / 2
                    .ColWidth(2) = .ColWidth(1) * 2 / 2
                    .TextMatrix(0, 1) = "Codice Regione"
                    .TextMatrix(0, 2) = "Regione"
                End With
                cmdElimina.Visible = False
            Case tpTIPOLOGIEMEDICO
                nomeTabella = "TIPOLOGIE_MEDICO"
                Me.Caption = Me.Caption & "Tipologie Medico"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) / 5
                    .ColWidth(2) = .ColWidth(1) * 9
                    .TextMatrix(0, 1) = "Codice Tipologia"
                    .TextMatrix(0, 2) = "Tipologia Medico"
                End With
                cmdElimina.Visible = False
            Case tpESENZIONI
                nomeTabella = "TIPOLOGIE_ESENZIONE"
                Me.Caption = Me.Caption & "Codici di esenzione"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) * 2 / 2
                    .ColWidth(2) = .ColWidth(1) * 2 / 2
                    .TextMatrix(0, 1) = "Codice Esenzione"
                    .TextMatrix(0, 2) = "Quota Regionale su Ricetta"
                End With
                cmdElimina.Visible = False
            Case tpEDTA
                nomeTabella = "EDTA"
                Me.Caption = Me.Caption & "E.D.T.A."
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) / 5
                    .ColWidth(2) = .ColWidth(1) * 10
                    .TextMatrix(0, 1) = "Cod."
                    .TextMatrix(0, 2) = "E.D.T.A."
                End With
        End Select
        
        Call CaricaFlx
        
        ElseIf tTabelle >= tpCOMUNI And tTabelle <= tpDISTRETTI Then
        ' tipo con tre informazioni
        flxGriglia.Cols = 4
        Select Case tTabelle
           Case tpCOMUNI
                nomeTabella = "COMUNI"
                Me.Caption = Me.Caption & "Comuni"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) / 2
                    .ColWidth(2) = .ColWidth(1) * 3 / 2 + 1000
                    .ColWidth(3) = .ColWidth(1) * 3 / 2
                    .TextMatrix(0, 1) = "Codice ISTAT"
                    .TextMatrix(0, 2) = "Comune"
                    .TextMatrix(0, 3) = "Regione"
                End With
                cmdElimina.Visible = False
            Case tpasl
                nomeTabella = "ASL"
                Me.Caption = Me.Caption & "ASL"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) / 2
                    .ColWidth(2) = .ColWidth(1) * 3 / 2
                    .ColWidth(3) = .ColWidth(1) * 3 / 2
                    .TextMatrix(0, 1) = "Codice ASL"
                    .TextMatrix(0, 2) = "ASL"
                    .TextMatrix(0, 3) = "Regione"
                End With
                cmdElimina.Visible = False
            Case tpDISTRETTI
                nomeTabella = "DISTRETTI"
                Me.Caption = Me.Caption & "Distretti"
                With flxGriglia
                    .ColWidth(1) = .ColWidth(1) / 2
                    .ColWidth(2) = .ColWidth(1) * 3 / 2
                    .ColWidth(3) = .ColWidth(1) * 3 / 2
                    .TextMatrix(0, 1) = "Codice Distretto"
                    .TextMatrix(0, 2) = "Distretto"
                    .TextMatrix(0, 3) = "Asl di riferimento"
                End With
                cmdElimina.Visible = False
        End Select
        
        Call CaricaFlx
        
    Else
        ' altro
        Select Case tTabelle
            Case tpESAME
                flxGriglia.ColWidth(1) = flxGriglia.ColWidth(1) * 2
                nomeTabella = "ESAMI"
                Me.Caption = Me.Caption & "Tipo di esame"
                fraListaMain.Caption = "Esami"
                ' allunga i frame
                fraListaSec.Height = 1575
                lstOrgani.Height = 1300
                flxGriglia.Height = 3375
                fraListaMain.Height = 3855
                ' sposta i frame
                fraListaSec.Top = fraListaMain.Top
                fraListaMain.Top = fraListaSec.Top + fraListaSec.Height
                fraAzioni.Top = fraListaMain.Top + fraListaMain.Height - 160
                Me.Height = fraAzioni.Top + fraAzioni.Height + 500
                fraListaSec.Visible = True
                ' carica solo la lista degli organi
                Set rsTabelle = New Recordset
                rsTabelle.Open "ORGANI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
                ' carica la seconda lista
                Do While Not rsTabelle.EOF
                    lstOrgani.AddItem rsTabelle("NOME")
                    rsTabelle.MoveNext
                Loop
                Set rsTabelle = Nothing
                cmdElimina.Visible = False
            Case tpNOMENCLATORE
                nomeTabella = "NOMENCLATORE_TARIFFARIO"
                Me.Caption = Me.Caption & "Nomenclatore Tariffario"
                With flxGriglia
                    .Cols = 5
                    .ColWidth(1) = .ColWidth(1) / 2
                    .ColWidth(2) = .ColWidth(1) * 2
                    .ColWidth(3) = .ColWidth(1) / 2
                    .ColWidth(4) = .ColWidth(1) / 1
                    .TextMatrix(0, 1) = "Codice"
                    .TextMatrix(0, 2) = "Descrizione Prestazione"
                    .TextMatrix(0, 3) = "Importo"
                    .TextMatrix(0, 4) = "Scontato"
                End With
                Call CaricaFlx
                cmdElimina.Visible = False
            Case tpRENI
                nomeTabella = "APPARATI"
                Me.Caption = Me.Caption & "Gestione Reni"
                cmdInserisci.Visible = False
                cmdElimina.Visible = False
                With flxGriglia
                    .Cols = 7
                    .ColWidth(1) = .ColWidth(1) * 1 / 2.8
                    .ColWidth(2) = .ColWidth(1) * 0.7
                    .ColWidth(3) = .ColWidth(1) * 1.9
                    .ColWidth(4) = .ColWidth(1) * 0.9
                    .ColWidth(5) = .ColWidth(1) * 0.8
                    .ColWidth(6) = .ColWidth(1) * 0.9
                    
                    
                    .TextMatrix(0, 1) = "Postazione"
                    .TextMatrix(0, 2) = "N° rene"
                    .TextMatrix(0, 3) = "Monitor"
                    .TextMatrix(0, 4) = "Matricola"
                    .TextMatrix(0, 5) = "Tipo"
                    .TextMatrix(0, 6) = "Dt.Rottam."
                    
                End With
                Call CaricaFlx
        End Select
    End If
    
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
    Dim strSql As String
    
    flxGriglia.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    
    strSql = "SELECT * FROM " & nomeTabella
    If tTabelle = tpESENZIONI Or tTabelle = tpDISTRETTI Or tTabelle = tpNOMENCLATORE Or tTabelle = tpTIPOLOGIEMEDICO Then
        strSql = strSql & " WHERE (NOT KEY=-1) ORDER BY CODICE"
    ElseIf tTabelle = tpRENI Then
        strSql = strSql & " WHERE TIPO_APPARATO = 'RENE ARTIFICIALE' ORDER BY SOSTITUITO DESC, DATA_ROTTAMAZIONE DESC, POSTAZIONE"
    Else
        strSql = strSql & " ORDER BY NOME"
    End If
    
    Set rsTabelle = New Recordset
    rsTabelle.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ' carica la lista
    If Not (rsTabelle.EOF And rsTabelle.BOF) Then
        Do While Not rsTabelle.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsTabelle("KEY")
                
                If (tTabelle >= tpRegioni And tTabelle <= tpTIPOLOGIEMEDICO) Or (tTabelle = tpEDTA) Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NOME") & ""
                ElseIf tTabelle = tpCOMUNI Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NOME") & ""
                    .TextMatrix(.Rows - 1, 3) = GetNome(rsTabelle("REGIONIID"), "REGIONI")
                ElseIf tTabelle = tpasl Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NOME") & ""
                    .TextMatrix(.Rows - 1, 3) = GetNome(rsTabelle("CODICE_REGIONE"), "REGIONI")
                ElseIf tTabelle = tpESENZIONI Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = IIf(CBool(rsTabelle("ESENZIONE_QUOTA")), ESENTE, NONESENTE)
                ElseIf tTabelle = tpDISTRETTI Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NOME") & ""
                    .TextMatrix(.Rows - 1, 3) = GetNome(rsTabelle("CODICE_ASL"), "ASL")
                ElseIf tTabelle = tpNOMENCLATORE Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("CODICE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NOME") & ""
                    .TextMatrix(.Rows - 1, 3) = VirgolaOrPunto(rsTabelle("IMPORTO"), ",")
                    .TextMatrix(.Rows - 1, 4) = VirgolaOrPunto(rsTabelle("IMPORTO_SCONTATO"), ",")
                ElseIf tTabelle = tpRENI Then
                    .TextMatrix(.Rows - 1, 1) = rsTabelle("POSTAZIONE")
                    '.TextMatrix(.Rows - 1, 2) = rsTabelle("NUMERO_RENE") & ""
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("NUMERO_APPARATO") & ""
                    '.TextMatrix(.Rows - 1, 3) = rsTabelle("TIPO_RENE") & ""
                    .TextMatrix(.Rows - 1, 3) = rsTabelle("MODELLO") & ""
                    .TextMatrix(.Rows - 1, 4) = rsTabelle("MATRICOLA") & ""
                    If rsTabelle("TIPO") = 0 Then
                        .TextMatrix(.Rows - 1, 5) = "NEG"
                    ElseIf rsTabelle("TIPO") = 1 Then
                        .TextMatrix(.Rows - 1, 5) = "HCV POS"
                    Else
                        .TextMatrix(.Rows - 1, 5) = "HBV POS"
                    End If
                    .Col = 6
                    .Row = .Rows - 1
                    .CellAlignment = vbRightJustify
                    .CellForeColor = vbRed
                    .TextMatrix(.Rows - 1, 6) = rsTabelle("DATA_ROTTAMAZIONE") & ""
                End If

                rsTabelle.MoveNext
            End With
        Loop
    End If
    Set rsTabelle = Nothing
    flxGriglia.Row = 0
End Sub

Private Sub SalvaModifiche()
    Dim keyId As Integer
    Dim tipoInfermiere As Integer
    Dim valore As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant

    With flxGriglia
        keyId = .TextMatrix(vRow, 0)
        
        If (tTabelle >= tpRegioni And tTabelle <= tpTIPOLOGIEMEDICO) Or (tTabelle = tpEDTA) Then
            v_Nomi = Array("KEY", "NOME", "CODICE")
            v_Val = Array(keyId, .TextMatrix(vRow, 2), .TextMatrix(vRow, 1))
        ElseIf tTabelle = tpCOMUNI Then
            v_Nomi = Array("KEY", "NOME", "CODICE", "REGIONIID")
            v_Val = Array(keyId, .TextMatrix(vRow, 2), .TextMatrix(vRow, 1), GetNumeroDaNome("REGIONI", "NOME", .TextMatrix(vRow, 3)))
        ElseIf tTabelle = tpasl Then
            v_Nomi = Array("KEY", "NOME", "CODICE", "CODICE_REGIONE")
            v_Val = Array(keyId, .TextMatrix(vRow, 2), .TextMatrix(vRow, 1), GetNumeroDaNome("REGIONI", "NOME", .TextMatrix(vRow, 3)))
        ElseIf tTabelle = tpESENZIONI Then
            v_Nomi = Array("KEY", "CODICE", "ESENZIONE_QUOTA")
            v_Val = Array(keyId, .TextMatrix(vRow, 1), IIf(.TextMatrix(vRow, 2) = ESENTE, True, False))
        ElseIf tTabelle = tpDISTRETTI Then
            v_Nomi = Array("KEY", "CODICE", "NOME", "CODICE_ASL")
            v_Val = Array(keyId, .TextMatrix(vRow, 1), .TextMatrix(vRow, 2), GetNumeroDaNome("ASL", "NOME", .TextMatrix(vRow, 3)))
        ElseIf tTabelle = tpESAME Then
            v_Nomi = Array("KEY", "NOME", "CODICE_ORGANO")
            v_Val = Array(keyId, .TextMatrix(vRow, 1), keyOrgano)
        ElseIf tTabelle = tpNOMENCLATORE Then
            v_Nomi = Array("KEY", "CODICE", "NOME", "IMPORTO", "IMPORTO_SCONTATO")
            v_Val = Array(keyId, .TextMatrix(vRow, 1), .TextMatrix(vRow, 2), .TextMatrix(vRow, 3), .TextMatrix(vRow, 4))
        ElseIf tTabelle = tpRENI Then
            v_Nomi = Array("KEY", "POSTAZIONE", "NUMERO_APPARATO", "MODELLO", "MATRICOLA", "TIPO", "DATA_ROTTAMAZIONE")
            If .TextMatrix(vRow, 5) = "HCV POS" Then
                valore = 1
            ElseIf .TextMatrix(vRow, 5) = "HBV POS" Then
                valore = 2
            Else
                valore = 0
            End If
            v_Val = Array(keyId, .TextMatrix(vRow, 1), .TextMatrix(vRow, 2), .TextMatrix(vRow, 3), .TextMatrix(vRow, 4), valore, IIf(.TextMatrix(vRow, 6) = "", Null, .TextMatrix(vRow, 6)))
        End If
        
        Set rsTabelle = New Recordset
        rsTabelle.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsTabelle.Update v_Nomi, v_Val
        Set rsTabelle = Nothing
        ' aggiorna la flx
        If tTabelle = tpESAME Then
            lstOrgani_Click
        ElseIf tTabelle = tpNOMENCLATORE And (vCol = 3 Or vCol = 4) Then
            Call AggiornaRicette(keyId)
        End If
    End With
End Sub

Private Sub AggiornaRicette(keyId As Integer)
    If MsgBox("Vuoi rielaborare le prestazioni?", vbQuestion + vbYesNo + vbDefaultButton2, "Rielabora prestazioni") = vbYes Then
        Unload frmMeseAnno
        Load frmMeseAnno
        frmMeseAnno.letKeyId = keyId
        frmMeseAnno.Letimporto = VirgolaOrPunto(flxGriglia.TextMatrix(vRow, vCol), ".")
        If vCol = 3 Then
            frmMeseAnno.LetimportoScontato = False
        Else
            frmMeseAnno.LetimportoScontato = True
        End If
        frmMeseAnno.Show 1
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

Private Sub cmdElimina_Click()
    Dim blnEliminato As Boolean
    Dim blnElimina As Boolean
    Dim intKey As Integer
    Dim strNome As String
    Dim rsDataset As Recordset
   
    With flxGriglia
        If .Row = 0 Then
            MsgBox "Selezionare il dato da eliminare", vbCritical, "Attenzione"
        Else
            intKey = .TextMatrix(vRow, 0)
            strNome = .TextMatrix(vRow, 1)

            blnElimina = False
            Select Case tTabelle
                Case tpNOMENCLATORE
                Case tpRegioni
                Case tpRENI
                    blnElimina = IsPossibleDelete("TURNI", "CODICE_RENE", intKey)
                    If blnElimina Then
                        blnElimina = IsPossibleDelete("STORICO_DIALISI_GIORNALIERA", "CODICE_RENE", intKey)
                    End If
                    strNome = .TextMatrix(vRow, 3)
                Case tpTIPOLOGIEMEDICO
                Case tpasl
                Case tpCOMUNI
                Case tpDISTRETTI
                Case tpESAME
                Case tpESENZIONI
                Case tpEDTA
                    blnElimina = IsPossibleDelete("ANAMNESI_NEFROLOGICHE", "CODICE_EDTA", intKey)
                    strNome = .TextMatrix(vRow, 2)
            End Select
                
            If blnElimina Then
                If MsgBox("Sicuro di voler eliminare " & strNome & "?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                    Set rsDataset = New Recordset
                    rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & intKey, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                    If rsDataset.EOF And rsDataset.BOF Then
                        MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
                    Else
                        rsDataset.Delete
                        blnEliminato = True
                    End If
                    Set rsDataset = Nothing
                End If
            Else
                MsgBox "Impossibile eliminare " & strNome & " perchè in relazione con altri dati del sistema", vbInformation, Me.Caption
            End If
            
            If blnEliminato Then
                ' rimuove dalla flx
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem (.Row)
                End If
                vRow = 0
                .Row = 0
                MsgBox "Eliminazione avvenuta con successo", vbInformation, Me.Caption
            End If
        End If
    End With
End Sub

Private Function EsisteValore() As Boolean
    Dim i As Integer
    
    Select Case tTabelle
        Case tpasl, tpDISTRETTI
            If tInput.v_valori(3) <> "" Then
                For i = 1 To flxGriglia.Rows - 1
                    If UCase(flxGriglia.TextMatrix(i, 1)) = UCase(tInput.v_valori(1)) And UCase(flxGriglia.TextMatrix(i, 3)) = UCase(cboAppo.List(GetCboListIndex(CInt(tInput.v_valori(3)), cboAppo))) Then
                        EsisteValore = i
                        Exit Function
                    End If
                Next i
            End If
            EsisteValore = 0
        Case tpRegioni
            If tInput.v_valori(2) <> "" Then
                For i = 1 To flxGriglia.Rows - 1
                    If flxGriglia.TextMatrix(i, 1) = tInput.v_valori(2) And flxGriglia.TextMatrix(i, 2) = tInput.v_valori(1) Then
                        EsisteValore = i
                        Exit Function
                    End If
                Next i
            End If
            EsisteValore = 0
        Case tpEDTA
            If tInput.v_valori(2) <> "" Then
                For i = 1 To flxGriglia.Rows - 1
                    If flxGriglia.TextMatrix(i, 1) = tInput.v_valori(2) Then ' And flxGriglia.TextMatrix(i, 2) = tInput.v_valori(1) Then
                        EsisteValore = i
                        Exit Function
                    End If
                Next i
            End If
            EsisteValore = 0
        Case Else
            EsisteValore = Esiste(flxGriglia, 1, 0, tInput.v_valori(1))
    End Select
End Function

Private Sub cmdInserisci_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim num As Integer
    Dim primo As Boolean
    
    If tTabelle = tpESAME And lstOrgani.ListIndex = -1 Then
        MsgBox "Selezionare l'organo a cui associare il nuovo esame", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    Select Case tTabelle
        Case tpRegioni, tpTIPOLOGIEMEDICO, tpEDTA
            tInput.Tipo = tpICOMPOSTO
        Case tpESAME
            tInput.Tipo = tpISINGOLO
        Case tpDISTRETTI
            tInput.Tipo = tpIDISTRETTI
        Case tpNOMENCLATORE
            tInput.Tipo = tpINOMENCLATORE
        Case tpCOMUNI
            tInput.Tipo = tpICOMUNI
        Case tpasl
            tInput.Tipo = tpIASL
        Case tpESENZIONI
            tInput.Tipo = tpIESENZIONE
        Case tpRENI
            tInput.Tipo = tpIRENI
    End Select
    
    primo = True
    tInput.mantieniDati = False
    Do
        If Not primo Then
            If tTabelle = tpRENI Then
                If MsgBox("Postazione già presente." & vbCrLf & "Vuoi duplicarla?", vbQuestion + vbYesNo + vbDefaultButton2, "Inserisci rene") = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                End If
            Else
                MsgBox "Il valore inserito è già presente", vbCritical, "ATTENZIONE!!!"
                tInput.mantieniDati = True
            End If
        End If
        Unload frmInput
        frmInput.Show 1
        primo = False
    Loop While EsisteValore

    If Not (tInput.v_valori(1) = "" And tInput.v_valori(2) = "") Then
        num = GetNumero(nomeTabella)
        Select Case tTabelle
            Case tpESAME
                v_Nomi = Array("KEY", "NOME", "CODICE_ORGANO")
                v_Val = Array(num, tInput.v_valori(1), keyOrgano)
            Case tpRegioni, tpTIPOLOGIEMEDICO, tpEDTA
                v_Nomi = Array("KEY", "CODICE", "NOME")
                v_Val = Array(num, tInput.v_valori(2), tInput.v_valori(1))
            Case tpESENZIONI
                v_Nomi = Array("KEY", "CODICE", "ESENZIONE_QUOTA")
                v_Val = Array(num, tInput.v_valori(1), CBool(tInput.v_valori(2)))
            Case tpDISTRETTI
                v_Nomi = Array("KEY", "CODICE", "NOME", "CODICE_ASL")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3))
            Case tpCOMUNI
                v_Nomi = Array("KEY", "CODICE", "NOME", "REGIONIID")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3))
            Case tpasl
                v_Nomi = Array("KEY", "CODICE", "NOME", "CODICE_REGIONE")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3))
            Case tpNOMENCLATORE
                v_Nomi = Array("KEY", "CODICE", "NOME", "IMPORTO", "IMPORTO_SCONTATO")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(4))
            Case tpRENI
                v_Nomi = Array("KEY", "POSTAZIONE", "MODELLO", "MATRICOLA", "TIPO", "DATA_ROTTAMAZIONE", "SOSTITUITO", "NUMERO_APPARATO")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(4), IIf(tInput.v_valori(5) = "", Null, tInput.v_valori(5)), False, IIf(tInput.v_valori(6) = "", Null, tInput.v_valori(6)))
        End Select
        
        Set rsTabelle = New Recordset
        rsTabelle.Open nomeTabella, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
        rsTabelle.AddNew v_Nomi, v_Val
        rsTabelle.Update
        Set rsTabelle = Nothing
        
        ' aggiorna la flx
        flxGriglia.Rows = 1
        If tTabelle = tpESAME Then
            lstOrgani_Click
        Else
            Call CaricaFlx
        End If
        
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
        vRow = flxGriglia.Row
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        If flxGriglia.Row > 8 Then
            flxGriglia.TopRow = flxGriglia.Row
        End If
        
 '       MsgBox "Inserimento valore effettuato", vbInformation, "Inserimento"
    End If
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

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If flxGriglia.Rows = 1 Then Exit Sub
    If flxGriglia.Row = flxGriglia.Rows - 1 Then
        i = 1
    Else
        i = flxGriglia.Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia.TextMatrix(i, 1), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia.Row = i
            If i >= 8 Or flxGriglia.TopRow > 8 Then
                flxGriglia.TopRow = i
            End If
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
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
        .SetFocus
        If .Col = 3 And Not tTabelle = tpNOMENCLATORE And Not tTabelle = tpRENI Then
            Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            cboAppo.Left = .colPos(.Col) + .Left + 130
            cboAppo.Width = .ColWidth(.Col) + 30
            cboAppo.Top = .rowPos(.Row) + .Top + 45
            cboAppo.ListIndex = GetIndex(cboAppo, .TextMatrix(.Row, .Col))
            cboAppo.Visible = True
            cboAppo.SetFocus
            Call SalvaModifiche
        ElseIf .Col = 2 And tTabelle = tpESENZIONI Then
            Select Case .TextMatrix(.Row, 2)
                Case Is = ESENTE
                    .TextMatrix(.Row, 2) = NONESENTE
                Case Is = NONESENTE
                    .TextMatrix(.Row, 2) = ESENTE
            End Select
            Call SalvaModifiche
        ElseIf .Col = 5 And tTabelle = tpRENI Then
            Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            ' TIPO puo essere neg o hcv o hbv
            If .TextMatrix(.Row, 5) = "NEG" Then
                .TextMatrix(.Row, 5) = "HCV POS"
            ElseIf .TextMatrix(.Row, 5) = "HCV POS" Then
                .TextMatrix(.Row, 5) = "HBV POS"
            Else
                .TextMatrix(.Row, 5) = "NEG"
            End If
            Call SalvaModifiche
        ElseIf .Col = 6 And tTabelle = tpRENI Then
 '           frmCalendario.Show 1
 '            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
 '            cmdAnnulla.Enabled = True
 '            .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
 '           Call SalvaModifiche
            ' cambia colonna per evitave di ricaricare il calendario
            .Col = 0
        Else
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.MaxLength = Len(.TextMatrix(.Row, .Col))
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        End If
    End With
End Sub

Private Sub lstOrgani_Click()
    ' carica la lista degli esami per quell'organo
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    Set rsTabelle = New Recordset
    rsTabelle.Open "ORGANI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    rsTabelle.Filter = ("NOME='" & Apostrophe(lstOrgani.List(lstOrgani.ListIndex)) & "'")
    keyOrgano = rsTabelle("KEY")
    ' devo chiudere tutto perche dopo il filter non mi funziona piu il recordset
    Set rsTabelle = Nothing
    Set rsTabelle = New Recordset
    rsTabelle.Open "SELECT * FROM ESAMI WHERE CODICE_ORGANO=" & keyOrgano & " ORDER BY NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsTabelle.EOF
        With flxGriglia
            .Rows = .Rows + 1
            ' lo aggiunge alla lista
            .TextMatrix(.Rows - 1, 0) = rsTabelle("KEY")
            .TextMatrix(.Rows - 1, 1) = rsTabelle("NOME")
            rsTabelle.MoveNext
        End With
    Loop
    Set rsTabelle = Nothing
End Sub

Private Sub cboAppo_Click()
    cboAppo.Visible = False
End Sub

Private Sub cboAppo_LostFocus()
    If flxGriglia.TextMatrix(vRow, vCol) <> cboAppo.Text Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(flxGriglia.Row, flxGriglia.Col) = cboAppo.Text
        Call SalvaModifiche
    End If
    cboAppo.Visible = False
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
    Select Case tTabelle
        Case tpESAME
            txtAppo.MaxLength = 50
        Case tpESENZIONI
            txtAppo.MaxLength = 10
        Case tpasl
            txtAppo.MaxLength = IIf(vCol = 1, 3, 25)
        Case tpCOMUNI
            txtAppo.MaxLength = IIf(vCol = 1, 6, 40)
        Case tpRegioni
            txtAppo.MaxLength = IIf(vCol = 1, 3, 30)
        Case tpTIPOLOGIEMEDICO
            txtAppo.MaxLength = IIf(vCol = 1, 1, 50)
        Case tpDISTRETTI
            txtAppo.MaxLength = IIf(vCol = 1, 5, 4)
        Case tpNOMENCLATORE
            txtAppo.MaxLength = Choose(vCol, 10, 100, 6, 6)
        Case tpRENI
            txtAppo.MaxLength = Choose(vCol, 3, 3, 50, 50)
        Case tpEDTA
            txtAppo.MaxLength = IIf(vCol = 2, 150, 3)
    End Select
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_Change()
    If (tTabelle = tpNOMENCLATORE And (vCol = 3 Or vCol = 4)) Or (tTabelle = tpRENI And (vCol = 2)) Then
        If Not (lettera = "." Or lettera = "") Then
            Call OnlyNumber(txtAppo, lettera)
        End If
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim PostazionePrecedente As String
    
    txtAppo.Visible = False
    If (flxGriglia.TextMatrix(vRow, vCol)) <> (txtAppo) Then
        If txtAppo = "" Then
            MsgBox "Impossibile memorizzare dati vuoti", vbCritical, "Attenzione"
            flxGriglia.Row = vRow
            flxGriglia.Col = vCol
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Sub
        End If
        
        If vCol = 1 Then
        PostazionePrecedente = flxGriglia.TextMatrix(vRow, vCol)
            If Esiste(flxGriglia, 1, vRow, txtAppo) Then
                flxGriglia.TextMatrix(vRow, vCol) = UCase((txtAppo.Text))
                If MsgBox("Valore già presente." & vbCrLf & "Vuoi duplicarlo?", vbQuestion + vbYesNo + vbDefaultButton2, "Inserimento Valori") = vbYes Then
                    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
                    cmdAnnulla.Enabled = True
                    Call SalvaModifiche
                Else
                    flxGriglia.TextMatrix(vRow, vCol) = PostazionePrecedente
                End If
            Else
                If tTabelle = tpCOMUNI And vCol = 1 Then
                    If Not Len(txtAppo) = 6 Then
                        MsgBox "Il codice ISTAT deve essere di 6 cifre", vbCritical, "ATTENZIONE!!!"
                        Exit Sub
                    End If
                End If
                Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
                cmdAnnulla.Enabled = True
                flxGriglia.TextMatrix(vRow, vCol) = UCase((txtAppo.Text))
                Call SalvaModifiche
            End If
        ElseIf tTabelle = tpRENI And vCol = 2 Then
            If Esiste(flxGriglia, 2, vRow, txtAppo) Then
                MsgBox "Il valore inserito è già presente", vbCritical, "Attenzione"
            Else
                Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
                cmdAnnulla.Enabled = True
                flxGriglia.TextMatrix(vRow, vCol) = UCase((txtAppo.Text))
                Call SalvaModifiche
            End If
        Else
            If (tTabelle = tpNOMENCLATORE And (vCol = 3 Or vCol = 4)) Or (tTabelle = tpRENI And (vCol = 1 Or vCol = 2)) Then
                If ControlloNumerico(txtAppo.Text) Then
                    txtAppo.Visible = False
                    Exit Sub
                End If
            End If
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            flxGriglia.TextMatrix(vRow, vCol) = (txtAppo.Text)
            Call SalvaModifiche
        End If
    End If
End Sub

Private Sub WheelCatcher1_WheelRotation(Rotation As Long, X As Long, Y As Long, CtrlHwnd As Long)
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
End Sub
