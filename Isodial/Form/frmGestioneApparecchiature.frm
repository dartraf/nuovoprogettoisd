VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGestioneApparecchiature 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   15000
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
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.ComboBox cboModalitaAcquisizioneProprieta 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "cboModalitaAcqusizioneProprieta"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3615
      End
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   3615
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
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   7200
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   6375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11245
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
         MouseIcon       =   "frmGestioneApparecchiature.frx":0000
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
      TabIndex        =   4
      Top             =   6600
      Width           =   10455
      Begin VB.CommandButton Command1 
         Caption         =   "&Annulla digitazione"
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
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1815
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
         Left            =   9120
         TabIndex        =   8
         Top             =   240
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
         Height          =   600
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
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
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGestioneApparecchiature"
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

Private Sub cboModalitaAcquisizioneProprieta_Click(Index As Integer)
    cboModalitaAcquisizioneProprieta(0).Visible = False
End Sub

Private Sub cboModalitaAcquisizioneProprieta_LostFocus(Index As Integer)
    
    If Len(cboModalitaAcquisizioneProprieta(0)) > 20 Then
        MsgBox "Impossibile memorizzare più di 20 caratteri", vbInformation, "Informazione"
        cboModalitaAcquisizioneProprieta(0).Text = ""
        Exit Sub
    End If
    
    If cboModalitaAcquisizioneProprieta(0).Text <> "" Then
        Call GestisciNuovo("MODALITA_ACQUISIZIONE", cboModalitaAcquisizioneProprieta(0))
    End If
    
    If flxGriglia.TextMatrix(vRow, vCol) <> cboModalitaAcquisizioneProprieta(0).Text Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(vRow, vCol) = cboModalitaAcquisizioneProprieta(0).Text
        Call SalvaModifiche
    End If
    cboModalitaAcquisizioneProprieta(0).Visible = False
    
    'Carico il valore nella flxgrid così quando inserisco un nuovo
    'valore nella combobox viene inserito direttamente nella flxgrid
    'flxGriglia.TextMatrix(vRow, 5) = cboModalitaAcquisizioneProprieta(0).Text

End Sub

Private Sub cmdAnnulla12_Click()
    frmGestioniApparecchiatureInput.Show 1
End Sub

Private Sub Command1_Click()
    frmGestioniApparecchiatureInput.Show 1
End Sub

Private Sub Form_Activate()
    Call RicaricaComboBox("MODALITA_ACQUISIZIONE", "NOME", cboModalitaAcquisizioneProprieta(0))
End Sub

Private Sub Form_Load()
    Dim i As Integer
       
    Set objAnnulla = New CAnnulla
    flxGriglia.Rows = 1
    
    Select Case tTabelle
            Case tpRENI
                nomeTabella = "RENI"
                frmGestioneApparecchiature.Caption = "Gestione Apparecchiature: RENI ARTIFICIALI"
                With flxGriglia
                    .Cols = 9
                    .ColWidth(1) = .ColWidth(1) * 0.3
                    .ColWidth(2) = .ColWidth(1) * 1.2   '1 / 2.8
                    .ColWidth(3) = .ColWidth(1) * 0.8
                    .ColWidth(4) = .ColWidth(1) * 1.8
                    .ColWidth(5) = .ColWidth(1) * 1#
                    .ColWidth(6) = .ColWidth(1) * 0.7
                    .ColWidth(7) = .ColWidth(1) * 1.5
                    .ColWidth(8) = .ColWidth(1) * 1.5
                    
                    '.ColWidth(5) = .ColWidth(1) * 0.9 ' ultimi campi modificare alla fine
                    '.ColWidth(6) = .ColWidth(1) * 0.8
                    '.ColWidth(7) = .ColWidth(1) * 0.9
                    
                    .TextMatrix(0, 1) = "N° Prog."
                    .TextMatrix(0, 2) = "Postazione"
                    .TextMatrix(0, 3) = "N° rene"
                    .TextMatrix(0, 4) = "Descrizione"
                    .TextMatrix(0, 5) = "Azienda"
                    .TextMatrix(0, 6) = "Mod.Acquisiz.Prop."
                    .TextMatrix(0, 7) = "Period.Ammortam."
                    .TextMatrix(0, 8) = "Dt.Installaz."
                    '.TextMatrix(0, 5) = "Matricola"  'ultimi campi modificare alla fine
                    '.TextMatrix(0, 6) = "Tipo"
                    '.TextMatrix(0, 7) = "Dt.Rottam."
                End With
                Call CaricaFlx
        End Select
    
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
    Dim Numero_Progressivo As Integer
    
    flxGriglia.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    
    strSql = "SELECT * FROM " & nomeTabella
    
    If tTabelle = tpRENI Then
        strSql = strSql & " ORDER BY SOSTITUITO DESC, DATA_ROTTAMAZIONE DESC, POSTAZIONE"
    End If
    
    Set rsTabelle = New Recordset
    rsTabelle.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ' carica la lista
    If Not (rsTabelle.EOF And rsTabelle.BOF) Then
        Do While Not rsTabelle.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsTabelle("KEY")
                
                If tTabelle = tpRENI Then
                    Numero_Progressivo = Numero_Progressivo + 1
                    .TextMatrix(.Rows - 1, 1) = Numero_Progressivo
                    .TextMatrix(.Rows - 1, 2) = rsTabelle("POSTAZIONE")
                    .TextMatrix(.Rows - 1, 3) = rsTabelle("NUMERO_RENE") & ""
                    .TextMatrix(.Rows - 1, 4) = rsTabelle("TIPO_RENE") & ""
                    .TextMatrix(.Rows - 1, 5) = rsTabelle("AZIENDA") & ""
                    .TextMatrix(.Rows - 1, 6) = rsTabelle("MODALITA_ACQUISIZIONE_PROPRIETA") & "" ' come caricare modalita acquisizione
                    .TextMatrix(.Rows - 1, 7) = rsTabelle("PERIODO_AMMORTAMENTO") & ""
                    .TextMatrix(.Rows - 1, 8) = rsTabelle("DATA_INSTALLAZIONE") & ""
                    
                    '.TextMatrix(.Rows - 1, 5) = rsTabelle("MATRICOLA") & ""
                    'If rsTabelle("TIPO") = 0 Then ' vedere alla fine il valore
                    '    .TextMatrix(.Rows - 1, 6) = "NEG"
                    'ElseIf rsTabelle("TIPO") = 1 Then
                    '    .TextMatrix(.Rows - 1, 6) = "HCV POS"
                    'Else
                    '    .TextMatrix(.Rows - 1, 6) = "HBV POS"
                    'End If
                    '.Col = 6       'vedere alla fine il valore
                    '.Row = .Rows - 1
                    '.CellAlignment = vbRightJustify
                    '.CellForeColor = vbRed
                    '.TextMatrix(.Rows - 1, 7) = rsTabelle("DATA_ROTTAMAZIONE") & ""
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
    Dim valore As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant

    With flxGriglia
        keyId = .TextMatrix(vRow, 0)
        
        If tTabelle = tpRENI Then ' modificare qui
            v_Nomi = Array("KEY", "POSTAZIONE", "NUMERO_RENE", "TIPO_RENE", "AZIENDA", "MODALITA_ACQUISIZIONE_PROPRIETA", "PERIODO_AMMORTAMENTO", "DATA_INSTALLAZIONE") ' "MATRICOLA", "TIPO", "DATA_ROTTAMAZIONE")
            '   If .TextMatrix(vRow, 6) = "HCV POS" Then
            '    valore = 1
            'ElseIf .TextMatrix(vRow, 6) = "HBV POS" Then
            '    valore = 2
            'Else
            '    valore = 0
            'End If
            v_Val = Array(keyId, .TextMatrix(vRow, 1), .TextMatrix(vRow, 2), .TextMatrix(vRow, 3), .TextMatrix(vRow, 4), .TextMatrix(vRow, 5), .TextMatrix(vRow, 6), IIf(.TextMatrix(vRow, 7) = "", Null, .TextMatrix(vRow, 7))) ' valori finali .TextMatrix(vRow, 5), valore, IIf(.TextMatrix(vRow, 7) = "", Null, .TextMatrix(vRow, 7)))
        End If
        
        Set rsTabelle = New Recordset
        
        rsTabelle.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsTabelle.Update v_Nomi, v_Val
        
        Set rsTabelle = Nothing
        
    End With
End Sub

Private Sub cmdAnnulla_Click()
    frmGestioniApparecchiatureInput.Show 1
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
                Case tpREGIONI
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
        Case tpREGIONI, tpCOMUNI, tpTIPOLOGIEMEDICO, tpEDTA
            tInput.Tipo = tpICOMPOSTO
        Case tpESAME
            tInput.Tipo = tpISINGOLO
        Case tpDISTRETTI
            tInput.Tipo = tpIDISTRETTI
        Case tpNOMENCLATORE
            tInput.Tipo = tpINOMENCLATORE
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
                MsgBox "Il valore inserito è già presente", vbCritical, "Attenzione"
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
            Case tpCOMUNI, tpREGIONI, tpTIPOLOGIEMEDICO, tpEDTA
                v_Nomi = Array("KEY", "CODICE", "NOME")
                v_Val = Array(num, tInput.v_valori(2), tInput.v_valori(1))
            Case tpESENZIONI
                v_Nomi = Array("KEY", "CODICE", "ESENZIONE_QUOTA")
                v_Val = Array(num, tInput.v_valori(1), CBool(tInput.v_valori(2)))
            Case tpDISTRETTI
                v_Nomi = Array("KEY", "CODICE", "NOME", "CODICE_ASL")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3))
            Case tpasl
                v_Nomi = Array("KEY", "CODICE", "NOME", "CODICE_REGIONE")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3))
            Case tpNOMENCLATORE
                v_Nomi = Array("KEY", "CODICE", "NOME", "IMPORTO", "IMPORTO_SCONTATO")
                v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(4))
            Case tpRENI
                v_Nomi = Array("KEY", "POSTAZIONE", "TIPO_RENE", "MATRICOLA", "TIPO", "DATA_ROTTAMAZIONE", "SOSTITUITO", "NUMERO_RENE")
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
                            
        ElseIf .Col = 5 And tTabelle = tpRENI Then  'modficare ì valori spostare la combbox
            Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            cboModalitaAcquisizioneProprieta(0).Left = .colPos(.Col) + .Left + 130
            cboModalitaAcquisizioneProprieta(0).Width = .ColWidth(.Col) + 30
            cboModalitaAcquisizioneProprieta(0).Top = .rowPos(.Row) + .Top + 45
            cboModalitaAcquisizioneProprieta(0).ListIndex = GetIndex(cboModalitaAcquisizioneProprieta(0), .TextMatrix(.Row, .Col))
            cboModalitaAcquisizioneProprieta(0).Visible = True
            cboModalitaAcquisizioneProprieta(0).SetFocus
            Call SalvaModifiche
            
        ElseIf .Col = 7 And tTabelle = tpRENI Then
            frmCalendario.Show 1
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
            cmdAnnulla.Enabled = True
            .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
            Call SalvaModifiche
           ' cambia colonna per evitave di ricaricare il calendario
            .Col = 0
            
        'ElseIf .Col = 6 And tTabelle = tpRENI Then 'modificare qui
        '    Call objAnnulla.Add(.TextMatrix(vRow, vCol), vCol, Int(.TextMatrix(vRow, 0)))
        '    cmdAnnulla.Enabled = True
            ' TIPO puo essere neg o hcv o hbv
        '    If .TextMatrix(.Row, 6) = "NEG" Then
        '        .TextMatrix(.Row, 6) = "HCV POS"
        '    ElseIf .TextMatrix(.Row, 6) = "HCV POS" Then
        '        .TextMatrix(.Row, 6) = "HBV POS"
        '    Else
        '        .TextMatrix(.Row, 6) = "NEG"
        '    End If
        '    Call SalvaModifiche
        
        'ElseIf .Col = 7 And tTabelle = tpRENI Then 'modificare qui
        '    frmCalendario.Show 1
        '    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
        '    cmdAnnulla.Enabled = True
        '    .TextMatrix(.Row, .Col) = IIf(laData <> "", laData, .TextMatrix(.Row, .Col))
        '    Call SalvaModifiche
            ' cambia colonna per evitave di ricaricare il calendario
        '    .Col = 0
        
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
        
        Case tpRENI
            txtAppo.MaxLength = Choose(vCol, 3, 3, 50, 50, 20, 2, 10) ' 50) è il monitor
            'Postazione, N° Rene, Descrizione, Azienda, Mod.Acquisizione Proprietà, Periodo Ammortamento, Data Installazione, Matricola
                                         
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
    ' vedere questa sub per decidere quali valori memorizzare
    If tTabelle = tpRENI And vCol = 2 Then
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
                If MsgBox("Postazione già presente." & vbCrLf & "Vuoi duplicarla?", vbQuestion + vbYesNo + vbDefaultButton2, "Inserisci rene") = vbYes Then
                    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, Int(flxGriglia.TextMatrix(vRow, 0)))
                    cmdAnnulla.Enabled = True
                    Call SalvaModifiche
                Else
                    flxGriglia.TextMatrix(vRow, vCol) = PostazionePrecedente
                End If
            Else
                If tTabelle = tpCOMUNI And vCol = 1 Then
                    If Not Len(txtAppo) = 6 Then
                        MsgBox "Il codice ISTAT deve essere di 6 caratteri", vbCritical, "Attenzione"
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


