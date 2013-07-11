VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Apparati"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Manutenzione Apparato"
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
      TabIndex        =   5
      Top             =   4440
      Width           =   14775
      Begin MSFlexGridLib.MSFlexGrid flxManutenzione 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5741
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
         MouseIcon       =   "frmApparati.frx":0000
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
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   2400
         TabIndex        =   4
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
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5741
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
         MouseIcon       =   "frmApparati.frx":015A
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   14775
      Begin VB.CommandButton cmdEliminaApparato 
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
         Left            =   11880
         TabIndex        =   11
         Top             =   170
         Width           =   1215
      End
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Inserisci Apparato"
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
         Left            =   13320
         TabIndex        =   3
         Top             =   170
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   7920
      Width           =   14775
      Begin VB.CommandButton cmdEliminaManutenzioneApparato 
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
         Left            =   7680
         TabIndex        =   12
         Top             =   170
         Width           =   1215
      End
      Begin VB.CommandButton cmdManutenzioneOrdinaria 
         Caption         =   "Inserisci Manut. &Ordinaria"
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
         TabIndex        =   10
         Top             =   170
         Width           =   1935
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
         Left            =   13440
         TabIndex        =   9
         Top             =   170
         Width           =   1215
      End
      Begin VB.CommandButton cmdManutenzioneStraordinaria 
         Caption         =   "Inserisci Manut. &Straordinaria"
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
         Left            =   11280
         TabIndex        =   8
         Top             =   170
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmApparati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsApparati As Recordset
Dim rsManutenziona As Recordset
Dim rsEliminaManutenzione As Recordset
Dim vRow As Integer             ' riga selezionata
Dim vCol As Integer             ' colonna selezionata
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx

Private Sub cmdEliminaApparato_Click()

    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare l' Apparato da eliminare", vbInformation, "Informazione"
    
    ElseIf flxManutenzione.Rows > 1 Then
        MsgBox "Impossibile eliminare un' Apparato" & vbCrLf & "in presenza di una scheda di manutenzione", vbInformation, "Informazione"
    
    ElseIf MsgBox("Sei sicuro di voler eliminare l' Apparato selezionato?", vbInformation + vbYesNo + vbDefaultButton2, "Informazione") = vbYes Then
        Call EliminaApparato
        Call CaricaFlx
        
    End If
    
End Sub

Private Sub EliminaApparato()

    Dim cmCommand As New Command
    
    'Elimina l' Apparato
    cmCommand.CommandType = adCmdText
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandText = "DELETE * FROM APPARATI WHERE KEY=" & KeyApparato
    cmCommand.Execute

End Sub

Private Sub cmdEliminaManutenzioneApparato_Click()
    
    If flxManutenzione.Row = 0 Then
        MsgBox "Selezionare la Manutenzione dell' Apparato da eliminare", vbInformation, "Informazione"
        Exit Sub
    End If
    
    Set rsEliminaManutenzione = New Recordset
        
        rsEliminaManutenzione.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE KEY= " & KeyReturnManutenzione, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If IsNull(rsEliminaManutenzione("DATA_EFFETTIVA_MANUTENZIONE")) = False And IsNull(rsEliminaManutenzione("NUMERO_DOCUMENTO")) = False And IsNull(rsEliminaManutenzione("DETTAGLI_INTERVENTO")) = False Then
            MsgBox "Impossibile eliminare la Manutenzione dell' Apparato in presenza dei seguenti campi:" & vbCrLf & "- Data Effettiva Manutenzione" & vbCrLf & "- N° Riferimento Documento" & vbCrLf & "- Dettagli Intervento", vbInformation, "Informazione"
            Exit Sub
        End If
        
        If IsNull(rsEliminaManutenzione("DATA_EFFETTIVA_MANUTENZIONE")) = True And IsNull(rsEliminaManutenzione("NUMERO_DOCUMENTO")) = True And IsNull(rsEliminaManutenzione("DETTAGLI_INTERVENTO")) = True Then
            ElseIf MsgBox("Sei sicuro di voler eliminare la Manutenzione selezionato?", vbInformation + vbYesNo + vbDefaultButton2, "Informazione") = vbYes Then
            Call EliminaManutenzione
            Call CaricaFlxManutenzione
        End If
    
    Set rsEliminaManutenzione = Nothing
           
End Sub

Private Sub EliminaManutenzione()
    
    Dim cmCommand As New Command
    
    'Elimina la Manutenzione
    cmCommand.CommandType = adCmdText
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandText = "DELETE * FROM MANUTENZIONE_APPARATI WHERE KEY=" & KeyReturnManutenzione
    cmCommand.Execute

End Sub
Private Sub cmdManutenzioneOrdinaria_Click()
    Dim num As Integer
    
    If KeyApparato = 0 Then
        MsgBox "Selezionare Un Apparato", vbInformation, "INFORMAZIONE"
    Else
        tTabellaManutenzione = tpMANUTENZIONEORDINARIA
        frmInserisciManutenzione.Show 1
        Call CaricaFlxManutenzione
    End If
    
    ' Funzione per Colorare il record
    If KeyReturnManutenzione = 0 Or KeyReturnManutenzione = -1 Then
        num = GetNumero("MANUTENZIONE_APPARATI") - 1
    Else
        num = KeyReturnManutenzione
    End If
    ' si posiziona sul record e lo seleziona
    flxManutenzione.Row = Esiste(flxManutenzione, 0, vRow, num)
    vRow = flxManutenzione.Row
    Call ColoraFlx(flxManutenzione, flxManutenzione.Cols - 1)
    If flxManutenzione.Row > 10 Then
        flxManutenzione.TopRow = flxManutenzione.Row
    End If
    ' Per evitare di Ricaricare lo stesso dato
    KeyReturnManutenzione = 0

End Sub

Private Sub cmdManutenzioneStraordinaria_Click()
    Dim num As Integer
    
    If KeyApparato = 0 Then
        MsgBox "Selezionare Un Apparato", vbInformation, "INFORMAZIONE"
    Else
        tTabellaManutenzione = tpMANUNTENZIONESTRAORDINARIA
        frmInserisciManutenzione.Show 1
        Call CaricaFlxManutenzione
    End If
    
    
    ' Funzione per Colorare il record
    If KeyReturnManutenzione = 0 Or KeyReturnManutenzione = -1 Then
        num = GetNumero("MANUTENZIONE_APPARATI") - 1
    Else
        num = KeyReturnManutenzione
    End If
    ' si posiziona sul record e lo seleziona
    flxManutenzione.Row = Esiste(flxManutenzione, 0, vRow, num)
    vRow = flxManutenzione.Row
    Call ColoraFlx(flxManutenzione, flxManutenzione.Cols - 1)
    If flxManutenzione.Row > 10 Then
        flxManutenzione.TopRow = flxManutenzione.Row
    End If
    ' Per evitare di Ricaricare lo stesso dato
    KeyReturnManutenzione = 0
    
End Sub

Private Sub flxManutenzione_Click()
    flxManutenzione.SetFocus
    If VerificaClickFlx(flxManutenzione) = False Then
        ' discolora
        Call ColoraFlx(flxManutenzione, flxManutenzione.Cols - 1, True)
        ' annulla le row e col
        flxManutenzione.Row = 0
        flxManutenzione.Col = 0
    Else
        vRow = flxManutenzione.Row
        vCol = flxManutenzione.Col
        Call ColoraFlx(flxManutenzione, flxManutenzione.Cols - 1)
        
        KeyReturnManutenzione = flxManutenzione.TextMatrix(vRow, 0)
    End If
End Sub

Private Sub flxManutenzione_DblClick()
    If VerificaClickFlx(flxManutenzione) = False Then Exit Sub
    
    ' Seleziono la key della manutenzione e la passo
    KeyReturnManutenzione = flxManutenzione.TextMatrix(vRow, 0)
    
    ' Se il campo della griglia Tipo Manutenzione è uguale a Straordinaria
    ' carica il cmdManutenzioneStraordinaria
    If flxManutenzione.TextMatrix(vRow, 1) = "STRAORDINARIA" Then
        cmdManutenzioneStraordinaria_Click
    ' Se il campo della griglia Tipo Manutenzione è diverso (per evitare di elencare tutti i campi)
    ' da Straordinaria carica il cmdManutenzioneOrdianria
    ElseIf flxManutenzione.TextMatrix(vRow, 1) <> "STRAORDINARIA" Then
        cmdManutenzioneOrdinaria_Click
    End If
    
    'per evitare di ricaricare l'apparato
    KeyReturnManutenzione = 0

End Sub

Private Sub Form_Load()
Dim i As Integer
       
    ' Griglia Apparato
    Set objAnnulla = New CAnnulla
    flxGriglia.Rows = 1
    
    With flxGriglia
        .Cols = 9
        .ColWidth(1) = .ColWidth(1) * 0.2
        .ColWidth(2) = .ColWidth(1) * 2.1
        .ColWidth(3) = .ColWidth(1) * 3.2
        .ColWidth(4) = .ColWidth(1) * 2.5
        .ColWidth(5) = .ColWidth(1) * 2.2
        .ColWidth(6) = .ColWidth(1) * 2.8
        .ColWidth(7) = .ColWidth(1) * 2.2
        .ColWidth(8) = .ColWidth(1) * 2.1
                                       
        .TextMatrix(0, 1) = "N° Inv."
        .TextMatrix(0, 2) = "N° Apparato"
        .TextMatrix(0, 3) = "Tipo Apparato"
        .TextMatrix(0, 4) = "Modello"
        .TextMatrix(0, 5) = "Matricola"
        .TextMatrix(0, 6) = "Produttore"
        .TextMatrix(0, 7) = "PROXREVFUN"
        .TextMatrix(0, 8) = "PROXREVSIC"
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
     
     
    ' Griglia Manutenzione
    Set objAnnulla = New CAnnulla
    flxManutenzione.Rows = 1
    
    With flxManutenzione
        .Cols = 8
        .ColWidth(1) = .ColWidth(1) * 0.4
        .ColWidth(2) = .ColWidth(1) * 1.4
        .ColWidth(3) = .ColWidth(1) * 1.4
        .ColWidth(4) = .ColWidth(1) * 1.4
        .ColWidth(5) = .ColWidth(1) * 1.5
        .ColWidth(6) = .ColWidth(1) * 1.5
        .ColWidth(7) = .ColWidth(1) * 0.9
                                       
        .TextMatrix(0, 1) = "Tp. Manutenz."
        .TextMatrix(0, 2) = "Dt. Scad. Manutenz."
        .TextMatrix(0, 3) = "Dt. Rich. Manutenz."
        .TextMatrix(0, 4) = "Dt. Effet. Manutenz."
        .TextMatrix(0, 5) = "Descr. Manutenz."
        .TextMatrix(0, 6) = "Dett. Intervento"
        .TextMatrix(0, 7) = "N° Doc. Lav."
    End With
        
    With flxManutenzione
        .ColWidth(0) = 0
         .Row = 0
         .MousePointer = flexCustom
         For i = 0 To flxManutenzione.Cols - 1
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
    
    Set rsApparati = New Recordset
    rsApparati.Open "SELECT * FROM APPARATI ORDER BY NUMERO_INVENTARIO ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not (rsApparati.EOF And rsApparati.BOF) Then
        Do While Not rsApparati.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsApparati("KEY")
                .TextMatrix(.Rows - 1, 1) = rsApparati("NUMERO_INVENTARIO")
                .TextMatrix(.Rows - 1, 2) = rsApparati("NUMERO_APPARATO") & ""
                .TextMatrix(.Rows - 1, 3) = rsApparati("TIPO_APPARATO") & ""
                .TextMatrix(.Rows - 1, 4) = rsApparati("MODELLO") & ""
                .TextMatrix(.Rows - 1, 5) = rsApparati("MATRICOLA") & ""
                .TextMatrix(.Rows - 1, 6) = rsApparati("PRODUTTORE") & ""
                .TextMatrix(.Rows - 1, 7) = rsApparati("PROXREVFUN") & ""
                .TextMatrix(.Rows - 1, 8) = rsApparati("PROXREVSIC") & ""
                rsApparati.MoveNext
            End With
        Loop
    End If
    Set rsApparati = Nothing
    flxGriglia.Row = 0
End Sub

Private Sub CaricaFlxManutenzione()
    
    flxManutenzione.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    
    Set rsManutenziona = New Recordset
    rsManutenziona.Open "SELECT * FROM MANUTENZIONE_APPARATI WHERE CODICE_APPARATO= " & KeyApparato & " ORDER BY KEY DESC ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not (rsManutenziona.EOF And rsManutenziona.BOF) Then
        Do While Not rsManutenziona.EOF
            With flxManutenzione
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsManutenziona("KEY")
                .TextMatrix(.Rows - 1, 1) = rsManutenziona("TIPO_MANUTENZIONE") & ""
                .TextMatrix(.Rows - 1, 2) = rsManutenziona("DATA_SCADENZA_MANUTENZIONE") & ""
                .TextMatrix(.Rows - 1, 3) = rsManutenziona("DATA_RICHIESTA_MANUTENZIONE") & ""
                .TextMatrix(.Rows - 1, 4) = rsManutenziona("DATA_EFFETTIVA_MANUTENZIONE") & ""
                .TextMatrix(.Rows - 1, 5) = rsManutenziona("DESCRIZIONE_MANUTENZIONE") & ""
                .TextMatrix(.Rows - 1, 6) = rsManutenziona("DETTAGLI_INTERVENTO") & ""
                .TextMatrix(.Rows - 1, 7) = rsManutenziona("NUMERO_DOCUMENTO") & ""
                rsManutenziona.MoveNext
            End With
        Loop
    End If
    
    Set rsManutenziona = Nothing
    flxManutenzione.Row = 0
End Sub

Private Sub cmdChiudi_Click()
    Unload frmApparati
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

Private Sub cmdInserisci_Click()
    Dim num As Integer

    frmApparatiInput.Show 1
    Call CaricaFlx
    
    ' Funzione per Colorare il Record
    If MantieniKeyReturn = 0 Or MantieniKeyReturn = -1 Then
        num = GetNumero("APPARATI") - 1
    Else
        num = MantieniKeyReturn
    End If
    ' si posiziona sul record e lo seleziona
    flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
    vRow = flxGriglia.Row
    Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    If flxGriglia.Row > 10 Then
        flxGriglia.TopRow = flxGriglia.Row
    End If
    ' Per evitare di caricare lo stesso dato
    MantieniKeyReturn = 0
    
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
        flxManutenzione.Rows = 1
    Else
        vRow = flxGriglia.Row
        vCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        
        ' seleziono la key dell' apparato per passarla alla tab Manutenzione
        KeyApparato = flxGriglia.TextMatrix(vRow, 0)
        Call CaricaFlxManutenzione
    End If
End Sub

Private Sub flxGriglia_DblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    
    ' Seleziono la key dell' apparato e la passo
    ' la passo con la variabile, altrimenti da errore
    tTrova.keyReturn = KeyApparato
    MantieniKeyReturn = tTrova.keyReturn
    cmdInserisci_Click
    tTrova.keyReturn = 0    'per evitare di ricaricare l'apparato
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

