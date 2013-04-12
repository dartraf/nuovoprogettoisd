VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGestisciPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Utenti"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
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
         Left            =   240
         MaxLength       =   25
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1920
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmGestisciPassword.frx":0000
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
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   9855
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
         Left            =   5640
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
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1335
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
         Height          =   495
         Left            =   4080
         TabIndex        =   3
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
         Left            =   8400
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGestisciPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmGestisciPassword.frm
'
' <b>Descrizione</b>: Scheda Gestione Utenti associata alla tab LOGIN
'
' @remarks
'
' @author
'
' @date 07/02/2011 18.11
Option Explicit

'' rs della scheda
Dim rsPassword As Recordset
Dim vCol As Integer
Dim vRow As Integer
'' oggetto che gestisce l'annullamento dei dati nelle flx
Dim objAnnulla As CAnnulla

Private Sub Form_Load()
    Dim i As Integer
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 5
            .ColAlignment(i) = vbLeftJustify
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
    ' carica l'oggetto
    Set objAnnulla = New CAnnulla
    Call CaricaFlx
End Sub

'' Salva le modifiche di un singolo campo del record
Private Sub SalvaModifiche()

    Dim nome As Variant
    Dim valore As Variant
    
    Select Case vCol
        Case 1:
            nome = "COGNOME"
        Case 2:
            nome = "NOME"
        Case 3
            nome = "CHIAVE"
        Case 4
            nome = "PASSWORD"
    End Select
    With flxGriglia
        If vCol <> 5 Then
            valore = .TextMatrix(vRow, vCol)
        End If
        Set rsPassword = New Recordset
        rsPassword.Open "SELECT * FROM LOGIN WHERE KEY=" & .TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsPassword.Update nome, valore
        If nome = "PASSWORD" Then
            rsPassword.Update "DATA", date
        End If
        rsPassword.Close
        Set rsPassword = Nothing
    End With
'    If (flxGriglia.TextMatrix(vRow, 5) = "Medico" Or flxGriglia.TextMatrix(vRow, 5) = "Infermiere") And (vCol = 1 Or vCol = 2) Then
'        If IsPresente Then
'            Call ModificaDaOrganigramma
'        End If
'    End If
End Sub

'' Carica la scheda nella flx
Private Sub CaricaFlx()
    ' pulisce la griglia
    flxGriglia.Rows = 1
    vRow = 0
    vCol = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    Set rsPassword = New Recordset
    rsPassword.Open "SELECT * FROM LOGIN WHERE ELIMINATO=FALSE ORDER BY TIPO DESC, COGNOME, NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsPassword.EOF And rsPassword.BOF) Then
        Do While Not rsPassword.EOF
          With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsPassword("KEY")
            .TextMatrix(.Rows - 1, 1) = rsPassword("COGNOME") & ""
            .TextMatrix(.Rows - 1, 2) = rsPassword("NOME") & ""
            .TextMatrix(.Rows - 1, 3) = rsPassword("CHIAVE")
            .TextMatrix(.Rows - 1, 4) = rsPassword("PASSWORD")
            If rsPassword("TIPO") = 1 Then
                .TextMatrix(.Rows - 1, 5) = "Medico"
            ElseIf rsPassword("TIPO") = 2 Then
                 .TextMatrix(.Rows - 1, 5) = "Infermiere"
            ElseIf rsPassword("TIPO") = 3 Then
                .TextMatrix(.Rows - 1, 5) = "Contabile"
            ElseIf rsPassword("TIPO") = 4 Then
                .TextMatrix(.Rows - 1, 5) = "Amministratore"
            End If
            rsPassword.MoveNext
          End With
        Loop
    End If
    Set rsPassword = Nothing
    flxGriglia.Row = 0
End Sub

'' Salva l'eliminazione nel db di tracciature
Private Sub SalvaEliminazione()
    Dim v_nome As Variant
    Dim v_Val As Variant
    Dim rsDataset As New Recordset
    v_nome = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_ELIMINATO")
    v_Val = Array(tAccesso.key, date, Time, flxGriglia.TextMatrix(flxGriglia.Row, 0))
    rsDataset.Open "E_UTENTI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_nome, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

'' Modifica il  medico o infermiere nella relativa tabella (MEDICI_DIALISI, INFERMIERI)
'Private Sub ModificaDaOrganigramma()
'    Dim rsDataset As New Recordset
'    Dim nomeTabella As String
'    Dim nome As String
'    Dim cognome As String
'    Dim num As Integer
    
'    If flxGriglia.TextMatrix(vRow, 5) = "Medico" Then
'        nomeTabella = "MEDICI_DIALISI"
'    Else
'        nomeTabella = "INFERMIERI"
'    End If
'    nome = flxGriglia.TextMatrix(vRow, 2)
'    cognome = flxGriglia.TextMatrix(vRow, 1)
    
'    rsDataset.Open "SELECT * FROM UTENTI_PERSONALE WHERE CODICE_UTENTE=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
'    num = rsDataset("CODICE_PERSONALE")
'    rsDataset.Close
    
'    rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & num, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
'    rsDataset.Update "COGNOME", UCase(cognome)
'    rsDataset.Update "NOME", UCase(nome)
'    Set rsDataset = Nothing
'End Sub

'' Elimina il nuovo medico o infermiere nella relativa tabella (MEDICI_DIALISI, INFERMIERI)
' e dal collegamento (UTENTI_PERSONALE)
' @param soloCollegamento se true elimina solo il collegamento (UTENTI_PERSONALE)
'Private Sub EliminaDaOrganigramma(soloCollegamento As Boolean)
'    Dim rsDataset As New Recordset
'    Dim nomeTabella As String
'    Dim num As Integer
    
'    If flxGriglia.TextMatrix(vRow, 5) = "Medico" Then
'        nomeTabella = "MEDICI_DIALISI"
'    Else
'        nomeTabella = "INFERMIERI"
'    End If
    
'    rsDataset.Open "SELECT * FROM UTENTI_PERSONALE WHERE CODICE_UTENTE=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
'    num = rsDataset("CODICE_PERSONALE")
'    rsDataset.Delete
'    rsDataset.Close
'    If Not soloCollegamento Then
'        rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & num, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
'        If Not (rsDataset.EOF And rsDataset.BOF) Then
'            rsDataset("ELIMINATO") = True
'            rsDataset.Update
'        End If
'        rsDataset.Close
'    End If
    
'    Set rsDataset = Nothing
'End Sub

'' Verifica se il medico o infermiere è presente nell'organigramma e quindi se c'è il collegamento in UTENTI_PERSONALE
'
' @return true se è presente
'Private Function IsPresente() As Boolean
'    Dim rsDataset As New Recordset
'    rsDataset.Open "SELECT * FROM UTENTI_PERSONALE WHERE CODICE_UTENTE=" & flxGriglia.TextMatrix(flxGriglia.Row, 0), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If Not (rsDataset.EOF And rsDataset.BOF) Then
'        IsPresente = True
'    Else
'        IsPresente = False
'    End If
'    Set rsDataset = Nothing
'End Function

'' Carica il nuovo medico o infermiere nella relativa tabella (MEDICI_DIALISI, INFERMIERI)
'Private Sub InserisciInOrganigramma(keyUtente As Integer)
'    Dim rsDataset As New Recordset
'    Dim nomeTabella As String
'    Dim num As Integer
'    Dim v_Nomi() As Variant
'    Dim v_Val() As Variant

'    If tInput.v_valori(5) = 1 Then
'        nomeTabella = "MEDICI_DIALISI"
'        num = GetNumero(nomeTabella)
'        v_Nomi = Array("KEY", "COGNOME", "NOME")
'        v_Val = Array(num, tInput.v_valori(2), tInput.v_valori(3))
'    Else
'        nomeTabella = "INFERMIERI"
'        num = GetNumero(nomeTabella)
'        v_Nomi = Array("KEY", "COGNOME", "NOME", "MANSIONE")
'        v_Val = Array(num, tInput.v_valori(2), tInput.v_valori(3), 1)
'    End If
    
'    rsDataset.Open nomeTabella, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
'    rsDataset.AddNew v_Nomi, v_Val
'    rsDataset.Update
'    rsDataset.Close
    ' aggiunge anche nella tabella di collegamento
'    rsDataset.Open "UTENTI_PERSONALE", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
'    rsDataset.AddNew
'    rsDataset("KEY") = GetNumero("UTENTI_PERSONALE")
'    rsDataset("CODICE_UTENTE") = keyUtente
'    rsDataset("CODICE_PERSONALE") = num
'    rsDataset("TIPO") = tInput.v_valori(5)
'    rsDataset.Update
'    rsDataset.Close
'    Set rsDataset = Nothing
'End Sub

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

Private Sub cmdElimina_Click()
    Dim eliminato As Boolean
   
    With flxGriglia
        If .Row = 1 Then
            MsgBox "L'utente amministratore non può essere eliminato", vbInformation, "Eliminazione"
            Exit Sub
        End If
        If .Row = 0 Then
            MsgBox "Selezionare l'utente da eliminare", vbCritical, "Attenzione"
        Else
            If MsgBox("Sei sicuro di eliminare: " & .TextMatrix(.Row, 1) & " " & .TextMatrix(.Row, 2) & " ?", vbQuestion + vbYesNo, "Eliminazione") = vbYes Then
                Set rsPassword = New Recordset
                rsPassword.Open "SELECT * FROM LOGIN WHERE KEY=" & .TextMatrix(.Row, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
                If rsPassword.EOF And rsPassword.BOF Then
                    MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
                Else
                    rsPassword("ELIMINATO") = True
                    rsPassword.Update
                    eliminato = True
                End If
                Set rsPassword = Nothing
'                If .TextMatrix(vRow, 5) = "Medico" Or .TextMatrix(vRow, 5) = "Infermiere" Then
'                    If IsPresente Then
'                        If MsgBox("Si vuole CANCELLARE questo utente dall'organigramma?", vbQuestion + vbYesNo, "Cancellazione") = vbYes Then
'                            Call EliminaDaOrganigramma(False)
'                        Else
'                            Call EliminaDaOrganigramma(True)
'                        End If
'                    End If
'                End If
            End If
            If eliminato And TRACCIATO Then
                Call SalvaEliminazione
            End If
            ' lo deve eliminare dopo dalla griglia
            If eliminato Then
                ' rimuove dalla flx
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem (.Row)
                End If
                vRow = 0
                .Row = 0
            End If
        End If
    End With
End Sub

'' Verifica se esiste l'utente
' @param strValore stringa da mostrare in caso di esistenza dell'utente
' @param row riga da non valutare
Private Function EsisteUtente(ByRef strValore As String, cognome As String, nome As String, utente As String, Row As Integer) As Boolean
    Dim i As Integer
    With flxGriglia
        For i = 1 To .Rows - 1
            If UCase(.TextMatrix(i, 1)) = UCase(cognome) And UCase(.TextMatrix(i, 2)) = UCase(nome) And Row <> i Then
                EsisteUtente = True
                strValore = "DATI ANAGRAFICI già presenti"
                Exit Function
            End If
            If UCase(.TextMatrix(i, 3)) = UCase(utente) And Row <> i Then
                EsisteUtente = True
                strValore = "NOME UTENTE già presente"
                Exit Function
            End If
        Next i
    End With
    EsisteUtente = False
End Function

Private Sub cmdInserisci_Click()
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim num As Integer
    Dim primo As Boolean
    Dim strValore As String

    tInput.mantieniDati = False
    tInput.Tipo = tpIPASSWORD
    primo = True
    Do
        If Not primo Then
            MsgBox strValore, vbCritical, "Attenzione"
            tInput.mantieniDati = True
        End If
        Unload frmInput
        frmInput.Show 1
        primo = False
    Loop While EsisteUtente(strValore, tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(1), 0)
       
      ' aggiorna la flx
        Call CaricaFlx
    
    If Not (tInput.v_valori(1) = "") Then
        v_Nomi = Array("KEY", "COGNOME", "NOME", "CHIAVE", "PASSWORD", "TIPO", "DATA")
        num = GetNumero("LOGIN")
        v_Val = Array(num, tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(1), tInput.v_valori(4), tInput.v_valori(5), date)
        
        Set rsPassword = New Recordset
        rsPassword.Open "LOGIN", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsPassword.AddNew v_Nomi, v_Val
        rsPassword.Update
        Set rsPassword = Nothing
        
 '       If (tInput.v_valori(5) = 1 Or tInput.v_valori(5) = 2) Then
 '           If tInput.v_valori(6) Then
                ' inserisce il nuovo utente nell'organigramma come medico o infermiere
 '               Call InserisciInOrganigramma(num)
 '           End If
 '       End If
        
        ' aggiorna la flx
        Call CaricaFlx
        
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
        vRow = flxGriglia.Row
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        If flxGriglia.Row > 14 Then
            flxGriglia.TopRow = flxGriglia.Row
        End If
        
    '    MsgBox "Inserimento effettuato", vbInformation, "Inserimento"
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
            If i >= 14 Or flxGriglia.TopRow > 14 Then
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
        .SetFocus
        If .Row = 1 Then
            If .Col = 1 Or .Col = 2 Or .Col = 4 Then
                txtAppo.Left = .colPos(.Col) + .Left + 45
                txtAppo.Top = .rowPos(.Row) + .Top + 45
                txtAppo.Width = .ColWidth(.Col)
                txtAppo.Text = .TextMatrix(.Row, .Col)
                txtAppo.Visible = True
                txtAppo.SetFocus
            End If
        Else
            If .Col <> 5 Then
                txtAppo.Left = .colPos(.Col) + .Left + 45
                txtAppo.Top = .rowPos(.Row) + .Top + 45
                txtAppo.Width = .ColWidth(.Col)
                txtAppo.Text = .TextMatrix(.Row, .Col)
                txtAppo.Visible = True
                txtAppo.SetFocus
            End If
        End If
    End With
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
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
'-------------------------

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
    If vCol = 1 Or vCol = 2 Then
        txtAppo.MaxLength = 25
    Else
        txtAppo.MaxLength = 20
    End If
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim strObbligatoria As String
    Dim strValore As String
    
    txtAppo.Visible = False
    With flxGriglia
        If UCase(.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
            If EsisteUtente(strValore, IIf(vCol = 1, txtAppo, .TextMatrix(vRow, 1)), IIf(vCol = 2, txtAppo, .TextMatrix(vRow, 2)), IIf(vCol = 3, txtAppo, .TextMatrix(vRow, 3)), vRow) Then
                MsgBox strValore, vbCritical, "Attenzione"
                Exit Sub
            End If
            If txtAppo = "" Then
                strObbligatoria = Choose(vCol, "COGNOME", "NOME", "CODICE UTENTE", "PASSWORD")
                MsgBox "Impossibile memorizzare dati vuoti" & vbCrLf & "Campo: " & strObbligatoria & " obbligatorio", vbCritical, "Attenzione"
                flxGriglia.Row = vRow
                Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Else
                If vCol = 3 Or vCol = 4 Then
                    ' verifica la lunchezza e il not null
                    If Len(txtAppo) < 8 Then
                        MsgBox "Il campo " & strObbligatoria & " deve avere una lunghezza non inferiore a 8 caratteri", vbCritical, "Attenzione"
                        flxGriglia.Row = vRow
                        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
                    Else
                        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
                        cmdAnnulla.Enabled = True
                        flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
                        Call SalvaModifiche
                    End If
                Else
                    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
                    cmdAnnulla.Enabled = True
                    flxGriglia.TextMatrix(vRow, vCol) = UCase(txtAppo.Text)
                    Call SalvaModifiche
                End If
            End If
        End If
    End With
End Sub

