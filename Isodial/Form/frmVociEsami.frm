VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmVociEsami 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esami di Laboratorio"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   14895
      Begin VB.TextBox txtCerca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   840
         TabIndex        =   0
         Top             =   210
         Width           =   3615
      End
      Begin VB.CommandButton cmdCerca 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   240
         Picture         =   "frmVociEsami.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   400
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   14895
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   2640
         TabIndex        =   10
         Top             =   480
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
         Left            =   240
         MaxLength       =   45
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   5025
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   $"frmVociEsami.frx":014E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmVociEsami.frx":0215
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
      TabIndex        =   8
      Top             =   4680
      Width           =   14895
      Begin VB.Frame fraPulsantiInterno 
         BorderStyle     =   0  'None
         Height          =   640
         Left            =   8760
         TabIndex        =   9
         Top             =   120
         Width           =   5895
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
            Left            =   4680
            TabIndex        =   6
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
            Left            =   2040
            TabIndex        =   3
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
            Left            =   720
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
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
            Height          =   480
            Left            =   3360
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmVociEsami"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDatasetCerca As Recordset
Dim rsVoci As Recordset
Dim rsAssEsami As Recordset
Dim rsGruppo As Recordset
Dim keyGruppo As Integer

Dim lettera As String
Dim vRow As Integer         ' riga selezionata
Dim vCol As Integer         ' colonna selezionata
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx
Const icsPOSNEG As String = "  X "
Const icsSTAMPA As String = "               X  "
Const icsSTAMPA_ESAMI As String = "                         X  "

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
        .ColAlignment(1) = vbLeftJustify
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 8
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
    Set objAnnulla = New CAnnulla
    Call CaricaFlx
    
End Sub

Private Sub SalvaModifiche()
    Dim keyId As Integer
    Dim v_Nomi(1 To 8) As Variant
    Dim v_Val(1 To 8) As Variant
    v_Nomi(1) = "KEY"
    v_Nomi(2) = "NOME"
    v_Nomi(3) = "PN"
    v_Nomi(4) = "UNITA"
    v_Nomi(5) = "MIN"
    v_Nomi(6) = "MAX"
    v_Nomi(7) = "STAMPA"
    v_Nomi(8) = "ESAMI_DA_STAMPARE"
    With flxGriglia
        keyId = .TextMatrix(vRow, 0)
        v_Val(1) = .TextMatrix(vRow, 0)
        v_Val(2) = .TextMatrix(vRow, 1)
        If .TextMatrix(vRow, 2) = icsPOSNEG Then
            v_Val(3) = True
            v_Val(4) = ""
            v_Val(5) = vbNull
            v_Val(6) = vbNull
        Else
            v_Val(3) = False
            v_Val(4) = .TextMatrix(vRow, 3)
            v_Val(5) = IIf(.TextMatrix(vRow, 4) = "", 0, .TextMatrix(vRow, 4))
            v_Val(6) = IIf(.TextMatrix(vRow, 5) = "", 0, .TextMatrix(vRow, 5))
        End If
        If .TextMatrix(vRow, 6) = icsSTAMPA Then
            v_Val(7) = True
        Else
            v_Val(7) = False
        End If
        If .TextMatrix(vRow, 7) = icsSTAMPA_ESAMI Then
            v_Val(8) = True
        Else
            v_Val(8) = False
        End If
        Set rsVoci = New Recordset
        rsVoci.Open "SELECT * FROM VOCI_ESAMI WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsVoci.Update v_Nomi, v_Val
        Set rsVoci = Nothing
    End With
End Sub

Private Sub CaricaFlx()
    Dim strSql As String
    ' azzera la griglia
    flxGriglia.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    Set rsVoci = New Recordset
    Set rsGruppo = New Recordset
    Set rsAssEsami = New Recordset
    rsVoci.Open "VOCI_ESAMI ORDER BY NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
 
    If Not (rsVoci.BOF And rsVoci.EOF) Then
        Do While Not rsVoci.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rsVoci("KEY")
                .TextMatrix(.Rows - 1, 1) = rsVoci("NOME") & ""
                If CBool(rsVoci("PN")) Then
                    .TextMatrix(.Rows - 1, 2) = icsPOSNEG
                Else
                    .TextMatrix(.Rows - 1, 2) = ""
                    .TextMatrix(.Rows - 1, 3) = rsVoci("UNITA") & ""
                    .TextMatrix(.Rows - 1, 4) = VirgolaOrPunto(rsVoci("MIN"), ",")
                    .TextMatrix(.Rows - 1, 5) = VirgolaOrPunto(rsVoci("MAX"), ",")
                End If
                If rsVoci("STAMPA") Then
                    .TextMatrix(.Rows - 1, 6) = icsSTAMPA
                Else
                    .TextMatrix(.Rows - 1, 6) = ""
                End If
                If rsVoci("ESAMI_DA_STAMPARE") Then
                    .TextMatrix(.Rows - 1, 7) = icsSTAMPA_ESAMI
                Else
                    .TextMatrix(.Rows - 1, 7) = ""
                End If
                    
            'cerca il codice del gruppo associato all'esame
                strSql = "SELECT CODICE_GRUPPO FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_ESAME =" & rsVoci("KEY")
                rsAssEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsAssEsami.BOF Or rsAssEsami.EOF) Then
                    keyGruppo = rsAssEsami("CODICE_GRUPPO")
                End If
                rsAssEsami.Close
        
            'cerca il gruppo associato all'esame
                If keyGruppo <> 0 Then
                    strSql = "SELECT * FROM GRUPPI_ESAMI WHERE KEY =" & keyGruppo
                    rsGruppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                    .TextMatrix(.Rows - 1, 8) = rsGruppo("NOME")
                    rsGruppo.Close
                Else
                    .TextMatrix(.Rows - 1, 8) = ""
                End If
                keyGruppo = 0
                rsVoci.MoveNext
            End With
        Loop
        Set rsVoci = Nothing
        Set rsGruppo = Nothing
        Set rsAssEsami = Nothing
        flxGriglia.Row = 0
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

Private Sub cmdElimina_Click()
    Dim blnElimina As Boolean
    Dim intKey As Integer
    Dim strNome As String
    Dim cmCommand As Command
   
    With flxGriglia
        If .Row = 0 Then
            MsgBox "Selezionare la Voce per Esami di Laboratorio da eliminare", vbCritical, "Attenzione"
        Else
            intKey = .TextMatrix(.Row, 0)
            strNome = .TextMatrix(.Row, 1)
            
            blnElimina = IsPossibleDelete("ESAMI_LAB", "CODICE_ESAME", intKey)
            
            If blnElimina Then
                If MsgBox("Sicuro di voler eliminare " & strNome & "?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                    
                    Set cmCommand = New Command
                    cmCommand.CommandType = adCmdText
                    cmCommand.ActiveConnection = cnPrinc
                    
                    cmCommand.CommandText = "Delete From Associazione_Esami_Lab Where Codice_Esame=" & intKey
                    cmCommand.Execute
                    
                    cmCommand.CommandText = "Delete From VOCI_ESAMI Where KEY=" & intKey
                    cmCommand.Execute
                                        
                    Set cmCommand = Nothing
                    
                    ' rimuove dalla flx
                    If .Rows = 2 Then
                        .Rows = 1
                    Else
                        .RemoveItem (.Row)
                    End If
                    vRow = 0
                    .Row = 0
                    
                    ' aggiorna il form raggruppamenti
                    Dim lForm As Form
                    For Each lForm In Forms
                        If lForm.Name = "frmTipiEsamiLab" Then
                            Set lForm = frmTipiEsamiLab
                            lForm.CaricaFlxVoci (lForm.flxNomi.TextMatrix(lForm.flxNomi.Row, 0))
                        End If
                    Next
                    
                    
                    MsgBox "Eliminazione avvenuta con successo", vbInformation, Me.Caption
                End If
            Else
                MsgBox "Impossibile eliminare " & strNome & " perchè in relazione con altri dati del sistema", vbInformation, Me.Caption
            End If
        End If
    End With
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdInserisci_Click()
    Dim v_Nomi(1 To 8) As Variant
    Dim v_Val() As Variant
    Dim num As Integer
    Dim primo As Boolean
    Dim ValoreMin As String
    Dim ValoreMax As String
    
    
    primo = True
    tInput.mantieniDati = False
    tInput.Tipo = tpIVOCI
    Do
        If Not primo Then
            MsgBox "L' esame inserito è già presente", vbCritical, "Attenzione"
            tInput.mantieniDati = True
        End If
        Unload frmInput
        tInput.v_valori(6) = numStampa
        tInput.v_valori(7) = numStampaEsami
        frmInput.Show 1
        primo = False
    Loop While Esiste(flxGriglia, 1, 0, tInput.v_valori(1))
    
    If Not (tInput.v_valori(1) = "") Then
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "NOME"
        v_Nomi(3) = "PN"
        v_Nomi(4) = "UNITA"
        v_Nomi(5) = "MIN"
        v_Nomi(6) = "MAX"
        v_Nomi(7) = "STAMPA"
        v_Nomi(8) = "ESAMI_DA_STAMPARE"
        
        num = GetNumero("VOCI_ESAMI")
        ValoreMin = tInput.v_valori(4)
        ValoreMax = tInput.v_valori(5)
        
        v_Val = Array(num, tInput.v_valori(1), CBool(tInput.v_valori(2)), tInput.v_valori(3), ValoreMin, ValoreMax, CBool(tInput.v_valori(6)), CBool(tInput.v_valori(7)))
        
        Set rsVoci = New Recordset
        rsVoci.Open "VOCI_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsVoci.AddNew v_Nomi, v_Val
        rsVoci.Update
        Set rsVoci = Nothing
        
        ' aggiorna la flx
        flxGriglia.Rows = 1
        Call CaricaFlx
        
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, vRow, num)
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        If flxGriglia.Row > 11 Then
            flxGriglia.TopRow = flxGriglia.Row
        End If
        
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
        vCol = flxGriglia.Col
        vRow = flxGriglia.Row
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_DblClick()
    ' fase di modifica
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    With flxGriglia
        .SetFocus
        ' esce se ha accettato valori pos e neg
        If (.Col = 3 Or .Col = 4 Or .Col = 5) And .TextMatrix(.Row, 2) = " X " Then Exit Sub
        If .Col = 2 Then
            Call objAnnulla.Add(.TextMatrix(.Row, .Col), .Col, .TextMatrix(.Row, 0))
            cmdAnnulla.Enabled = True
            ' mette o toglie una X per il campo valori positivo negativo
            If .TextMatrix(.Row, 2) = "" Then
                .TextMatrix(.Row, 2) = icsPOSNEG
                ' cancella le altre informazioni e salva
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
            Else
                .TextMatrix(.Row, 2) = ""
            End If
            Call SalvaModifiche
        ElseIf .Col = 6 Then
            Call objAnnulla.Add(.TextMatrix(.Row, .Col), .Col, .TextMatrix(.Row, 0))
            cmdAnnulla.Enabled = True
            If .TextMatrix(.Row, 6) = "" Then
                If numStampa < 3 Then
                    .TextMatrix(.Row, 6) = icsSTAMPA
                Else
                    MsgBox "Impossibile stampare più di tre esami nella cartella clinica", vbCritical, "Attenzione"
                    Exit Sub
                End If
            Else
                .TextMatrix(.Row, 6) = ""
            End If
            Call SalvaModifiche
        ElseIf .Col = 7 Then
            Call objAnnulla.Add(.TextMatrix(.Row, .Col), .Col, .TextMatrix(.Row, 0))
            cmdAnnulla.Enabled = True
            If .TextMatrix(.Row, 7) = "" Then
                If numStampaEsami < 16 Then
                    .TextMatrix(.Row, 7) = icsSTAMPA_ESAMI
                Else
                    MsgBox "Impossibile stampare più di 16 esami nella Scheda Dialitica Settimanale", vbCritical, "Attenzione"
                    Exit Sub
                End If
            Else
                .TextMatrix(.Row, 7) = ""
            End If
            Call SalvaModifiche
        ElseIf .Col = 8 Then
           'non permette l'editing della colonna
        Else
            ' altri campi
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        End If
    End With
End Sub

Private Function numStampaEsami() As Integer
    Dim i As Integer
    Dim num As Integer
    
    For i = 1 To flxGriglia.Rows - 1
        If flxGriglia.TextMatrix(i, 7) = icsSTAMPA_ESAMI Then
            num = num + 1
        End If
    Next i
    
    numStampaEsami = num
End Function

Private Function numStampa() As Integer
    Dim i As Integer
    Dim num As Integer
    
    For i = 1 To flxGriglia.Rows - 1
        If flxGriglia.TextMatrix(i, 6) = icsSTAMPA Then
            num = num + 1
        End If
    Next i
    
    numStampa = num
End Function

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
            If i >= 11 Or flxGriglia.TopRow > 11 Then
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

Private Sub txtAppo_Change()
    If flxGriglia.Col = 4 Or flxGriglia.Col = 5 Then
        If Not (lettera = "." Or lettera = "") Then
            Call OnlyNumber(txtAppo, lettera)
        End If
    End If
End Sub

Private Sub txtAppo_GotFocus()
    If flxGriglia.Col = 4 Or flxGriglia.Col = 5 Then
        txtAppo.Alignment = 1 'destra per i numeri
    Else
        txtAppo.Alignment = 0 'sinistra
    End If
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim controlla As Boolean
    Dim min As Double
    Dim max As Double
    
    txtAppo.Visible = False
    If UCase(flxGriglia.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
        If txtAppo = "" Then
            MsgBox "Impossibile memorizzare dati vuoti", vbCritical, "Attenzione"
            flxGriglia.Row = vRow
            flxGriglia.Col = vCol
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Sub
        End If
        If vCol = 1 Then
            If Esiste(flxGriglia, 1, vRow, txtAppo) Then
                MsgBox "Il nome inserito è già presente", vbCritical, "Attenzione"
            Else
                Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
                cmdAnnulla.Enabled = True
                flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
                Call SalvaModifiche
            End If
        ElseIf vCol = 4 Or vCol = 5 Then
            ' verifica la corettezza del min e max
            If txtAppo = "" Then
                txtAppo = 0
            End If
            
            If vCol = 4 Then
                controlla = (flxGriglia.TextMatrix(vRow, 5) <> "")
                min = CDbl(txtAppo)
                If controlla Then
                    max = CDbl(flxGriglia.TextMatrix(vRow, 5))
                End If
            ElseIf vCol = 5 Then
                controlla = (flxGriglia.TextMatrix(vRow, 4) <> "")
                max = CDbl(txtAppo)
                If controlla Then
                    min = CDbl(flxGriglia.TextMatrix(vRow, 4))
                End If
            End If
            If controlla Then
                If min > max Then
                    ' min > max
                    MsgBox "Errato inserimento dei valori max e min", vbCritical, "Attenzione"
                Else
                    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
                    cmdAnnulla.Enabled = True
                    flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
                    Call SalvaModifiche
                End If
            Else
                Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
                cmdAnnulla.Enabled = True
                flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
                Call SalvaModifiche
            End If
        Else
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, flxGriglia.TextMatrix(vRow, 0))
            cmdAnnulla.Enabled = True
            flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
            Call SalvaModifiche
        End If
    End If
End Sub

Private Sub txtAppo_Validate(Cancel As Boolean)
    If flxGriglia.Col = 4 Or flxGriglia.Col = 5 Then
        If txtAppo = "" Then
            Cancel = False
        Else
            Cancel = ControlloNumerico(txtAppo.Text)
        End If
    Else
        Cancel = False
    End If
End Sub

Private Sub txtCerca_Change()
    Call Cerca
End Sub

Private Sub txtCerca_GotFocus()
    txtCerca.BackColor = colArancione
End Sub

Private Sub txtCerca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdChiudi_Click
    End If
End Sub

Private Sub txtCerca_LostFocus()
    txtCerca.BackColor = vbWhite
End Sub

Private Sub Cerca()
    ' cerca l'esame
    Dim chiaveRic As String
    Dim strSql As String
    Dim condizione As String
    
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    chiaveRic = UCase(txtCerca.Text)
        
        condizione = IIf(tTrova.condizione <> "", " AND ", "") & tTrova.condizione
        strSql = "SELECT * FROM VOCI_ESAMI WHERE NOME LIKE '" & Apostrophe(chiaveRic) & "%' " & condizione & "ORDER BY NOME"
        Set rsDatasetCerca = New Recordset
        Set rsGruppo = New Recordset
        Set rsAssEsami = New Recordset
        rsDatasetCerca.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsDatasetCerca.EOF
         With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDatasetCerca("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDatasetCerca("NOME")
            .TextMatrix(.Rows - 1, 2) = rsDatasetCerca("PN")
            If CBool(rsDatasetCerca("PN")) Then
                .TextMatrix(.Rows - 1, 2) = icsPOSNEG
            Else
                .TextMatrix(.Rows - 1, 2) = ""
            End If
            .TextMatrix(.Rows - 1, 3) = rsDatasetCerca("UNITA") & ""
            .TextMatrix(.Rows - 1, 4) = VirgolaOrPunto(rsDatasetCerca("MIN"), ",")
            .TextMatrix(.Rows - 1, 5) = VirgolaOrPunto(rsDatasetCerca("MAX"), ",")
            If rsDatasetCerca("STAMPA") Then
                .TextMatrix(.Rows - 1, 6) = icsSTAMPA
            Else
                .TextMatrix(.Rows - 1, 6) = ""
            End If
            If rsDatasetCerca("ESAMI_DA_STAMPARE") Then
              .TextMatrix(.Rows - 1, 7) = icsSTAMPA_ESAMI
            Else
              .TextMatrix(.Rows - 1, 7) = ""
            End If
           
           'cerca il codice del gruppo associato all'esame
            strSql = "SELECT CODICE_GRUPPO FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_ESAME =" & rsDatasetCerca("KEY")
            rsAssEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsAssEsami.BOF Or rsAssEsami.EOF) Then
                keyGruppo = rsAssEsami("CODICE_GRUPPO")
            End If
            rsAssEsami.Close
        
            'cerca il gruppo associato all'esame
            If keyGruppo <> 0 Then
                strSql = "SELECT * FROM GRUPPI_ESAMI WHERE KEY =" & keyGruppo
                rsGruppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                .TextMatrix(.Rows - 1, 8) = rsGruppo("NOME")
                rsGruppo.Close
             Else
                .TextMatrix(.Rows - 1, 8) = ""
             End If
                       
            rsDatasetCerca.MoveNext
            End With
        Loop
        keyGruppo = 0
        flxGriglia.Row = 0
        Set rsDatasetCerca = Nothing
        Set rsGruppo = Nothing
        Set rsAssEsami = Nothing
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

