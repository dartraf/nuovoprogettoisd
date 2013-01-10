VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmTsat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calcolo TSAT %"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabScheda 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6773
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tabella"
      TabPicture(0)   =   "frmTsat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Grafico 2D"
      TabPicture(1)   =   "frmTsat.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grafico(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Grafico 3D"
      TabPicture(2)   =   "frmTsat.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grafico(1)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
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
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11895
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
            ItemData        =   "frmTsat.frx":0054
            Left            =   960
            List            =   "frmTsat.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtAppo 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   720
         End
         Begin MSFlexGridLib.MSFlexGrid flxGriglia 
            Height          =   1695
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   2990
            _Version        =   393216
            Rows            =   5
            Cols            =   13
            ScrollTrack     =   -1  'True
            MousePointer    =   15
            FormatString    =   $"frmTsat.frx":0058
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmTsat.frx":0134
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Anno"
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
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   11895
         Begin VB.CommandButton cmdEsportaEsame 
            Caption         =   "&Esporta TSAT%"
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
            Left            =   3840
            TabIndex        =   21
            Top             =   240
            Width           =   1935
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
            Left            =   2280
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdImportaEsami 
            Caption         =   "&Importa esami"
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
            Left            =   6000
            TabIndex        =   10
            Top             =   240
            Width           =   1815
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
            Height          =   480
            Left            =   8040
            TabIndex        =   8
            Top             =   240
            Width           =   2295
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
            Left            =   10560
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   3495
         Index           =   0
         Left            =   -74880
         OleObjectBlob   =   "frmTsat.frx":028E
         TabIndex        =   5
         Top             =   360
         Width           =   11895
      End
      Begin MSChart20Lib.MSChart grafico 
         Height          =   3495
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmTsat.frx":2F3A
         TabIndex        =   9
         Top             =   360
         Width           =   11895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmTsat.frx":5BCF
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblCognome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6840
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblEta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   11160
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anni"
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
         Index           =   3
         Left            =   10440
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmTsat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lettera As String
Dim rsTSAT As Recordset
Dim vCol As Integer
Dim vRow As Integer
Dim objAnnulla As CAnnulla
Dim stoCaricando As Boolean
Dim rsDialisi As Recordset
Dim intPazientiKey As Integer

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim k As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxGriglia
        .MousePointer = flexCustom
        .Row = 0
        For i = 0 To flxGriglia.Cols - 1
            .Col = i
            .CellFontBold = True
        Next i
        .Col = 0
        For i = 1 To flxGriglia.Rows - 1
            .Row = i
            .CellFontBold = True
            .CellBackColor = RGB(231, 255, 255)
        Next i
    End With
    stoCaricando = True
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    stoCaricando = False
    For k = 0 To 1
        ' valore massimo di ktv è 4
        grafico(k).Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 100
        For i = 1 To 12
            grafico(k).Column = 1
            grafico(k).Row = i
            grafico(k).data = 0
            grafico(k).RowLabel = UCase(MonthName(i, True))
        Next i
    Next k
    tabScheda.Tab = 0
    Set objAnnulla = New CAnnulla
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intPazientiKey = 0
End Sub

Private Sub SalvaModifiche()
    Dim nome_campo As String
    With flxGriglia
        Select Case vRow
            Case 1: nome_campo = "GIORNO"
            Case 2: nome_campo = "SIDEREMIA"
            Case 3: nome_campo = "TRANSFERRINA"
        End Select
        Set rsTSAT = New Recordset
        rsTSAT.Open "SELECT * FROM TSAT WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND MESE=" & vCol & " AND ANNO=" & cboAnno.Text, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If rsTSAT.EOF And rsTSAT.BOF Then
            ' nn esiste e quindi lo aggiunge
            rsTSAT.AddNew
            rsTSAT("KEY") = GetNumero("TSAT")
            rsTSAT("CODICE_PAZIENTE") = intPazientiKey
            rsTSAT("ANNO") = cboAnno.Text
            rsTSAT("MESE") = vCol
            rsTSAT(nome_campo) = IIf(.TextMatrix(vRow, vCol) = "", Null, VirgolaOrPunto(.TextMatrix(vRow, vCol), ","))
            rsTSAT.Update
        Else
           '  esiste e lo modifica
            rsTSAT(nome_campo) = IIf(.TextMatrix(vRow, vCol) = "", Null, VirgolaOrPunto(.TextMatrix(vRow, vCol), ","))
            rsTSAT.Update
        End If
        Set rsTSAT = Nothing
    End With
End Sub

Private Sub PulisciTutto()
    intPazientiKey = 0
    lblCognome = ""
    lblEta = ""
    lblNome = ""
    Call Pulisci
    objAnnulla.Refresh
End Sub

Private Sub Pulisci()
    Dim i As Integer
    Dim k As Integer
    For i = 1 To 12
        For k = 1 To flxGriglia.Rows - 1
            flxGriglia.TextMatrix(k, i) = ""
        Next k
        For k = 0 To 1
            grafico(k).Column = 1
            grafico(k).Row = i
            grafico(k).data = 0
        Next k
    Next i
End Sub

Private Function CalcolaTsat(vCol As Integer) As Double
    On Error GoTo gestione

    'Sideremia/Transferrina*70,9
    With flxGriglia
        If .TextMatrix(2, vCol) <> "" And .TextMatrix(3, vCol) <> "" Then
            CalcolaTsat = Format(VirgolaOrPunto(.TextMatrix(2, vCol), ".") / VirgolaOrPunto(.TextMatrix(3, vCol), ".") * CSng("70,9"), "##.##")
        Else
            Exit Function
        End If
    End With
    
    Exit Function
gestione:
    If Err.Number = 13 Or Err.Number = 5 Then
        'MsgBox "Impossibile calcolare il valore del TSAT con i dati presenti", vbCritical, "Attenzione"
    Else
        MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
    End If
End Function

Public Sub ColoraColonna(Optional colore As ColorConstants = vbCyan)
    ' colora l'intera colonna di una flex
    Dim i As Integer
    Dim k As Integer
    Dim Col As Integer
    Dim riga As Integer
    Dim colAppo As ColorConstants
    
    If flxGriglia.Row = 0 Or flxGriglia.Col = 0 Then Exit Sub
    
    riga = flxGriglia.Row
    Col = flxGriglia.Col
    ' discolora la colonna colorata
    For k = 1 To flxGriglia.Cols - 1
        flxGriglia.Col = k
        ' utilizzo un var di appoggio perche cosi funziona
        colAppo = flxGriglia.CellBackColor
        If colAppo <> vbWhite And colAppo <> 0 Then
            For i = flxGriglia.FixedRows To flxGriglia.Rows - 1
                flxGriglia.Row = i
                flxGriglia.CellBackColor = vbWhite
            Next i
            Exit For
        End If
    Next k
    flxGriglia.Col = Col
    ' cambia colore della riga
    For i = flxGriglia.FixedRows To flxGriglia.Rows - 1
        flxGriglia.Row = i
        flxGriglia.CellBackColor = colore
    Next i
    flxGriglia.Row = riga
End Sub

Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        Exit Sub
    End If
    ' carica i dati del paziente
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    lblCognome = rsDataset("COGNOME")
    lblNome = rsDataset("NOME")
    Dim somma As Integer
    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
    Set rsDataset = Nothing
    ' carica la scheda
    Call CaricaScheda
End Sub

'' elimina l'esame tsat del mese
Private Sub Elimina()
    Set rsTSAT = New Recordset
    rsTSAT.Open "SELECT * FROM TSAT WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND MESE=" & vCol & " AND ANNO=" & cboAnno.Text, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsTSAT.EOF And rsTSAT.BOF) Then
        rsTSAT.Delete
        rsTSAT.Update
    End If
    Set rsTSAT = Nothing
End Sub

Private Sub CaricaScheda()
    Dim i As Integer
    Dim k As Integer
    Dim valore As Single
    vRow = 0
    vCol = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnulla.Enabled = False
    With flxGriglia
        For i = 1 To 12
            Set rsTSAT = New Recordset
            rsTSAT.Open "SELECT * FROM TSAT WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND ANNO=" & cboAnno.Text & " AND MESE=" & i, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsTSAT.EOF And rsTSAT.BOF) Then
                .TextMatrix(1, i) = rsTSAT("GIORNO") & ""
                .TextMatrix(2, i) = VirgolaOrPunto(rsTSAT("SIDEREMIA") & "", ",")
                .TextMatrix(3, i) = VirgolaOrPunto(rsTSAT("TRANSFERRINA") & "", ",")
                
                valore = CalcolaTsat(i)
                .TextMatrix(4, i) = VirgolaOrPunto(CStr(valore), ",")
                For k = 0 To 1
                    grafico(k).Column = 1
                    grafico(k).Row = i
                    grafico(k).data = valore
                Next k
            End If
            rsTSAT.Close
        Next i
        Set rsTSAT = Nothing
    End With
End Sub

Private Sub cmdEsportaEsame_Click()
    Dim rsDataset As Recordset
    Dim keyEsame As Integer
    Dim keyGruppo As Integer
    Dim keyAnamnesi As Integer
    Dim keyRecord As Integer
    Dim keyNuovo As Integer
    
    If flxGriglia.TextMatrix(4, vCol) = "" Then
        MsgBox "IMPOSSIBILE ESPORTARE!!!" & vbCrLf & "Valori non definiti", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    If flxGriglia.Col = 0 Then
        MsgBox "Selezionare il mese dell'esame da esportare", vbCritical, "Attenzione"
    Else
        If flxGriglia.TextMatrix(4, vCol) <> 0 Then
            Set rsDataset = New Recordset
            
            ' verifica se esiste l'esame kt/v
            rsDataset.Open "SELECT * FROM VOCI_ESAMI WHERE NOME like '%TSAT%'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                keyEsame = rsDataset("KEY")
            Else
                MsgBox "Esame TSAT% non presente nella lista esami di laboratorio", vbCritical, "Attenzione"
            End If
            rsDataset.Close
            
            If keyEsame <> 0 Then
                ' verifica se esiste l'associazione con qualche gruppo esami lab
                rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_ESAME=" & keyEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    keyGruppo = rsDataset("CODICE_GRUPPO")
                Else
                    MsgBox "Esame TSAT% non è associato a nessun raggruppamento esami di laboratorio", vbCritical, "Attenzione"
                End If
                rsDataset.Close
                
                If keyGruppo <> 0 Then
                    ' verifica se esiste un record anamnesi per il mese selezionato del gruppo
                    rsDataset.Open "SELECT * FROM ANAMNESI_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & keyGruppo & " AND DATA BETWEEN #" & Format(vCol, "00") & "/01/" & cboAnno.Text & "# AND #" & GetUltimoGiorno(vCol, cboAnno.Text, True) & "#", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                    If Not (rsDataset.EOF And rsDataset.BOF) Then
                        If rsDataset.RecordCount > 1 Then
                            rsDataset.Filter = "DATA>=" & DateValue(flxGriglia.TextMatrix(1, vCol) & "/" & vCol & "/" & cboAnno.Text)
                            If rsDataset.RecordCount > 1 Then
                                ' mostra un pannello con le date filtrare e fa scegliere all'utente
                                tElenca.Tipo = tpESPORTAESAMI
                                tElenca.condizione = "WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & keyGruppo & " AND DATA>= #" & DateValue(vCol & "/" & flxGriglia.TextMatrix(1, vCol) & "/" & cboAnno.Text) & "# AND DATA<=#" & GetUltimoGiorno(vCol, cboAnno.Text, True) & "#"
                                frmElencaDate.Show 1
                                If laData = "" Then Exit Sub
                                rsDataset.Filter = "DATA=#" & laData & "#"
                                If rsDataset.RecordCount > 0 Then
                                    keyAnamnesi = rsDataset("KEY")
                                Else
                                    keyAnamnesi = 0
                                End If
                            Else
                                keyAnamnesi = rsDataset("KEY")
                            End If
                            rsDataset.Filter = ""
                        Else
                            keyAnamnesi = rsDataset("KEY")
                        End If
                    Else
                        keyAnamnesi = 0
                    End If
                    rsDataset.Close
                    
                    If keyAnamnesi <> 0 Then
                        ' verifica se esiste un record per l'esame per l'anamnesi trovata
                        rsDataset.Open "SELECT * FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & keyAnamnesi & " AND CODICE_ESAME=" & keyEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                        If Not (rsDataset.EOF And rsDataset.BOF) Then
                           keyRecord = rsDataset("KEY")
                        End If
                        rsDataset.Close
                        
                        If keyRecord <> 0 Then
                            If MsgBox("Il valore del Tsat% è già presente" & vbCrLf & "Vuoi sovrascriverlo?", vbQuestion + vbYesNo + vbDefaultButton2, "Esporta esame") = vbYes Then
                                ' modifica il dato esistente
                                rsDataset.Open "SELECT * FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & keyAnamnesi & " AND CODICE_ESAME=" & keyEsame, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                                If Not (rsDataset.EOF And rsDataset.BOF) Then
                                   rsDataset("VALORE") = flxGriglia.TextMatrix(4, vCol)
                                   rsDataset.Update
                                End If
                                rsDataset.Close
                                MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                            End If
                        Else
                            ' inserisce un nuovo esame
                            keyNuovo = GetNumero("ESAMI_LAB")
                            rsDataset.Open "ESAMI_LAB", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
                            rsDataset.AddNew
                            rsDataset("KEY") = keyNuovo
                            rsDataset("CODICE_ESAME") = keyEsame
                            rsDataset("CODICE_ANAMNESI_ESAMI") = keyAnamnesi
                            rsDataset("VALORE") = flxGriglia.TextMatrix(4, vCol)
                            rsDataset.Update
                            rsDataset.Close
                            MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                        End If
                    Else
                        ' inserisce la nuova data del gruppo esami
                        keyAnamnesi = GetNumero("ANAMNESI_ESAMI")
                        rsDataset.Open "ANAMNESI_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                        rsDataset.AddNew
                        rsDataset("KEY") = keyAnamnesi
                        rsDataset("CODICE_PAZIENTE") = intPazientiKey
                        rsDataset("DATA") = DateValue(flxGriglia.TextMatrix(1, vCol) & "/" & vCol & "/" & cboAnno.Text)
                        rsDataset("CODICE_GRUPPO") = keyGruppo
                        rsDataset("UTENTE_MODIFICATORE") = tAccesso.key
                        rsDataset.Update
                        rsDataset.Close
                        
                        ' inserisce un nuovo esame
                        keyNuovo = GetNumero("ESAMI_LAB")
                        rsDataset.Open "ESAMI_LAB", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
                        rsDataset.AddNew
                        rsDataset("KEY") = keyNuovo
                        rsDataset("CODICE_ESAME") = keyEsame
                        rsDataset("CODICE_ANAMNESI_ESAMI") = keyAnamnesi
                        rsDataset("VALORE") = flxGriglia.TextMatrix(4, vCol)
                        rsDataset.Update
                        rsDataset.Close
                        MsgBox "Esame esportato con successo", vbInformation, "Esporta esame"
                    End If
                End If
            End If
            Set rsDataset = Nothing
        Else
            MsgBox "Per il mese selezionato impossibile calcolare il TSAT%", vbCritical, "Attenzione"
        End If
    End If
End Sub

Private Sub cmdImportaEsami_Click()
    Dim rsDataset As Recordset
    Dim continua As Boolean
    Dim strSql As String
    
    If flxGriglia.Col = 0 Then
        MsgBox "Selezionare il mese degli esami da importare", vbInformation, "Informazione"
    Else
        strSql = "SELECT    VOCI_ESAMI.NOME AS VOCI_ESAMINOME, VALORE, DATA " & _
                "FROM       ((ANAMNESI_ESAMI " & _
                "           INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                "           INNER JOIN VOCI_ESAMI ON VOCI_ESAMI.KEY=ESAMI_LAB.CODICE_ESAME) " & _
                "WHERE      CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                "           (NOME LIKE '%Sideremia%' OR NOME LIKE '%Trasferrina%' OR NOME LIKE '%Transferrina%') AND " & _
                "           MONTH([DATA])=" & flxGriglia.Col & " AND " & _
                "           YEAR([DATA])=" & cboAnno.Text & " " & _
                "ORDER BY   DATA DESC"
        Set rsDataset = New Recordset
        rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            If flxGriglia.TextMatrix(2, flxGriglia.Col) <> "" Then
                If MsgBox("I valori degli esami sono già presenti" & vbCrLf & "Vuoi sovrascriverli?", vbQuestion + vbYesNo + vbDefaultButton2, "Importa esami") = vbYes Then
                    continua = True
                Else
                    continua = False
                End If
            Else
                continua = True
            End If
            If continua Then
                rsDataset.Filter = "VOCI_ESAMINOME LIKE '%Trasferrina%' OR VOCI_ESAMINOME LIKE '%Transferrina%'"
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    flxGriglia.TextMatrix(3, flxGriglia.Col) = VirgolaOrPunto(rsDataset("VALORE"), ",")
                    vRow = 3
                    Call SalvaModifiche
                End If
                rsDataset.Filter = "VOCI_ESAMINOME LIKE '%Sideremia%'"
                If Not (rsDataset.EOF And rsDataset.BOF) Then
                    flxGriglia.TextMatrix(2, flxGriglia.Col) = VirgolaOrPunto(rsDataset("VALORE"), ",")
                    vRow = 2
                    Call SalvaModifiche
                End If
                ' tsat
                flxGriglia.TextMatrix(4, flxGriglia.Col) = VirgolaOrPunto(CStr(CalcolaTsat(flxGriglia.Col)), ",")
                ' giorno
                flxGriglia.TextMatrix(1, flxGriglia.Col) = Day(rsDataset("DATA"))
                vRow = 1
                Call SalvaModifiche
            End If
        Else
            MsgBox "Nessun esame per il mese selezionato", vbInformation, "Importa esami"
        End If
        Set rsDataset = Nothing
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim codiceId As Integer
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Attenzione"
        Exit Sub
    End If

      
    Set rsDialisi = New Recordset
    rsDialisi.Open "SELECT COGNOME, NOME, DATA_NASCITA, CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsDialisi("COGNOME") & " " & rsDialisi("NOME")
    structIntestazione.sDataPaziente = rsDialisi("DATA_NASCITA")
    codiceId = rsDialisi("CODICE_ID")
    Set rsDialisi = Nothing

    strSql = "SHAPE APPEND  NEW adVarChar (10) as GIORNO_GEN, " & _
                    "       NEW adVarChar (10) as GIORNO_FEB, " & _
                    "       NEW adVarChar (10) as GIORNO_MAR, " & _
                    "       NEW adVarChar (10) as GIORNO_APR, " & _
                    "       NEW adVarChar (10) as GIORNO_MAG, " & _
                    "       NEW adVarChar (10) as GIORNO_GIU, " & _
                    "       NEW adVarChar (10) as GIORNO_LUG, " & _
                    "       NEW adVarChar (10) as GIORNO_AGO, " & _
                    "       NEW adVarChar (10) as GIORNO_SET, " & _
                    "       NEW adVarChar (10) as GIORNO_OTT, " & _
                    "       NEW adVarChar (10) as GIORNO_NOV, " & _
                    "       NEW adVarChar (10) as GIORNO_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as SIDEREMIA_GEN, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_FEB, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_MAR, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_APR, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_MAG, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_GIU, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_LUG, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_AGO, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_SET, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_OTT, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_NOV, " & _
                    "       NEW adVarChar (10) as SIDEREMIA_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as TRANSFERRINA_GEN, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_FEB, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_MAR, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_APR, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_MAG, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_GIU, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_LUG, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_AGO, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_SET, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_OTT, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_NOV, " & _
                    "       NEW adVarChar (10) as TRANSFERRINA_DIC, "
    strSql = strSql & _
                    "       NEW adVarChar (10) as TSAT_GEN, " & _
                    "       NEW adVarChar (10) as TSAT_FEB, " & _
                    "       NEW adVarChar (10) as TSAT_MAR, " & _
                    "       NEW adVarChar (10) as TSAT_APR, " & _
                    "       NEW adVarChar (10) as TSAT_MAG, " & _
                    "       NEW adVarChar (10) as TSAT_GIU, " & _
                    "       NEW adVarChar (10) as TSAT_LUG, " & _
                    "       NEW adVarChar (10) as TSAT_AGO, " & _
                    "       NEW adVarChar (10) as TSAT_SET, " & _
                    "       NEW adVarChar (10) as TSAT_OTT, " & _
                    "       NEW adVarChar (10) as TSAT_NOV, " & _
                    "       NEW adVarChar (10) as TSAT_DIC "

                 
         
     ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSql, cnConn, adOpenStatic, adLockOptimistic
        
    Dim vett(1 To 12) As String
    vett(1) = "GEN"
    vett(2) = "FEB"
    vett(3) = "MAR"
    vett(4) = "APR"
    vett(5) = "MAG"
    vett(6) = "GIU"
    vett(7) = "LUG"
    vett(8) = "AGO"
    vett(9) = "SET"
    vett(10) = "OTT"
    vett(11) = "NOV"
    vett(12) = "DIC"
    Dim i As Integer
    Dim valore As Double
    
    With rsMain
        .AddNew
        
        Set rsTSAT = New Recordset
        For i = 1 To 12
            rsTSAT.Open "SELECT * FROM TSAT WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND ANNO=" & cboAnno.Text & " AND MESE=" & i, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not (rsTSAT.EOF And rsTSAT.BOF) Then
                .Fields("GIORNO_" & vett(i)) = rsTSAT("GIORNO") & ""
                .Fields("SIDEREMIA_" & vett(i)) = VirgolaOrPunto(rsTSAT("SIDEREMIA") & "", ",")
                .Fields("TRANSFERRINA_" & vett(i)) = VirgolaOrPunto(rsTSAT("TRANSFERRINA") & "", ",")
                If .Fields("SIDEREMIA_" & vett(i)) <> "" And .Fields("TRANSFERRINA_" & vett(i)) <> "" Then
                    valore = Format(.Fields("SIDEREMIA_" & vett(i)) / .Fields("TRANSFERRINA_" & vett(i)) * CSng("70,9"), "##.##")
                    .Fields("TSAT_" & vett(i)) = VirgolaOrPunto(CStr(valore), ",")
                Else
                    .Fields("TSAT_" & vett(i)) = ""
                End If
            Else
                .Fields("GIORNO_" & vett(i)) = ""
                .Fields("SIDEREMIA_" & vett(i)) = ""
                .Fields("TRANSFERRINA_" & vett(i)) = ""
                .Fields("TSAT_" & vett(i)) = ""
            End If
            rsTSAT.Close
        Next i
        Set rsTSAT = Nothing
    End With

    Set rptCalcoloTsat.DataSource = rsMain
    rptCalcoloTsat.TopMargin = 0
    rptCalcoloTsat.BottomMargin = 0
    rptCalcoloTsat.Sections("Intestazione").Controls.Item("lblPaziente").Caption = structIntestazione.sPaziente
    rptCalcoloTsat.Sections("Intestazione").Controls.Item("lblDataNascita").Caption = structIntestazione.sDataPaziente
    rptCalcoloTsat.Sections("Intestazione").Controls.Item("lblEta").Caption = lblEta.Caption
    rptCalcoloTsat.Sections("Intestazione").Controls.Item("lblAnno").Caption = cboAnno.Text
    rptCalcoloTsat.PrintReport True, rptRangeAllPages
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdAnnulla_Click()
    Dim k As Integer
    Dim Dato As String
    Dim Col As Integer
    Dim Row As Integer
    Dim valore As Single
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    ' row identifica la riga e non il key
    Row = objAnnulla.Row
    With flxGriglia
        .TextMatrix(Row, Col) = Dato
        valore = CalcolaTsat(Col)
        .TextMatrix(4, Col) = VirgolaOrPunto(CStr(valore), ",")
        For k = 0 To 1
            grafico(k).Column = 1
            grafico(k).Row = Col
            grafico(k).data = valore
        Next k
        objAnnulla.Remove
        ' modifica anche il db
        vRow = Row
        Call SalvaModifiche
        If objAnnulla.Vuoto = True Then
            cmdAnnulla.Enabled = False
        End If
    End With
End Sub

Private Sub flxGriglia_Click()
    
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraColonna(vbWhite)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        vCol = flxGriglia.Col
        vRow = flxGriglia.Row
        Call ColoraColonna
    End If
End Sub

Private Sub flxGriglia_DblClick()
    With flxGriglia
        .SetFocus
        If .Col <> 0 And (.Row = 1 Or .Row = 2 Or .Row = 3) Then
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Width = .ColWidth(.Col)
            If .Row = 1 And .TextMatrix(.Row, .Col) = "" Then
                txtAppo.Text = Day(Now)
            Else
                txtAppo.Text = .TextMatrix(.Row, .Col)
            End If
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

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    intPazientiKey = tTrova.keyReturn
    If tTrova.keyReturn = 0 Then
        Unload frmTsat
    Else
        Call CaricaPaziente
    End If
End Sub

Private Sub cboAnno_Click()
    If stoCaricando Then Exit Sub
    Call Pulisci
    Call CaricaScheda
End Sub

Private Sub txtAppo_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtAppo, lettera)
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If flxGriglia.Row = 8 Then Exit Sub
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii <> vbKeyReturn Then
        If KeyAscii = 44 Then KeyAscii = 46
        lettera = Chr(KeyAscii)
    Else
        txtAppo_Validate (False)
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim k As Integer
    Dim valore As Single
    
    txtAppo.Visible = False
    flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
    Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, vRow)
    cmdAnnulla.Enabled = True
    Call SalvaModifiche
    valore = CalcolaTsat(vCol)
    flxGriglia.TextMatrix(4, vCol) = VirgolaOrPunto(CStr(valore), ",")
    For k = 0 To 1
        grafico(k).Column = 1
        grafico(k).Row = vCol
        grafico(k).data = valore
    Next k

    With flxGriglia
        If (.TextMatrix(1, vCol) = "" And .TextMatrix(2, vCol) = "" And .TextMatrix(3, vCol) = "") Then
            Call Elimina
            .TextMatrix(1, vCol) = ""
            .TextMatrix(2, vCol) = ""
            .TextMatrix(3, vCol) = ""
            .TextMatrix(4, vCol) = ""
        End If
    End With

    txtAppo.MaxLength = 0
End Sub

Private Sub txtAppo_Validate(Cancel As Boolean)
    If txtAppo = "" Then
        Cancel = False
    Else
        If vRow = 1 Then
            Cancel = Not IsDate(txtAppo & "/" & vCol & "/" & cboAnno.Text)
        Else
            Cancel = ControlloNumerico(txtAppo.Text)
        End If
    End If
    If Not Cancel Then
        flxGriglia.SetFocus
    Else
        Beep
    End If
End Sub

