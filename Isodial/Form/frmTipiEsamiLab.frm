VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmTipiEsamiLab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabella Raggruppamento Esami di Laboratorio"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   8280
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
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
         MaxLength       =   45
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   7200
      End
      Begin MSFlexGridLib.MSFlexGrid flxNomi 
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         FormatString    =   $"frmTipiEsamiLab.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTipiEsamiLab.frx":0098
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
      TabIndex        =   4
      Top             =   2640
      Width           =   8055
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
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAnnullaEsami 
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
         Left            =   6480
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserisci 
         Caption         =   "&Nuovo gruppo"
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
         Left            =   4440
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   8055
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.TextBox txtDesrizioneEsame 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2160
         MaxLength       =   45
         TabIndex        =   16
         Top             =   300
         Width           =   4560
      End
      Begin VB.CommandButton cmdSposta 
         Height          =   255
         Index           =   0
         Left            =   7680
         Picture         =   "frmTipiEsamiLab.frx":01F2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdSposta 
         Height          =   255
         Index           =   1
         Left            =   7680
         Picture         =   "frmTipiEsamiLab.frx":033C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   16776960
         ForeColorSel    =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         FormatString    =   $"frmTipiEsamiLab.frx":0486
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTipiEsamiLab.frx":0510
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
      Top             =   6840
      Width           =   8055
      Begin VB.CommandButton cmdEliminaEsame 
         Caption         =   "E&limina"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserisciDescr 
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
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "C&hiudi"
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
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCambiaGruppo 
         Caption         =   "&Cambia Gruppo"
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
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTipiEsamiLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDataset As Recordset
Dim vRow As Integer             ' riga selezionata  (nomi)
Dim vCol As Integer             ' colonna selezionata   (nomi)
Dim vRow2 As Integer            ' griglia
Dim objAnnulla As CAnnulla      ' oggetto che gestisce l'annullamento dei dati nelle flx

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxGriglia
        .Row = 0
        .Col = 1
        .CellFontBold = True
        .Rows = 1
        .ColAlignment(1) = vbLeftJustify
        .ColWidth(0) = 0
        .MousePointer = flexCustom
    End With
    With flxNomi
        .ColWidth(0) = 0
        .Row = 0
        .Col = 1
        .CellFontBold = True
        .ColAlignment(1) = vbLeftJustify
        .MousePointer = flexCustom
    End With
    vRow = 0
    vRow2 = 0
    vCol = 0
    ' carica l'oggetto annulla
    Set objAnnulla = New CAnnulla        ' per gli esami
    ' carica gli esami
    Call CaricaFlx
End Sub

Public Sub CaricaFlxVoci(codiceGruppo As Integer)
    Dim strSql As String
    
    flxGriglia.Rows = 1
    vRow2 = 0
    strSql = "SELECT    VOCI_ESAMI.NOME, VOCI_ESAMI.KEY " & _
            "FROM       ((GRUPPI_ESAMI " & _
            "           INNER JOIN ASSOCIAZIONE_ESAMI_LAB ON ASSOCIAZIONE_ESAMI_LAB.CODICE_GRUPPO=GRUPPI_ESAMI.KEY) " & _
            "           INNER JOIN VOCI_ESAMI ON VOCI_ESAMI.KEY=ASSOCIAZIONE_ESAMI_LAB.CODICE_ESAME) " & _
            "WHERE      CODICE_GRUPPO=" & codiceGruppo & " " & _
            "ORDER BY   ORDINE_VISUALIZZAZIONE"
    Set rsDataset = New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("NOME")
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    Set rsDataset = Nothing
    txtDesrizioneEsame.Text = flxNomi.TextMatrix(vRow, 1)
    flxGriglia.Row = 0
End Sub

Private Sub CaricaFlx()
    flxNomi.Rows = 1
    vCol = 0
    vRow = 0
    ' pulisce l'oggetto
    objAnnulla.Refresh
    cmdAnnullaEsami.Enabled = False
    Set rsDataset = New Recordset
    rsDataset.Open "GRUPPI_ESAMI ORDER BY NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Do While Not rsDataset.EOF
        With flxNomi
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = "" & rsDataset("NOME")
            rsDataset.MoveNext
        End With
    Loop
    rsDataset.Close
    Set rsDataset = Nothing
    flxNomi.Row = 0
End Sub

Private Sub ModificaDati()
    Dim nome As Variant
    Dim valore As Variant
    With flxNomi
        nome = "NOME"
        valore = .TextMatrix(vRow, 1)
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM GRUPPI_ESAMI WHERE KEY=" & .TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsDataset.Update nome, valore
        Set rsDataset = Nothing
        ' pulisce la flx azzerando le righe
        flxGriglia.Rows = 1
        txtDesrizioneEsame.Text = ""
    End With
End Sub

''
' Aggiorna il campo associazione nelle richieste esami di lab
'
' @param codiceAssociazioneVecchia codice della vecchia associazione
' @param codiceGruppoNuovo codice del nuovo gruppo esami
' @param codiceEsame codice dell'esame
Private Sub AggiornaRichiesteEsami(codiceAssociazioneVecchia, codiceGruppoNuovo, codiceEsame)
    Dim rsDataset As New Recordset
    Dim codiceAssociazioneNuova As Integer
    
    rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppoNuovo & " AND CODICE_ESAME=" & codiceEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        codiceAssociazioneNuova = rsDataset("KEY")
    Else
        codiceAssociazioneNuova = 0
    End If
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM RICHIESTE_ESAMI WHERE CODICE_ASSOCIAZIONE=" & codiceAssociazioneVecchia, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    Do While Not rsDataset.EOF
        rsDataset("CODICE_ASSOCIAZIONE") = codiceAssociazioneNuova
        rsDataset.Update
        rsDataset.MoveNext
    Loop
    rsDataset.Close
End Sub

''
' Aggiorna gli esami di lab
'
' @param codiceGruppoVecchio codice del vecchio gruppo esami
' @param codiceGruppoNuovo codice del nuovo gruppo esami
' @param codiceEsame codice dell'esame
Private Sub AggiornaEsamiLab(codiceGruppoVecchio As Integer, codiceGruppoNuovo As Integer, codiceEsame As Integer)
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    Dim cmCommand As New Command
    Dim codiceAnamnesi As Integer
    Dim codiceAnamnesiSostituta As Integer
    Dim keyEsame As Long
    Dim strSql As String
    Dim strSqlInsert As String
    Dim tempo As Single
    
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    
    tempo = Timer
    strSql = "SELECT    ANAMNESI_ESAMI.KEY AS ANAMNESI_ESAMIKEY, ESAMI_LAB.KEY AS ESAMI_LABKEY, DATA, CODICE_PAZIENTE, UTENTE_MODIFICATORE, VALORE " & _
            "FROM       (ANAMNESI_ESAMI " & _
            "           INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) " & _
            "WHERE      CODICE_GRUPPO=" & codiceGruppoVecchio & " AND " & _
            "           CODICE_ESAME=" & codiceEsame
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        codiceAnamnesi = rsDataset("ANAMNESI_ESAMIKEY")
        codiceAnamnesiSostituta = 0
        
        rsAppo.Open "SELECT KEY FROM ANAMNESI_ESAMI WHERE CODICE_GRUPPO=" & codiceGruppoNuovo & " AND DATA=#" & rsDataset("DATA") & "# AND CODICE_PAZIENTE=" & rsDataset("CODICE_PAZIENTE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsAppo.EOF And rsAppo.BOF) Then
            codiceAnamnesiSostituta = rsAppo("KEY")
        End If
        rsAppo.Close
        
        ' cambia l'esame
        If codiceAnamnesiSostituta = 0 Then
            ' bisogna creare anche l'anamnesi
            codiceAnamnesiSostituta = GetNumero("ANAMNESI_ESAMI")
            strSqlInsert = "INSERT INTO ANAMNESI_ESAMI (`KEY`, CODICE_PAZIENTE, DATA, CODICE_GRUPPO, UTENTE_MODIFICATORE) " & _
                            " VALUES (" & codiceAnamnesiSostituta & "," & _
                            rsDataset("CODICE_PAZIENTE") & "," & _
                            "#" & rsDataset("DATA") & "#" & "," & _
                            codiceGruppoNuovo & "," & _
                            rsDataset("UTENTE_MODIFICATORE") & ")"
            cmCommand.CommandText = strSqlInsert
            cmCommand.Execute
        End If
        
        ' crea l'esame nuovo
        keyEsame = GetNumero("ESAMI_LAB")
        strSql = "INSERT INTO ESAMI_LAB (`KEY`, CODICE_ANAMNESI_ESAMI, CODICE_ESAME, VALORE) " & _
                " VALUES (" & keyEsame & "," & _
                codiceAnamnesiSostituta & "," & _
                codiceEsame & "," & _
                VirgolaOrPunto(rsDataset("VALORE"), ",") & ")"
        cmCommand.CommandText = strSql
        cmCommand.Execute
        
        ' elimina l'esame spostato
        cmCommand.CommandText = "DELETE * FROM ESAMI_LAB WHERE KEY=" & rsDataset("ESAMI_LABKEY")
        cmCommand.Execute
        
        ' verifica se l'anamnesi è rimasta priva di esami
        rsAppo.Open "SELECT KEY FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & codiceAnamnesi, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsAppo.EOF And rsAppo.BOF Then
            ' elimina l'anamnesi perche priva di esami
            cmCommand.CommandText = "DELETE * FROM ANAMNESI_ESAMI WHERE KEY=" & codiceAnamnesi
            cmCommand.Execute
        End If
        rsAppo.Close
        
        rsDataset.MoveNext
        frmBarra.prgBar.Value = frmBarra.prgBar + 1
    Loop
    rsDataset.Close
    Debug.Print Timer - tempo
End Sub

''
' Inserisce la nuova associazione e cancella la vecchia
'
' @param codiceGruppoVecchio codice del vecchio gruppo esami
' @param codiceGruppoNuovo codice del nuovo gruppo esami
' @param codiceEsame codice dell'esame
' @return codice della vecchia associazione
Private Function AggiornaAssociazione(codiceGruppoVecchio As Integer, codiceGruppoNuovo As Integer, codiceEsame As Integer) As Integer
    Dim rsDataset As New Recordset
    Dim ordineVisualizzazione As Integer
    Dim codiceVecchiaAssociazione As Integer
    Dim codiceNuovaAssociazione As Integer
    
    rsDataset.Open "SELECT ORDINE_VISUALIZZAZIONE FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppoNuovo & " ORDER BY ORDINE_VISUALIZZAZIONE DESC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        ordineVisualizzazione = rsDataset("ORDINE_VISUALIZZAZIONE") + 1
    Else
        ordineVisualizzazione = 1
    End If
    rsDataset.Close
    
    codiceNuovaAssociazione = GetNumero("ASSOCIAZIONE_ESAMI_LAB")
    rsDataset.Open "ASSOCIAZIONE_ESAMI_LAB", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew
    rsDataset("KEY") = codiceNuovaAssociazione
    rsDataset("CODICE_GRUPPO") = codiceGruppoNuovo
    rsDataset("CODICE_ESAME") = codiceEsame
    rsDataset("ORDINE_VISUALIZZAZIONE") = ordineVisualizzazione
    rsDataset.Update
    rsDataset.Close
    
    rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppoVecchio & " AND CODICE_ESAME=" & codiceEsame, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        codiceVecchiaAssociazione = rsDataset("KEY")
        rsDataset.Delete
    End If
    rsDataset.Close

    AggiornaAssociazione = codiceVecchiaAssociazione
End Function

''
' Verifica che nel gruppo è gia presente l'esame
'
' @param codiceGruppo codice del nuovo gruppo esami
' @param codiceEsame codice dell'esame
' @return true se l'esame è gia presente nel gruppo
Private Function esameGiaPresente(codiceGruppo As Integer, codiceEsame As Integer) As Boolean
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppo & " AND CODICE_ESAME=" & codiceEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.RecordCount = 0 Then
        esameGiaPresente = False
    Else
        esameGiaPresente = True
    End If
End Function

''
' Elimina le richieste esami lab associati al gruppo
Private Sub EliminaRichiesteEsamiLab()
    Dim cmCommand As New Command
    
    cmCommand.CommandType = adCmdText
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandText = "DELETE * FROM RICHIESTE_ESAMI WHERE NOT CODICE_ASSOCIAZIONE IN (SELECT KEY FROM ASSOCIAZIONE_ESAMI_LAB )"
    cmCommand.Execute
End Sub

''
' Elimina le associazioni del gruppo
'
' @param codiceGruppo codice del nuovo gruppo esami
' @param codiceEsame codice dell'esame
Private Sub EliminaAssociazioni(codiceGruppo As Integer, Optional codiceEsame As Integer = 0)
    Dim cmCommand As New Command
    Dim condEsame As String
    
    If codiceEsame <> 0 Then
        condEsame = " AND CODICE_ESAME=" & codiceEsame
    End If
    
    cmCommand.CommandType = adCmdText
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandText = "DELETE * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppo & condEsame
    cmCommand.Execute
End Sub

''
' Elimina il gruppo
'
' @param codiceGruppo codice del nuovo gruppo esami
Private Sub EliminaGruppo(codiceGruppo As Integer)
    Dim cmCommand As New Command
    
    cmCommand.CommandType = adCmdText
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandText = "DELETE * FROM GRUPPI_ESAMI WHERE KEY=" & codiceGruppo
    cmCommand.Execute
End Sub

''
' Verifica se ci sono riferimenti per l'esame selezionato
'
' @param codiceGruppo codice del gruppo di esami
' @param codiceEsame codice dell'esame
' @param esamiPresenti true se ci sono riferimenti in esami di lab
Private Sub VerificaRiferimentiEsamiLab(codiceGruppo As Integer, codiceEsame As Integer, ByRef esamiLabPresenti)
    Dim rsDataset As New Recordset
    Dim condEsame As String
    
    If codiceEsame <> 0 Then
        condEsame = " AND CODICE_ESAME=" & codiceEsame
    End If
        
    rsDataset.Open "SELECT ANAMNESI_ESAMI.KEY FROM (ANAMNESI_ESAMI INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) WHERE CODICE_GRUPPO=" & codiceGruppo & condEsame, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.RecordCount <> 0 Then
        esamiLabPresenti = True
    Else
        esamiLabPresenti = False
    End If
    rsDataset.Close

End Sub

''
' Aggiorna le griglie eliminando i vecchi valori
Private Sub AggiornaGriglie(grigliaGruppi As Boolean)
    If grigliaGruppi Then
        If flxNomi.Rows = 2 Then
            flxNomi.Rows = 1
        Else
            flxNomi.RemoveItem (flxNomi.Row)
        End If
        ' discolora
   '     Call ColoraFlx(flxNomi, flxNomi.Cols - 1, True)
        ' annulla le row e col
        flxNomi.Row = 0
        flxNomi.Col = 0
        ' pulisce la flxgriglia
        flxGriglia.Rows = 1
        vRow = 0
        vRow2 = 0
        txtDesrizioneEsame.Text = ""
    Else
    
        Call CaricaFlxVoci(flxNomi.TextMatrix(flxNomi.Row, 0))
    End If
End Sub

Private Sub AggiornaGriglieDopoInserimento(inCodiceGruppo As Integer)
    flxNomi.Rows = 1
    Call CaricaFlx
    
    ' si posiziona sul record e lo seleziona
    flxNomi.Row = Esiste(flxNomi, 0, vRow, inCodiceGruppo)
    Call ColoraFlx(flxNomi, flxNomi.Cols - 1)
    If flxNomi.Row > 4 Then
        flxNomi.TopRow = flxNomi.Row
    End If
    
    vRow = flxNomi.Row
    Call CaricaFlxVoci(inCodiceGruppo)
End Sub

Private Sub cmdCambiaGruppo_Click()
    'On Error GoTo gestione
    
    Dim esamiGiaPresenti As Boolean
    Dim codiceEsame As Integer
    Dim codiceGruppo As Integer
    Dim codiceVecchiaAssociazione As Integer
    Dim i As Integer
    Dim indiceInf As Integer
    Dim indiceSup As Integer
    Dim intValoreMax As Integer
        
    Dim strElencoCodiceEsami As String
    Dim rsDataset As New Recordset

    cnPrinc.BeginTrans
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare l'esame a cui cambiare gruppo", vbCritical, "Attenzione"
    Else
        codiceGruppo = flxNomi.TextMatrix(flxNomi.Row, 0)
                  
            tSelezionaDaCbo.tipoCampo = tpGRUPPI_ESAMI
            tSelezionaDaCbo.valoreDaEvitare = codiceGruppo
            frmSelezionaDaCbo.Show 1
            
            If tSelezionaDaCbo.valoreSelezionato <> 0 Then
                If flxGriglia.Row > flxGriglia.RowSel Then
                    indiceInf = flxGriglia.RowSel
                    indiceSup = flxGriglia.Row
                Else
                    indiceInf = flxGriglia.Row
                    indiceSup = flxGriglia.RowSel
                End If
                
                Unload frmSelezionaDaCbo
                
                For i = indiceInf To indiceSup
                    strElencoCodiceEsami = strElencoCodiceEsami & flxGriglia.TextMatrix(i, 0) & ","
                Next i
                strElencoCodiceEsami = strElencoCodiceEsami & "0"
                Set rsDataset = New Recordset
                rsDataset.Open "SELECT COUNT(ESAMI_LAB.KEY) AS TOTALE FROM (ANAMNESI_ESAMI INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) WHERE CODICE_GRUPPO=" & codiceGruppo & " AND CODICE_ESAME IN (" & strElencoCodiceEsami & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                intValoreMax = rsDataset("TOTALE") + 1
                rsDataset.Close
                Set rsDataset = Nothing
                
                If intValoreMax > 30 Then
                    Call StartProgressBar(intValoreMax, 0, Me)
                End If
                
                For i = indiceInf To indiceSup
                    codiceEsame = flxGriglia.TextMatrix(i, 0)
                    If Not esameGiaPresente(tSelezionaDaCbo.valoreSelezionato, codiceEsame) Then
                        codiceVecchiaAssociazione = AggiornaAssociazione(codiceGruppo, tSelezionaDaCbo.valoreSelezionato, codiceEsame)
                        Call AggiornaRichiesteEsami(codiceVecchiaAssociazione, tSelezionaDaCbo.valoreSelezionato, codiceEsame)
                        Call AggiornaEsamiLab(codiceGruppo, tSelezionaDaCbo.valoreSelezionato, codiceEsame)
                    Else
                        If flxGriglia.Row = flxGriglia.RowSel Then
                            MsgBox "CAMBIO GRUPPO NON PERMESSO!!!" & vbCrLf & "Esame già presente nel gruppo di destinazione", vbCritical, "Attenzione"
                        Exit Sub
                        Else
                            esamiGiaPresenti = True
                        End If
                    End If
                Next i
                
                Call StopProgressBar(Me)
                                    
                If tSelezionaDaCbo.nuovoInserimento Then
                    Call AggiornaGriglieDopoInserimento(tSelezionaDaCbo.valoreSelezionato)
                Else
                    Call AggiornaGriglie(False)
                End If
                
                If esamiGiaPresenti Then
                    MsgBox "CAMBIO PARZIALE - ESAMI IDENTICI NEL GRUPPO DI DESTINAZIONE!!!" & vbCrLf & "        Non tutti gli esami sono stati trasferiti", vbCritical, "Attenzione"
                Else
                    MsgBox "L'esame è stato trasferito!!!", vbInformation, "Cambio Gruppo"
                End If
            End If
        End If
     cnPrinc.CommitTrans
    Exit Sub
    
gestione:
    cnPrinc.RollbackTrans
End Sub

Private Sub cmdElimina_Click()
    Dim gruppoVuoto As Boolean
    Dim esamiLabPresenti As Boolean
    Dim codiceGruppo As Integer
    Dim rsDataset As New Recordset
    
    If flxNomi.Row = 0 Then
        MsgBox "Selezionare il gruppo da eliminare", vbCritical, "Attenzione"
    Else
        codiceGruppo = flxNomi.TextMatrix(flxNomi.Row, 0)
    
        rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & codiceGruppo, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsDataset.RecordCount <> 0 Then
            gruppoVuoto = False
        Else
            gruppoVuoto = True
        End If
        rsDataset.Close
        
        If Not gruppoVuoto Then
            Call VerificaRiferimentiEsamiLab(codiceGruppo, 0, esamiLabPresenti)
        End If
        
        If gruppoVuoto Then
            If MsgBox("Si conferma l'eliminazione del gruppo: " & flxNomi.TextMatrix(flxNomi.Row, 1) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminazione Gruppo") = vbYes Then
                ' il gruppo puo essere eliminato senza problemi
                Call EliminaGruppo(codiceGruppo)
                Call AggiornaGriglie(True)
                MsgBox "Il Gruppo è stato eliminato!!!", vbInformation, "Elimina Gruppo"
            End If
        Else
            If Not esamiLabPresenti Then
                If MsgBox("Si conferma l'eliminazione?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminazione Gruppo") = vbYes Then
                    
                    Call EliminaGruppo(codiceGruppo)
                    Call EliminaAssociazioni(codiceGruppo)
                    Call EliminaRichiesteEsamiLab
                    
                    Call AggiornaGriglie(True)
                    MsgBox "Il Gruppo è stato eliminato!!!", vbInformation, "Elimina Gruppo"
                End If
            Else
                MsgBox "ELIMINAZIONE GRUPPO NON PERMESSA - Presenza di esami registrati!!!" & vbCrLf & "Trasferire prima gli esami in un altro gruppo e poi procedere con l'eliminazione", vbCritical, "Attenzione"
                Call ColoraFlx(flxNomi, flxNomi.Cols - 1)
            End If
        End If
        
        Set rsDataset = Nothing
    End If
End Sub

Private Sub cmdEliminaEsame_Click()
    Dim esamiLabPresenti As Boolean
    Dim eliminazioneParziale As Boolean
    Dim codiceEsame As Integer
    Dim codiceGruppo As Integer
    Dim i As Integer
    Dim indiceInf As Integer
    Dim indiceSup As Integer
    
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare l'esame da eliminare", vbCritical, "Attenzione"
    Else
        codiceGruppo = flxNomi.TextMatrix(flxNomi.Row, 0)
        
        If flxGriglia.Row > flxGriglia.RowSel Then
            indiceInf = flxGriglia.RowSel
            indiceSup = flxGriglia.Row
        Else
            indiceInf = flxGriglia.Row
            indiceSup = flxGriglia.RowSel
        End If
        
        If MsgBox("Si conferma l'eliminazione " & IIf(flxGriglia.Row = flxGriglia.RowSel, "dell'esame selezionato", "degli esami selezionati") & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminazione Esame") = vbYes Then
            For i = indiceInf To indiceSup
                
                codiceEsame = flxGriglia.TextMatrix(i, 0)
                Call VerificaRiferimentiEsamiLab(codiceGruppo, codiceEsame, esamiLabPresenti)
        
                If Not esamiLabPresenti Then
                    Call EliminaAssociazioni(codiceGruppo, codiceEsame)
                    Call EliminaRichiesteEsamiLab
                    
                Else
                    eliminazioneParziale = True
                End If
            Next i
            
            If indiceInf <> indiceSup Then
                Call AggiornaGriglie(False)
                If eliminazioneParziale Then
                    MsgBox "ELIMINAZIONE PARZIALE - " & vbCrLf & "Alcuni esami hanno valori registrati", vbCritical, "Attenzione"
                Else
                    MsgBox "L'esame e' stato eliminato!!!", vbInformation, "Elimina Esame"
                End If
            Else
                If esamiLabPresenti Then
                    MsgBox "   ELIMINAZIONE NON PERMESSA!!!" & vbCrLf & " L'esame selezionato ha valori registrati", vbCritical, "Attenzione"
                Else
                    Call AggiornaGriglie(False)
                    MsgBox "L'esame e' stato eliminato!!!", vbInformation, "Elimina Esame"
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdElimina_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 And Shift Then
        MsgBox strConnectionStringCentro & " " & strConnectionStringTracciatura
    End If
End Sub

Private Sub cmdAnnullaEsami_Click()
    Dim Dato As String
    Dim Col As Integer
    Dim RowKey As Integer
    Dim i As Integer
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    RowKey = objAnnulla.Row
    ' cerca la riga con il key memorizzato in rowkey
    With flxNomi
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = RowKey Then
                ' annulla
                .TextMatrix(i, Col) = Dato
                objAnnulla.Remove
                ' modifica anche il db
                vRow = i
                Call ModificaDati
                If objAnnulla.Vuoto = True Then
                    cmdAnnullaEsami.Enabled = False
                End If
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdInserisci_Click()
    Dim v_Nomi(1 To 2) As Variant
    Dim v_Val(1 To 2) As Variant
    Dim intCodiceGruppo As Integer
    Dim primo As Boolean
    
    primo = True
    tInput.mantieniDati = False
    tInput.Tipo = tpIESAMI
    Do
        If Not primo Then
            MsgBox "Il gruppo inserito è già presente.", vbCritical, "Attenzione"
            tInput.mantieniDati = True
        End If
        Unload frmInput
        frmInput.Show 1
        primo = False
    Loop While Esiste(flxNomi, 1, 0, tInput.v_valori(1))
    
    If Not (tInput.v_valori(1) = "") Then
        v_Nomi(1) = "KEY"
        v_Nomi(2) = "NOME"
        intCodiceGruppo = GetNumero("GRUPPI_ESAMI")
        v_Val(1) = intCodiceGruppo
        v_Val(2) = tInput.v_valori(1)
        
        Set rsDataset = New Recordset
        rsDataset.Open "GRUPPI_ESAMI", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
        
        Call AggiornaGriglieDopoInserimento(intCodiceGruppo)
                
'        MsgBox "Inserimento effettuato.", vbInformation, "Inserimento"
    End If
End Sub

Private Sub cmdInserisciDescr_Click()
    Dim primo As Boolean
    Dim strSql As String
     
    If vRow = 0 Then
        MsgBox "Selezionare l'esame a cui associare la voce", vbCritical, "Attenzione"
        Exit Sub
    End If
    
    tInput.Tipo = tpITIPIESAMILAB
    tInput.mantieniDati = False
    
    Do
        primo = True
        Unload frmInput
        frmInput.Show 1

        If Esiste(flxGriglia, 0, 0, Int(tInput.v_valori(1))) <> 0 Then
            MsgBox "La voce scelta è stata già inserita", vbCritical, "Attenzione"
            tInput.mantieniDati = True
            primo = False
        Else
            strSql = "SELECT CODICE_ESAME FROM ASSOCIAZIONE_ESAMI_LAB " & _
                     "WHERE CODICE_ESAME= " & tInput.v_valori(1) & " AND CODICE_GRUPPO <> " & flxNomi.TextMatrix(vRow, 0) & ""

            Set rsDataset = New Recordset
            rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            If Not (rsDataset.EOF And rsDataset.BOF) Then
                MsgBox "INSERIMENTO NON PERMESSO - La voce scelta è presente in un altro gruppo", vbCritical, "ATTENZIONE!!!"
                rsDataset.Close
                Set rsDataset = Nothing
                tInput.mantieniDati = True
                primo = False
            End If
        End If
        If primo Then
            Exit Do
        End If
    Loop
    
    If Not (tInput.v_valori(1) = -1) Then
        Set rsDataset = New Recordset
        rsDataset.Open "ASSOCIAZIONE_ESAMI_LAB", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
              
        rsDataset.AddNew
        rsDataset("KEY") = GetNumero("ASSOCIAZIONE_ESAMI_LAB")
        rsDataset("CODICE_GRUPPO") = flxNomi.TextMatrix(vRow, 0)
        rsDataset("CODICE_ESAME") = tInput.v_valori(1)
        rsDataset("ORDINE_VISUALIZZAZIONE") = flxGriglia.Rows
        rsDataset.Update
        rsDataset.Close
        Set rsDataset = Nothing
        
        ' aggiorna la flx
        Call CaricaFlxVoci(flxNomi.TextMatrix(vRow, 0))
        
        ' si posiziona sul record e lo seleziona
        flxGriglia.Row = Esiste(flxGriglia, 0, vRow2, Int(tInput.v_valori(1)))
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
        If flxGriglia.Row > 9 Then
            flxGriglia.TopRow = flxGriglia.Row
        End If
        
 '       MsgBox "Inserimento effettuato", vbInformation, "Inserimento"
    End If
End Sub

Private Sub cmdSposta_Click(Index As Integer)
    Dim num As Integer
    Dim nomeDestinazione As String
    Dim keyDestinazione As Integer
    Dim ordineAppo As Integer
    Dim rsAppo As Recordset
    
    If flxGriglia.Row <> 0 Then
        If Index = 0 Then   ' su
            If flxGriglia.Row <> 1 Then
                num = -1
            Else
                Exit Sub
            End If
        Else                ' giu
            If flxGriglia.Row <> flxGriglia.Rows - 1 Then
                num = 1
            Else
                Exit Sub
            End If
        End If
        ' swap
        keyDestinazione = flxGriglia.TextMatrix(flxGriglia.Row + num, 0)
        nomeDestinazione = flxGriglia.TextMatrix(flxGriglia.Row + num, 1)
        flxGriglia.TextMatrix(flxGriglia.Row + num, 0) = flxGriglia.TextMatrix(flxGriglia.Row, 0)
        flxGriglia.TextMatrix(flxGriglia.Row + num, 1) = flxGriglia.TextMatrix(flxGriglia.Row, 1)
        flxGriglia.TextMatrix(flxGriglia.Row, 0) = keyDestinazione
        flxGriglia.TextMatrix(flxGriglia.Row, 1) = nomeDestinazione
        
        Set rsAppo = New Recordset
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & flxNomi.TextMatrix(vRow, 0) & " AND CODICE_ESAME=" & keyDestinazione, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        rsAppo.Open "SELECT * FROM ASSOCIAZIONE_ESAMI_LAB WHERE CODICE_GRUPPO=" & flxNomi.TextMatrix(vRow, 0) & " AND CODICE_ESAME=" & flxGriglia.TextMatrix(flxGriglia.Row + num, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        ordineAppo = rsAppo("ORDINE_VISUALIZZAZIONE")
        rsAppo("ORDINE_VISUALIZZAZIONE") = rsDataset("ORDINE_VISUALIZZAZIONE")
        rsDataset("ORDINE_VISUALIZZAZIONE") = ordineAppo
        rsDataset.Update
        rsAppo.Update
        Set rsAppo = Nothing
        Set rsDataset = Nothing
        
        flxGriglia.Row = flxGriglia.Row + num
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    ElseIf ControlHWnd = flxNomi.hWnd Then
'        If flxNomi.TopRow - Rotation > 0 Then
'            If flxNomi.TopRow - Rotation < flxNomi.Rows Then
'                flxNomi.TopRow = flxNomi.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'-------------------------------------

Private Sub flxGriglia_Click()
    Dim i As Integer
    Dim numAppo As Integer
    
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        If flxGriglia.Row = flxGriglia.RowSel Then
            ' seleziona singola
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            vRow2 = flxGriglia.Row
        Else
            ' selezione multipla
            numAppo = flxGriglia.RowSel
            For i = 0 To 1
                flxGriglia.Col = i
                flxGriglia.CellBackColor = vbCyan
            Next
            flxGriglia.RowSel = numAppo
            vRow2 = 0
        End If
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
            If i >= 10 Or flxGriglia.TopRow > 10 Then
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

Private Sub flxnomi_Click()
    vCol = flxNomi.Col
    flxNomi.SetFocus
    If VerificaClickFlx(flxNomi) = False Then
        ' discolora
        Call ColoraFlx(flxNomi, flxNomi.Cols - 1, True)
        ' annulla le row e col
        flxNomi.Row = 0
        flxNomi.Col = 0
        ' pulisce la flxgriglia
        flxGriglia.Rows = 1
        vRow = 0
        vRow2 = 0
        txtDesrizioneEsame.Text = ""
    Else
        Call ColoraFlx(flxNomi, flxNomi.Cols - 1)
        flxNomi.Col = vCol
        vRow = flxNomi.Row
        Call CaricaFlxVoci(flxNomi.TextMatrix(vRow, 0))
    End If
End Sub

Private Sub flxnomi_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxNomi.rowPos(flxNomi.Row) + flxNomi.Top + 45
    End If
End Sub

Private Sub flxnomi_DblClick()
    ' fase di modifica
    If VerificaClickFlx(flxNomi) = False Then Exit Sub
    With flxNomi
        .SetFocus
        txtAppo.Left = .colPos(.Col) + .Left + 45
        txtAppo.Top = .rowPos(.Row) + .Top + 45
        txtAppo.Width = .ColWidth(.Col)
        txtAppo.Text = .TextMatrix(.Row, .Col)
        txtAppo.Visible = True
        txtAppo.SetFocus
    End With
End Sub

Private Sub flxNomi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If flxNomi.Rows = 1 Then Exit Sub
    If flxNomi.Row = flxNomi.Rows - 1 Then
        i = 1
    Else
        i = flxNomi.Row + 1
    End If
    Do
        If UCase(Mid(flxNomi.TextMatrix(i, 1), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxNomi.Row = i
            If i > 5 Or flxNomi.TopRow > 5 Then
                flxNomi.TopRow = i
            End If
            Call ColoraFlx(flxNomi, flxNomi.Cols - 1)
            Exit Do
        End If
        If i = flxNomi.Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxNomi.Row
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        flxNomi.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    txtAppo.Visible = False
    If UCase(flxNomi.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
        If txtAppo = "" Then
            MsgBox "MEMORIZZAZIONE CAMPI VUOTI NON PERMESSA", vbCritical, "ATTENZIONE!!!"
            flxNomi.Row = vRow
            Call ColoraFlx(flxNomi, flxNomi.Cols - 1)
            ' essendo l'esame selezionato bisogna mostrare le descrizioni
            Call CaricaFlxVoci(flxNomi.TextMatrix(vRow, 0))
            Exit Sub
        End If
        If Esiste(flxNomi, 1, vRow, txtAppo) Then
            MsgBox "Il nome inserito è già presente", vbCritical, "Attenzione"
        Else
            Call objAnnulla.Add(flxNomi.TextMatrix(vRow, vCol), vCol, flxNomi.TextMatrix(vRow, 0))
            cmdAnnullaEsami.Enabled = True
            flxNomi.TextMatrix(vRow, vCol) = txtAppo.Text
            Call ModificaDati
        End If
    End If
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

