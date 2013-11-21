VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAnamnesiPat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ANAMNESI PATOLOGICA REMOTA E FAMILIARE"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmAnemnesiPat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2160
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
         Left            =   6720
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
         Left            =   11040
         TabIndex        =   17
         Top             =   360
         Width           =   615
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
         TabIndex        =   16
         Top             =   360
         Width           =   1005
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
         TabIndex        =   15
         Top             =   360
         Width           =   630
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
         TabIndex        =   14
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   12015
      Begin VB.TextBox txtFamiliare 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   480
         Width           =   11775
      End
      Begin VB.Label lblUtenteModificatoreFamiliare 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo aggiornamento del dr./dr.ssa:"
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
         Left            =   4500
         TabIndex        =   20
         Top             =   165
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anamnesi familiare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   12015
      Begin VB.CheckBox chkStampa 
         Caption         =   "Stampa Cartella Clinica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3400
         Width           =   2895
      End
      Begin VB.TextBox txtPatologica 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   4440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   7455
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   11
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblUtenteModificatorePat 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo aggiornamento del dr./dr.ssa:"
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
         Left            =   4500
         TabIndex        =   21
         Top             =   165
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anamnesi patologica remota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3435
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   6240
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
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
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
         Left            =   8880
         TabIndex        =   5
         Top             =   240
         Width           =   1455
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
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAnamnesiPat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAnemnesiPat.frm
'
' <b>Descrizione</b>: Scheda Anamnesi Patologica e Familiare associata alla tab ANAMNESI_PAT
'
' @remarks
'
' @author
'
' @date 04/02/2011 19.32
Option Explicit

'' testo da memorizzare nel campo SCHEDA_CLINICA
Dim v_testo(1 To 11) As String
'' true se deve stampare il dato, altrimenti false
Dim v_stampa(1 To 11) As Boolean
'' rs della scheda
Dim rsAnamnesi As Recordset
'' indica se si è in fase di modifica
Dim modifica As Boolean
'' key del record aperto
Dim keyId As Integer
'' rs per la tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer
Dim blnModificato As Boolean
Dim VisualizzaUtenteModificatorePat As Boolean

'' Apre frmTrova se non c'è nessun paziente gia caricato
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    Select Case CaricaPazienteInAperturaForm(Me.Caption, blnModificato, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
    End Select
    
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    modifica = False
    Dim i As Integer
    For i = 1 To 11
        v_stampa(i) = True
    Next i
    Call ApriRsDisconnesso

    With flxGriglia
        .TextMatrix(0, 0) = "Generale"
        .TextMatrix(1, 0) = "Malattie cardiovascolari"
        .TextMatrix(2, 0) = "Malattie polmonari"
        .TextMatrix(3, 0) = "Malattie tubo digerente"
        .TextMatrix(4, 0) = "Malattie endocrino - dismetaboliche"
        .TextMatrix(5, 0) = "Malattie nefro - uro - genitali"
        .TextMatrix(6, 0) = "Malattie infettive"
        .TextMatrix(7, 0) = "Malattie osteo articolari"
        .TextMatrix(8, 0) = "Interventi chirurgici"
        .TextMatrix(9, 0) = "Ricoveri"
        .TextMatrix(10, 0) = "Varie"
        .ColWidth(0) = 4085
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        oPazientiKey.OnClosingForm (Me.Caption)
        intPazientiKey = 0
        blnModificato = False
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim rsDataset As New Recordset
    rsDataset.Open "ANAMNESI_PATOLOGICHE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
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
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "NOME_CAMPI", "VECCHI_VALORI")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rs("KEY"), nome_campi, valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_PATOLOGICHE", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    Dim i As Integer
    modifica = False
    intPazientiKey = 0
    Call PulisciForm(Me)
    chkStampa.Value = Checked
    For i = 1 To 11
        v_stampa(i) = True
        v_testo(i) = ""
        flxGriglia.Row = i - 1
        flxGriglia.CellBackColor = vbWhite
    Next i
    lblUtenteModificatoreFamiliare = "Ultimo aggiornamento del dr./dr.ssa: "
    lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: "
    cmdTrova.SetFocus
    blnModificato = False
End Sub

'' Pulisce solo i campi
Private Sub PulisciCampi()
    Dim i As Integer
    chkStampa.Value = Checked
    txtFamiliare.Text = ""
    txtPatologica.Text = ""
    For i = 1 To 11
        v_stampa(i) = True
        v_testo(i) = ""
        flxGriglia.Row = i - 1
        flxGriglia.CellBackColor = vbWhite
    Next i
    lblUtenteModificatoreFamiliare = "Ultimo aggiornamento del dr./dr.ssa: "
    lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: "
End Sub

Private Sub cmdStampa_Click()
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Sub
    End If
    If Not modifica Then
        MsgBox "La scheda deve essere prima memorizzata", vbCritical, "Attenzione"
        Exit Sub
    End If
      
    Set rsAnamnesi = New Recordset
    rsAnamnesi.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsAnamnesi("COGNOME") & " " & rsAnamnesi("NOME")
    structIntestazione.sDataPaziente = rsAnamnesi("DATA_NASCITA")
    Set rsAnamnesi = Nothing
           
    Call StampaSecondaParte(False, intPazientiKey)
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdMemorizza_Click()
    Dim i As Integer
    Dim v_Val(36) As Variant
    Dim v_Nomi(36) As Variant

    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbCritical, "Attenzione"
        Exit Sub
    End If

    v_Nomi(0) = "KEY"
    v_Nomi(1) = "CODICE_PAZIENTE"
    v_Nomi(2) = "ANAMNESI_FAMILIARE"
    For i = 1 To 11
        v_Nomi(2 + i) = "SCHEDA_CLINICA" & i
        v_Nomi(13 + i) = "STAMPA" & i
        v_Nomi(24 + i) = "UTENTE_MODIFICATORE" & i
    Next i
    v_Nomi(36) = "UTENTE_MODIFICATORE_FAMILIARE"
    
    If Not modifica Then
        keyId = GetNumero("ANAMNESI_PATOLOGICHE")
    End If
    v_Val(0) = keyId
    v_Val(1) = intPazientiKey
    v_Val(2) = txtFamiliare
    For i = 1 To 10
        v_Val(2 + i) = v_testo(i + 1)
        v_Val(13 + i) = v_stampa(i + 1)
        If modifica Then
            If rsDisco.Fields("SCHEDA_CLINICA" & i) <> v_testo(i + 1) Then
                v_Val(24 + i) = tAccesso.key
            Else
                If (rsDisco.Fields("SCHEDA_CLINICA" & i)) = "" Then
                    v_Val(24 + i) = 0
                Else
                    v_Val(24 + i) = rsDisco.Fields(v_Nomi(24 + i))
                End If
            End If
        Else
            If v_testo(i + 1) <> "" Then
                v_Val(24 + i) = tAccesso.key
            End If
        End If
    Next i
    v_Val(13) = v_testo(1)
    v_Val(24) = v_stampa(1)
    If modifica Then
        If rsDisco.Fields("SCHEDA_CLINICA11") <> v_testo(1) Then
            v_Val(35) = tAccesso.key
        Else
            If (rsDisco.Fields("SCHEDA_CLINICA11")) = "" Then
                v_Val(35) = 0
            Else
                v_Val(35) = rsDisco.Fields(v_Nomi(35))
            End If
        End If
        If rsDisco.Fields("ANAMNESI_FAMILIARE") <> txtFamiliare Then
            v_Val(36) = tAccesso.key
        Else
            If (rsDisco.Fields("ANAMNESI_FAMILIARE")) = "" Then
                v_Val(36) = 0
            Else
                v_Val(36) = rsDisco.Fields(v_Nomi(36))
            End If
        End If
    Else
        If v_testo(1) <> "" Then
            v_Val(35) = tAccesso.key
        End If
        If txtFamiliare <> "" Then
            v_Val(36) = tAccesso.key
        End If
    End If
    
    Set rsAnamnesi = New Recordset
    If modifica Then
        rsAnamnesi.Open "SELECT * FROM ANAMNESI_PATOLOGICHE WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        'rsAnamnesi.Update v_nomi, v_val
        For i = 0 To 36
            rsAnamnesi(v_Nomi(i)) = v_Val(i)
        Next
        rsAnamnesi.Update
        If TRACCIATO Then
            Call Confronta(rsAnamnesi)
        End If
    Else
        rsAnamnesi.Open "ANAMNESI_PATOLOGICHE", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsAnamnesi.AddNew v_Nomi, v_Val
        Call Upd_rsDisco
    End If
    Set rsAnamnesi = Nothing
    
    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    blnModificato = False
    modifica = True
End Sub

Private Sub cmdTrova_Click()
    If ControlloChiusuraForm(blnModificato, Me.Caption) Then
        ' pulisce per evitare problemi
        Call PulisciTutto
        tTrova.Tipo = tpPAZIENTE
        tTrova.condizione = ""
        tTrova.condStato = ""
        frmTrova.Show 1
        If tTrova.keyReturn = 0 Then
            Unload Me
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

'' Carica i dati dai vettori nel form
Private Sub flxGriglia_Click()
    Dim i As Integer
    Dim vRow As Integer
    Dim blnModificatoAppo As Boolean
    blnModificatoAppo = blnModificato
    
    txtPatologica = v_testo(flxGriglia.Row + 1)
    chkStampa.Value = IIf(v_stampa(flxGriglia.Row + 1) = True, Checked, Unchecked)
    If flxGriglia.Row = 0 Then
        i = 11
    Else
        i = flxGriglia.Row
    End If
    If rsDisco.RecordCount <> 0 Then
        If Not IsNull(rsDisco.Fields("UTENTE_MODIFICATORE" & i)) Then
            If VisualizzaUtenteModificatorePat = False Then
                lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: "
            Else
                lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsDisco.Fields("UTENTE_MODIFICATORE" & i))
            End If
        Else
            lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: "
            End If
    Else
        lblUtenteModificatorePat = "Ultimo aggiornamento del dr./dr.ssa: "
    End If
    vRow = flxGriglia.Row
    For i = 2 To 11
        If v_testo(i) <> "" Then
            flxGriglia.Row = i - 1
            flxGriglia.CellBackColor = vbGreen
        End If
    Next i
    If v_testo(1) <> "" Then
        flxGriglia.Row = 0
        flxGriglia.CellBackColor = vbGreen
    End If
    flxGriglia.Row = vRow
    Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    
    blnModificato = blnModificatoAppo
End Sub

'' Cambia il valore nel vettore v_stampa
Private Sub chkStampa_Click()
    If flxGriglia.Row = -1 Then
        Exit Sub
    Else
        v_stampa(flxGriglia.Row + 1) = IIf(chkStampa.Value = Checked, True, False)
    End If
    blnModificato = True
End Sub

'' Carica i dati sul paziente e i dati della scheda dalla tabella nel form e nel rs per la tracciatura
Private Sub CaricaPaziente()
    Dim i As Integer
    Dim rsDataset As Recordset
    
    If intPazientiKey = 0 Then Exit Sub

    ' carica i dati del paziente
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
    
    Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
    ' cerca i riferimenti al paziente
    Set rsAnamnesi = New Recordset
    rsAnamnesi.Open "SELECT * FROM ANAMNESI_PATOLOGICHE WHERE CODICE_PAZIENTE=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsAnamnesi.BOF And rsAnamnesi.EOF Then
        ' il paziente non ha una scheda clinica
        modifica = False
        ' pulisco i campi per evitare di caricare i valori ad un paziente senza dati
        ' con l' anamnesi di un paziente caricato precedentemente con i valori
        VisualizzaUtenteModificatorePat = False
        Call PulisciCampi
    Else
        keyId = rsAnamnesi("KEY")
        modifica = True
        VisualizzaUtenteModificatorePat = True
        ' carica tutto nei vattori
        With rsAnamnesi
            For i = 2 To 11
                v_testo(i) = .Fields("SCHEDA_CLINICA" & i - 1)
                If v_testo(i) <> "" Then
                    flxGriglia.Row = i - 1
                    flxGriglia.CellBackColor = vbGreen
                Else
                    flxGriglia.Row = i - 1
                    flxGriglia.CellBackColor = vbWhite
                End If
                v_stampa(i) = CBool(.Fields("STAMPA" & i - 1))
            Next i
            v_testo(1) = .Fields("SCHEDA_CLINICA11")
            If v_testo(1) <> "" Then
                flxGriglia.Row = 0
                flxGriglia.CellBackColor = vbGreen
            Else
                flxGriglia.Row = 0
                flxGriglia.CellBackColor = vbWhite
            End If
            v_stampa(1) = CBool(.Fields("STAMPA11"))
            ' anamnesi familiare
            txtFamiliare = .Fields("ANAMNESI_FAMILIARE")
            If Not IsNull(.Fields("UTENTE_MODIFICATORE_FAMILIARE")) Then
                lblUtenteModificatoreFamiliare = "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(.Fields("UTENTE_MODIFICATORE_FAMILIARE"))
            Else
                lblUtenteModificatoreFamiliare = "Ultimo aggiornamento del dr./dr.ssa: "
            End If
        End With
        
      Call Upd_rsDisco
      
    End If
    Set rsAnamnesi = Nothing
    
    blnModificato = False
End Sub

Private Sub txtFamiliare_Lostfocus()
    txtFamiliare.BackColor = vbWhite
End Sub

Private Sub txtFamiliare_GotFocus()
    txtFamiliare.BackColor = colArancione
End Sub

'' Carica i dati nel vettore v_testo
Private Sub txtPatologica_Change()
    Static fatto As Boolean
    If flxGriglia.Row <> -1 Then
        v_testo(flxGriglia.Row + 1) = txtPatologica.Text
    Else
        fatto = Not fatto
        If Not fatto Then
            MsgBox "Selezionare la patologia", vbCritical, "Attenzione"
        End If
        txtPatologica = ""
    End If
    blnModificato = True
End Sub

Private Sub txtPatologica_GotFocus()
    txtPatologica.BackColor = colArancione
End Sub

Private Sub txtPatologica_LostFocus()
    txtPatologica.BackColor = vbWhite
End Sub

'******** Gestione Modificato
Private Sub txtFamiliare_Change()
    blnModificato = True
End Sub

Private Sub Upd_rsDisco()
 ' aggiorna i dati nel rsDisco
   Dim i As Integer
   Do While Not rsDisco.EOF
      rsDisco.Delete
      rsDisco.MoveNext
      Loop
      rsDisco.AddNew
      For i = 0 To rsAnamnesi.Fields.count - 1
          rsDisco.Fields(i) = rsAnamnesi.Fields(i)
      Next i
      rsDisco.Update
End Sub
