VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGestioneDocumentiEsterni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione Documenti"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.Label lblTesto 
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   5475
      End
      Begin VB.Label lblTesto 
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5475
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
      Begin VB.CommandButton cmdVisualizza 
         Caption         =   "&Visualizza"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdScansione 
         Caption         =   "&Scansiona"
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
         Left            =   3000
         TabIndex        =   3
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
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdImporta 
         Caption         =   "&Importa"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlApri 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgAppo 
      Height          =   495
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmGestioneDocumentiEsterni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim codicePaziente As Integer
Dim codiceRecord As Integer
Dim codiceCentro As Integer         ' solo per i trapianti
Dim nomeFile As String
Dim strTestoDoc As String

Dim rsDataset As Recordset
Dim numPag As Integer
Dim v_nomeTabella(4) As String

Public Property Get getCodicePaziente() As Integer
    getCodicePaziente = codicePaziente
End Property

Public Property Let LetCodicePaziente(ByVal vCodicePaziente As Integer)
    codicePaziente = vCodicePaziente
End Property

Public Property Get getCodiceRecord() As Integer
    getCodiceRecord = codiceRecord
End Property

Public Property Let letcodiceRecord(ByVal vcodiceRecord As Integer)
    codiceRecord = vcodiceRecord
End Property

Public Property Get getCodiceCentro() As Integer
    getCodiceCentro = codiceCentro
End Property

Public Property Let LetcodiceCentro(ByVal vcodiceCentro As Integer)
    codiceCentro = vcodiceCentro
End Property

Public Property Get getNomeFile() As String
    getNomeFile = nomeFile
End Property

Public Property Let LetNomeFile(ByVal vnomeFile As String)
    nomeFile = vnomeFile
End Property

Private Sub Form_Activate()
    Call Aggiorna
End Sub

Private Sub Form_Load()
    v_nomeTabella(0) = "SCAN_ESAMI_STRUMENTALI"
    v_nomeTabella(1) = "SCAN_PSICO_SOCIALE"
    v_nomeTabella(2) = "SCAN_TRAPIANTI"
    v_nomeTabella(3) = "SCAN_TRATT_ACQUE"
    v_nomeTabella(4) = "SCAN_DOCUMENTI_PAZIENTI"
    
    If tDocumentiEsterni = tpSCANDOCPAZIENTI Then
        strTestoDoc = "documento"
    Else
        strTestoDoc = "referto"
    End If
End Sub

Private Sub Aggiorna()
    Dim condTrapianto As String
    
    Set rsDataset = New Recordset
    
    If tDocumentiEsterni = tpSCANTRAPIANTI Then
        condTrapianto = " AND CODICE_CENTRO=" & codiceCentro
    End If
    
    rsDataset.Open "SELECT * FROM " & v_nomeTabella(tDocumentiEsterni) & " WHERE CODICE_SCHEDA=" & codiceRecord & condTrapianto, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        numPag = rsDataset.RecordCount
        If tDocumentiEsterni = tpSCANTRATTAMENTOACQUE Then
            lblTesto(0) = "Per questa scheda e' presente un " & strTestoDoc & " di " & numPag & " pagin" & IIf(numPag = 1, "a", "e")
        ElseIf tDocumentiEsterni = tpSCANDOCPAZIENTI Then
            lblTesto(0) = "Il documento selezionato ha " & numPag & " pagin" & IIf(numPag = 1, "a", "e")
        Else
            lblTesto(0) = Space(17) & "E' presente un " & strTestoDoc & " di " & numPag & " pagin" & IIf(numPag = 1, "a", "e")
        End If
        lblTesto(1) = "I successivi documenti o immagini saranno accodati"
        cmdVisualizza.Enabled = True
    Else
        If tDocumentiEsterni = tpSCANTRATTAMENTOACQUE Then
            lblTesto(0) = Space(6) & "Per questa scheda NON sono presenti referti"
        ElseIf tDocumentiEsterni = tpSCANDOCPAZIENTI Then
            lblTesto(0) = "Il documento selezionato ha 0 pagine"
        Else
            lblTesto(0) = Space(23) & "NON sono presenti referti"
        End If
        lblTesto(1) = ""
        numPag = 0
        cmdVisualizza.Enabled = False
    End If
    rsDataset.Close
    
    Set rsDataset = Nothing
End Sub

Private Sub EliminaDocSospeso()
    Dim rsDataset As Recordset
    
    ' controlla eventuali scansioni memorizzate in sospeso
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & v_nomeTabella(tDocumentiEsterni) & " WHERE NOME_FILE='" & Apostrophe(nomeFile & "01") & "'", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        rsDataset.Delete
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

Private Sub cmdScansione_Click()
    On Error GoTo gestore
    Dim nome As String
    Dim nPixTypes As Long
    Dim keyScan As Integer

    ' pulisco la clipboar che dovra contenere l'immagine
    Clipboard.Clear
    If TWAIN_AcquireToClipboard(Me.hWnd, nPixTypes) = 0 Then
        MsgBox "Scansione non riuscita", vbInformation, "Impossibile aggiornare"
        Exit Sub
    Else
        nome = nomeFile & Format(numPag + 1, "00")
        imgAppo = Clipboard.getData(2)
        SavePicture imgAppo.Picture, "C:\temp.bmp"
        ' converte in jpg e lo salva sul disco
        Call BmpToJpeg("C:\temp.bmp", structApri.pathDB & "\" & nome & ".jpg", COMPRESSIONE)
        ' elimina il temp
        Kill "C:\temp.bmp"
        ' aggiorna il db
        keyScan = GetNumero(v_nomeTabella(tDocumentiEsterni))
        Set rsDataset = New Recordset
        rsDataset.Open v_nomeTabella(tDocumentiEsterni), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew
        rsDataset("KEY") = keyScan
        If tDocumentiEsterni = tpSCANTRAPIANTI Then
            rsDataset("CODICE_CENTRO") = codiceCentro
        ElseIf tDocumentiEsterni = tpSCANDOCPAZIENTI Then
            rsDataset("CODICE_PAZIENTE") = codicePaziente
            rsDataset("DATA") = date
            If numPag = 0 Then
                frmScanDocumenti.flxGriglia.TextMatrix(frmScanDocumenti.flxGriglia.Row, 0) = keyScan
                'elimina il vecchio record in sospeso
                Call EliminaDocSospeso
                codiceRecord = keyScan
            End If
        End If
        rsDataset("CODICE_SCHEDA") = codiceRecord
        rsDataset("NOME_FILE") = nome
        rsDataset.Update
        Set rsDataset = Nothing
        
        Call Aggiorna

        Unload frmVisualizzaScansione
        Load frmVisualizzaScansione
        frmVisualizzaScansione.LetNomeFile = nome
        frmVisualizzaScansione.letcodiceRecord = codiceRecord
        frmVisualizzaScansione.letNumPag = numPag
        frmVisualizzaScansione.Show 1
    End If
    
    Exit Sub
gestore:
    MsgBox Err.Description, vbCritical, "Attenzione"
End Sub

Private Sub cmdChiudi_Click()
    If tDocumentiEsterni = tpSCANDOCPAZIENTI And numPag = 0 Then
        Call EliminaDocSospeso
        frmScanDocumenti.LetAggiorna = True
    End If
    Unload Me
End Sub

Private Sub cmdVisualizza_Click()
    Dim nome As String
    Dim Visualizza As Boolean
    Dim keyScan As Integer
    Dim numPag As Integer
    Dim condTrapianto As String
    
    If tDocumentiEsterni = tpSCANTRAPIANTI Then
        condTrapianto = " AND CODICE_CENTRO=" & codiceCentro
    End If
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & v_nomeTabella(tDocumentiEsterni) & " WHERE CODICE_SCHEDA=" & codiceRecord & condTrapianto, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        nome = rsDataset("NOME_FILE")
        keyScan = rsDataset("KEY")
        numPag = rsDataset.RecordCount
        Visualizza = True
    Else
        Visualizza = False
    End If
    rsDataset.Close
    
    If Visualizza Then
        Unload frmVisualizzaScansione
        Load frmVisualizzaScansione
        frmVisualizzaScansione.LetNomeFile = nome
        frmVisualizzaScansione.letcodiceRecord = codiceRecord
        frmVisualizzaScansione.LetcodiceCentro = codiceCentro
        frmVisualizzaScansione.letNumPag = numPag
        frmVisualizzaScansione.Show 1
    Else
        MsgBox "Nessun " & strTestoDoc & " da visualizzare", vbCritical, "Attenzione"
    End If
End Sub

Private Sub cmdImporta_Click()
    On Error GoTo gestione

    Dim nomePathFile As String
    Dim nome As String
    Dim keyScan As Integer
    
    Dim img As ImageFile
    Dim pic As ImageFile
    Dim prcs As New ImageProcess
    
    cdlApri.CancelError = True
    cdlApri.Filter = "File immagine jpg|*.jpg|File immagine bmp|*.bmp|File immagine gif|*.gif|File immagine tif|*.tif|File immagine png|*.png|Documento pdf|*.pdf"
    cdlApri.FilterIndex = 1
    cdlApri.ShowOpen
    nomePathFile = cdlApri.FileName
    nome = nomeFile & Format(numPag + 1, "00")
    Select Case Mid(nomePathFile, Len(nomePathFile) - 2, 3)
        Case Is = "jpg"
            FileCopy nomePathFile, structApri.pathDB & "\" & nome & ".jpg"
        Case Is = "pdf"
            FileCopy nomePathFile, structApri.pathDB & "\" & nome & ".pdf"
        Case Else
            While (prcs.Filters.count > 0)
                prcs.Filters.Remove 1
            Wend
            Set img = New ImageFile
            img.LoadFile nomePathFile
                
            prcs.Filters.Add prcs.FilterInfos("Convert").FilterID
            prcs.Filters(1).Properties(1).Value = wiaFormatJPEG
            Set pic = prcs.Apply(img)
            If pic Is Nothing Then
                MsgBox "Impossibile caricare il documento", vbCritical, "Attenzione"
                Exit Sub
            End If
            pic.SaveFile structApri.pathDB & "\" & nome & ".jpg"
    End Select

    ' aggiorna il db
    keyScan = GetNumero(v_nomeTabella(tDocumentiEsterni))
    Set rsDataset = New Recordset
    rsDataset.Open v_nomeTabella(tDocumentiEsterni), cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew
    rsDataset("KEY") = keyScan
    If tDocumentiEsterni = tpSCANTRAPIANTI Then
        rsDataset("CODICE_CENTRO") = codiceCentro
    ElseIf tDocumentiEsterni = tpSCANDOCPAZIENTI Then
        rsDataset("CODICE_PAZIENTE") = codicePaziente
        rsDataset("DATA") = date
        If numPag = 0 Then
            frmScanDocumenti.flxGriglia.TextMatrix(frmScanDocumenti.flxGriglia.Row, 0) = keyScan
            'elimina il vecchio record in sospeso
            Call EliminaDocSospeso
            codiceRecord = keyScan
        End If
    End If
    rsDataset("CODICE_SCHEDA") = codiceRecord
    rsDataset("NOME_FILE") = nome
    rsDataset.Update
    Set rsDataset = Nothing
    
    Call Aggiorna

    Unload frmVisualizzaScansione
    Load frmVisualizzaScansione
    frmVisualizzaScansione.LetNomeFile = nome
    frmVisualizzaScansione.letcodiceRecord = codiceRecord
    frmVisualizzaScansione.letNumPag = numPag
    frmVisualizzaScansione.Show 1
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

