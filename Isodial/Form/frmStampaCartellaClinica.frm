VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStampaCartellaClinica 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa cartella clinica"
   ClientHeight    =   4296
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4296
   ScaleWidth      =   7212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Referti Esami Strumentali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   5415
      End
      Begin VB.ComboBox cboAnno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmStampaCartellaClinica.frx":0000
         Left            =   3720
         List            =   "frmStampaCartellaClinica.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Diario clinico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Accessi Vascolari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Terapia Domiciliare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Terapia Dialitica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Esami di Laboratorio dall'anno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Esami Strumentali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Anamnesi Dialitica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Anamnesi Nefrologica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Anamnesi Patologica e Familiare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox chkSezioni 
         Caption         =   "Anagrafica"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   6975
      Begin VB.CommandButton cmdDeselezionaTutto 
         Caption         =   "&Deseleziona tutto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelezionaTutto 
         Caption         =   "S&eleziona tutto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStampaCartella 
         Caption         =   "&Stampa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStampa 
      Height          =   495
      Left            =   0
      ScaleHeight     =   444
      ScaleWidth      =   324
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmStampaCartellaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim codicePaziente As Integer
Dim codiceId As Integer

Public Property Get getCodicePaziente() As Integer
    getCodicePaziente = codicePaziente
End Property

Public Property Let LetCodicePaziente(ByVal vCodicePaziente As Integer)
    codicePaziente = vCodicePaziente
End Property

Public Property Get getCodiceId() As Integer
    getCodiceId = codiceId
End Property

Public Property Let LetCodiceId(ByVal vCodiceId As Integer)
    codiceId = vCodiceId
End Property

'' Stampa informazioni generali
Private Sub StampaPrimaParte()
    Dim strSqlStampa As String
    Dim strSql As String

    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsPazientiStampa As Recordset          ' tabella pazienti
    Dim rsCodiceEdtaMorte As Recordset
        
    ' carica la stringa
    '--------- paziente
    strSqlStampa = "       NEW adVarChar (50) as INDIRIZZO, " & _
            "       NEW adVarChar (60) as CITTA_NASCITA, " & _
            "       NEW adVarChar (60) as CITTA_RESIDENZA, " & _
            "       NEW adDate AS DATA_NASCITA, " & _
            "       NEW adVarChar (70) as TELEFONO, " & _
            "       NEW adVarChar (5) as DISTRETTO, " & _
            "       NEW adVarChar (20) as CODICE_DOCUMENTO, " & _
            "       NEW adVarChar (2) as TIPO_DOCUMENTO, " & _
            "       NEW adVarChar (35) as CODICE_FISCALE, " & _
            "       NEW adVarChar (20) as ASL, " & _
            "       NEW adVarChar (2) as GRUPPO_SANGUIGNO, " & _
            "       NEW adVarChar (2) as RH, " & _
            "       NEW adVarChar (15) as ESENZIONE, " & _
            "       NEW adInteger as STATO, " & _
            "       NEW adVarChar (5) as NOMEDATA, " & _
            "       NEW adVarChar (12) as DATASTATO, " & _
            "       NEW adVarChar (11) as EDTA_MORTE, " & _
            "       NEW adLongVarChar as NOME_EDTA_MORTE, "
    '--------- medici di base
    strSqlStampa = strSqlStampa & _
            "       NEW adVarChar (50) as COGNOME_MEDICO, " & _
            "       NEW adVarChar (50) as NOME_MEDICO, " & _
            "       NEW adVarChar (20) as TELEFONO_MEDICO "
    
                
    ' stringa di shape
    strSqlStampa = "SHAPE APPEND " & strSqlStampa
     
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    
    ' carica il recordset padre
    Set rsPazientiStampa = New Recordset
    strSql = "SELECT    PAZIENTI.*, COMUNI.NOME AS COMUNINOME, ASL.NOME AS ASLNOME, DISTRETTI.NOME AS DISTRETTINOME, " & _
            "           TIPOLOGIE_ESENZIONE.CODICE AS TIPOLOGIE_ESENZIONECODICE, MEDICI_BASE.NOME AS MEDICI_BASENOME, " & _
            "           MEDICI_BASE.COGNOME AS MEDICI_BASECOGNOME, MEDICI_BASE.TELEFONO AS MEDICI_BASETELEFONO " & _
            " FROM (((((PAZIENTI " & _
            "       LEFT OUTER JOIN ASL ON ASL.KEY=PAZIENTI.CODICE_ASL) " & _
            "       LEFT OUTER JOIN DISTRETTI ON DISTRETTI.KEY=PAZIENTI.CODICE_DISTRETTO) " & _
            "       LEFT OUTER JOIN TIPOLOGIE_ESENZIONE ON TIPOLOGIE_ESENZIONE.KEY=PAZIENTI.CODICE_ESENZIONE) " & _
            "       LEFT OUTER JOIN MEDICI_BASE ON MEDICI_BASE.KEY=PAZIENTI.CODICE_MEDICO) " & _
            "       LEFT OUTER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) " & _
            "       WHERE (PAZIENTI.KEY=" & codicePaziente & ")"
    rsPazientiStampa.Open strSql, cnPrinc, adOpenDynamic, adLockPessimistic, adCmdText
    If Not (rsPazientiStampa.EOF And rsPazientiStampa.BOF) Then
        With rsMain
            .AddNew
            ' pazienti
            .Fields("INDIRIZZO") = rsPazientiStampa("INDIRIZZO")
            .Fields("CITTA_NASCITA") = rsPazientiStampa("CITTA_NASCITA") & " (" & rsPazientiStampa("PROV_NASCITA") & ")"
            .Fields("DATA_NASCITA") = rsPazientiStampa("DATA_NASCITA")
            .Fields("CITTA_RESIDENZA") = rsPazientiStampa("COMUNINOME") & " (" & rsPazientiStampa("PROV_RESIDENZA") & ")"
            .Fields("TELEFONO") = rsPazientiStampa("TELEFONO") & "  " & rsPazientiStampa("CELLULARE")
            .Fields("ASL") = rsPazientiStampa("ASLNOME")
            .Fields("DISTRETTO") = rsPazientiStampa("DISTRETTINOME")
            .Fields("CODICE_DOCUMENTO") = rsPazientiStampa("CODICE_DOCUMENTO")
            .Fields("TIPO_DOCUMENTO") = rsPazientiStampa("TIPO_DOCUMENTO")
            .Fields("CODICE_FISCALE") = rsPazientiStampa("CODICE_FISCALE")
            .Fields("GRUPPO_SANGUIGNO") = rsPazientiStampa("G_SANGUIGNO")
            .Fields("RH") = rsPazientiStampa("RH")
            .Fields("ESENZIONE") = rsPazientiStampa("TIPOLOGIE_ESENZIONECODICE")
            .Fields("STATO") = rsPazientiStampa("STATO")
            .Fields("NOMEDATA") = IIf(rsPazientiStampa("STATO") = 0, "", "Data")
            .Fields("DATASTATO") = rsPazientiStampa("STATODATA") & ""
            
            ' se è deceduto carica il codice edta
            If rsPazientiStampa("STATO") = 1 Then
                If IsNull(rsPazientiStampa("CODICE_EDTA_MORTE")) Then
                    .Fields("EDTA_MORTE") = "Causa morte"
                    .Fields("NOME_EDTA_MORTE") = "- -"
                Else
                    Dim CodiceEdtaMorte As Integer
                    CodiceEdtaMorte = rsPazientiStampa("CODICE_EDTA_MORTE")
                    Set rsCodiceEdtaMorte = New Recordset
                    rsCodiceEdtaMorte.Open "SELECT * FROM EDTA_MORTE WHERE KEY=" & CodiceEdtaMorte, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                    .Fields("EDTA_MORTE") = "Causa morte"
                    .Fields("NOME_EDTA_MORTE") = rsCodiceEdtaMorte("NOME")
                    Set rsCodiceEdtaMorte = Nothing
                End If
            End If
            
            .Fields("COGNOME_MEDICO") = rsPazientiStampa("MEDICI_BASECOGNOME") & ""
            .Fields("NOME_MEDICO") = rsPazientiStampa("MEDICI_BASENOME") & ""
            .Fields("TELEFONO_MEDICO") = rsPazientiStampa("MEDICI_BASETELEFONO") & ""
        End With
    End If

    Set rptCartellaClinica_1 = Nothing

    Set rptCartellaClinica_1.DataSource = rsMain
    rptCartellaClinica_1.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
    rptCartellaClinica_1.PrintReport False, rptRangeAllPages
End Sub

'' Stampa i referti degli esami strumentali scannerizzati
Private Sub StampaReferti()
    Dim rsDataset As New Recordset
    Dim strSql As String
    
    strSql = "SELECT    NOME_FILE " & _
            "FROM       ((PAZIENTI " & _
            "           INNER JOIN ESAMI_STRUMENTALI ON ESAMI_STRUMENTALI.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
            "           INNER JOIN SCAN_ESAMI_STRUMENTALI ON SCAN_ESAMI_STRUMENTALI.CODICE_SCHEDA=ESAMI_STRUMENTALI.KEY) " & _
            "WHERE      PAZIENTI.KEY=" & codicePaziente
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        If Dir(structApri.pathDB & "\" & rsDataset("NOME_FILE") & ".jpg") <> "" Then
            picStampa.Picture = LoadPicture(structApri.pathDB & "\" & rsDataset("NOME_FILE") & ".jpg")
            Printer.ScaleMode = vbPixels
            Printer.PaintPicture picStampa.Picture, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
            Printer.EndDoc
        ElseIf Dir(structApri.pathDB & "\" & rsDataset("NOME_FILE") & ".pdf") <> "" Then
            ShellExecute Me.hWnd, "print", structApri.pathDB & "\" & rsDataset("NOME_FILE") & ".pdf", "", "", 1
            Unload Me
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdDeselezionaTutto_Click()
    Dim i As Integer
    
    For i = 1 To 10
        chkSezioni(i).Value = Unchecked
    Next i
End Sub

Private Sub cmdSelezionaTutto_Click()
    Dim i As Integer
    
    For i = 1 To 10
        chkSezioni(i).Value = Checked
    Next i
End Sub

Private Sub cmdStampaCartella_Click()
    On Error GoTo gestione
    
    Dim quantimesi As Integer
    Dim condizione As String
    Dim data_min As Date
    Dim data_max As Date
    
    quantimesi = 12
    data_min = DateValue("01/01/" & cboAnno.Text)
    data_max = DateValue(Month(date) & "/" & Day(date) & "/" & Year(date))
    condizione = " AND ANAMNESI_ESAMI.DATA BETWEEN #" & data_min & "# AND #" & data_max & "# "
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    If chkSezioni(10).Value = Checked Then
        Call StampaReferti
    End If
    Call StampaPrimaParte
    If chkSezioni(1).Value = Checked Then
        Call StampaSecondaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(2).Value = Checked Then
        Call StampaTerzaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(3).Value = Checked Then
        Call StampaQuartaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(4).Value = Checked Then
        Call StampaQuintaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(5).Value = Checked Then
        Call StampaSestaParte(True, codicePaziente, condizione, quantimesi, codiceId)
    End If
    If chkSezioni(6).Value = Checked Then
        Call StampaSettimaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(7).Value = Checked Then
        Call StampaOttavaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(8).Value = Checked Then
        Call StampaNonaParte(True, codicePaziente, codiceId)
    End If
    If chkSezioni(9).Value = Checked Then
        Call StampaDecimaParte(True, codicePaziente, codiceId)
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 5
        cboAnno.AddItem Year(Now) - i
    Next i
    cboAnno.ListIndex = 0
End Sub
