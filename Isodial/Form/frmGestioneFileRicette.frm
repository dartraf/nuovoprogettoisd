VERSION 5.00
Begin VB.Form frmGestioneFileRicette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione File C"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
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
         ItemData        =   "frmGestioneFileRicette.frx":0000
         Left            =   4200
         List            =   "frmGestioneFileRicette.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdScegli 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   3
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtPercorso 
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
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   4575
      End
      Begin VB.ComboBox cboMese 
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
         ItemData        =   "frmGestioneFileRicette.frx":0004
         Left            =   840
         List            =   "frmGestioneFileRicette.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Salva file in"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1230
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
         Index           =   32
         Left            =   3480
         TabIndex        =   7
         Top             =   260
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mese"
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
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   260
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5175
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGenera 
         Caption         =   "&Genera"
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
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmGestioneFileRicette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmGestioneFileRicette.frm
'
' <b>Descrizione</b>: Pannello per la gestione dei file C e XML
'
' @remarks
'
' @author
'
' @date 07/02/2011 17.16
Option Explicit

'' oggetto documentoXML
Dim doc As New DOMDocument60
Dim ret As Boolean

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    If tFileRicette = tpFILEC Then
        Me.Caption = "Gestione File C"
    Else
        Me.Caption = "Gestione File XML"
    End If
    
    cboAnno.AddItem Year(Now)
    cboAnno.AddItem Year(Now) - 1
    cboAnno.ListIndex = 0
    
    cboMese.ListIndex = Month(Now) - 1
    
    txtPercorso = Environ$("USERPROFILE") & "\Desktop"
End Sub

'' Scrive la data secondo lo standard del file XML
'
' @param data data da modificare
Private Function sistemaData(data As Date) As String
    sistemaData = Year(data) & "-" & Format(Month(data), "00") & "-" & Format(Day(data), "00")
End Function

'' Apre un broserforfolder
'
' @param selectedPath path di default
' @return path selezionato
Private Function BrowseForFolder(selectedPath As String) As String
    Dim Browse_for_folder As BROWSEINFOTYPE
    Dim itemID As Long
    Dim selectedPathPointer As Long
    Dim tmpPath As String * 256
    With Browse_for_folder
        .hOwner = Me.hWnd
        .lpszTitle = "Seleziona una cartella:" ' Titolo dialogo
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Funzione di callback per la preselezione cartella
        selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocamento stringa
        CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1
        .Lparam = selectedPathPointer ' Cartella preselezionata
    End With
    itemID = SHBrowseForFolder(Browse_for_folder) ' Apertura finestra di dialogo
    If itemID Then
        If SHGetPathFromIDList(itemID, tmpPath) Then
            BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Elimina gli spazi nulli
        End If
        Call CoTaskMemFree(itemID)
    End If
    Call LocalFree(selectedPathPointer)
End Function

'' Crea un singolo nodo del file XML
'
' @param nome nome del nodo
' @param valore valore da inserire nel nodo
' @return nodo da aggiungere al documento XML
Private Function CreaNodo(nome As String, valore As String) As IXMLDOMNode
    Dim nodo As IXMLDOMNode
    Set nodo = doc.createElement(nome)
    nodo.Text = valore
    Set CreaNodo = nodo
End Function

'' Genera il file XML
'
' @return true se l'operazione è andata a buon fine
Private Function GeneraFileXML() As Boolean
    Dim proc As IXMLDOMProcessingInstruction
    Dim nodo1 As IXMLDOMNode
    Dim nodo2 As IXMLDOMNode
    Dim attr As IXMLDOMAttribute
    Dim root As IXMLDOMElement
    Dim frag As IXMLDOMDocumentFragment
    
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    Dim strSql As String
    Dim codiceSTS As String
    Dim codiceAsl As String
    Dim tipoStampaPC As Integer
    Dim totaleAssistito As Single
    Dim totaleScontato As Single
    Dim ticket As Single
    Dim quotaAggiuntiva As Single
    Dim coefficienteQuotaAggiuntiva As Single
    Dim quotaNazionale As Single
    Dim blnPazienteEstero As Boolean

    
    Set doc = Nothing
    ' versione
    Set proc = doc.createProcessingInstruction("xml", "version='1.0'")
    doc.appendChild proc
    
    ' root
    Set root = doc.createElement("Ricette")
    Set attr = doc.createAttribute("xmlns:xsi")
    attr.Value = "http://www.w3.org/2001/XMLSchema-instance"
    root.setAttributeNode attr
    Set attr = doc.createAttribute("xsi:noNamespaceSchemaLocation")
    attr.Value = "XmlPrestazione2.0.xsd"
    root.setAttributeNode attr
    doc.appendChild root
    
    ' header
    Set frag = doc.createDocumentFragment
    frag.appendChild doc.createTextNode(vbNewLine + vbTab)
    frag.appendChild doc.createElement("Header")
    frag.appendChild doc.createTextNode(vbNewLine + vbTab)
    frag.appendChild doc.createElement("Telematico1")
    frag.appendChild doc.createTextNode(vbNewLine + vbTab)
    frag.appendChild doc.createElement("Telematico2")
    frag.appendChild doc.createTextNode(vbNewLine + vbTab)
    frag.appendChild doc.createElement("Telematico3")
    root.appendChild frag
    
    ' testata
    rsDataset.Open "SELECT ASL.CODICE FROM (INTESTAZIONE_STAMPA LEFT OUTER JOIN ASL ON ASL.KEY=INTESTAZIONE_STAMPA.CODICE_ASL)", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set nodo1 = doc.createElement("Testata")
    nodo1.appendChild CreaNodo("RegStruttura", "150")
    codiceAsl = rsDataset("CODICE")
    nodo1.appendChild CreaNodo("CodAsl", codiceAsl)
    codiceSTS = structIntestazione.sCodiceSTS
    nodo1.appendChild CreaNodo("CodStruttura", codiceSTS)
    rsDataset.Close
    
    ' verifica se ci sono ricette
    rsDataset.Open "SELECT * FROM RICETTE  WHERE (ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        MsgBox "Nessuna ricetta per il mese di " & cboMese.Text, vbInformation, "Genera file XML"
        GeneraFileXML = False
        Exit Function
    Else
        If rsDataset.State = ADODB.adStateOpen Then rsDataset.Close
        rsDataset.Open "SELECT * FROM RICETTE  WHERE (VALIDATA=FALSE AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsDataset.EOF And rsDataset.BOF Then
            MsgBox "NON E' POSSIBILE GENERARE UN NUOVO FILE XML" & vbCrLf & "Per il mese di " & cboMese.Text & " è stato già generato", vbInformation, "Genera file XML"
            GeneraFileXML = False
            Exit Function
        End If
    End If
    nodo1.appendChild CreaNodo("TotRic", rsDataset.RecordCount)
    rsDataset.Close
    
    ' carica ticket e quota aggiuntiva
    rsDataset.Open "SELECT TICKET,QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    'If MsgBox("CIFRARE?", vbYesNo) = vbYes Then
    Call CalcolaCodiciCifrati
    
    ' carica il totale delle prestazioni
    rsAppo.Open "SELECT SUM(QUANTITA) AS TOTALEQ  FROM PRESCRIZIONI WHERE CODICE_RICETTA IN (" & _
                "SELECT KEY FROM RICETTE  WHERE (VALIDATA=FALSE AND NOT FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & "))", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    nodo1.appendChild CreaNodo("TotPrest", rsAppo("TOTALEQ"))
    rsAppo.Close
    
    ' carica il coefficienteQuotaAggiuntiva (1/2 per chi ha codice esenzione ma non esente, intero per chi non ha codice esenzione, E05 non la paga)
    rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALER FROM (RICETTE R INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE)  WHERE ((NOT R.CODICE_ESENZIONE=-1) AND (NOT T.CODICE='E05') AND (NOT ESENZIONE_DOPPIA=TRUE) AND ESENZIONE_QUOTA=FALSE AND VALIDATA=FALSE AND NOT FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    coefficienteQuotaAggiuntiva = rsAppo("TOTALER") / 2
    rsAppo.Close
    ' codice_esenzione=-1 è un record fittizio per fare gli inner join
    rsAppo.Open "SELECT COUNT(KEY) AS TOTALER FROM RICETTE WHERE (CODICE_ESENZIONE=-1 AND VALIDATA=FALSE AND NOT FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    coefficienteQuotaAggiuntiva = coefficienteQuotaAggiuntiva + rsAppo("TOTALER")
    rsAppo.Close
    
    ' calcola il ticket (pagato solo da chi non ha codice_esenzione o da E05) e il totale generale
    rsAppo.Open "SELECT COUNT(R.KEY) AS TOTALER FROM (RICETTE R INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE)  WHERE ((R.CODICE_ESENZIONE=-1 OR T.CODICE='E05') AND VALIDATA=FALSE AND NOT FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    nodo1.appendChild CreaNodo("TotImpCaricoAss", VirgolaOrPunto(Format(rsAppo("TOTALER") * (ticket + quotaNazionale) + coefficienteQuotaAggiuntiva * quotaAggiuntiva, "#######0.00"), ","))
    totaleAssistito = Format(rsAppo("TOTALER") * (ticket + quotaNazionale) + coefficienteQuotaAggiuntiva * quotaAggiuntiva, "#######0.00")
    rsAppo.Close
    
    ' calcola i totali generali
    rsAppo.Open "SELECT SUM(QUANTITA*IMPORTO) AS TOTALE, SUM(QUANTITA*IMPORTO_SCONTATO) AS TOTALE_SCONTATO FROM PRESCRIZIONI WHERE CODICE_RICETTA IN (" & _
                "SELECT KEY FROM RICETTE  WHERE (VALIDATA=FALSE AND NOT FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & "))", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    nodo1.appendChild CreaNodo("TotValRicInviate", VirgolaOrPunto(Format(rsAppo("TOTALE"), "#####.00"), ","))
    nodo1.appendChild CreaNodo("TotImpCaricoSSN", VirgolaOrPunto(Format(rsAppo("TOTALE_SCONTATO") - totaleAssistito, "#####.00"), ","))
    rsAppo.Close
    
    ' calcola i dati sul numero ricette I, V, C
    rsAppo.Open "SELECT COUNT(KEY) AS TOTALE_INS FROM RICETTE WHERE (VALIDATA=FALSE AND FLAG=1 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")"
    nodo1.appendChild CreaNodo("TotRicNuove", rsAppo("TOTALE_INS"))
    rsAppo.Close
    rsAppo.Open "SELECT COUNT(KEY) AS TOTALE_VAR FROM RICETTE WHERE (VALIDATA=FALSE AND FLAG=2 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")"
    nodo1.appendChild CreaNodo("TotRicVariaz", rsAppo("TOTALE_VAR"))
    rsAppo.Close
    rsAppo.Open "SELECT COUNT(KEY) AS TOTALE_CAN FROM RICETTE WHERE (VALIDATA=FALSE AND FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")"
    nodo1.appendChild CreaNodo("TotRicCanc", rsAppo("TOTALE_CAN"))
    rsAppo.Close
    nodo1.appendChild CreaNodo("TotStrutture", "1")
    nodo1.appendChild CreaNodo("AnnoMeseNoInvio", "")
    root.appendChild nodo1

    frmBarra.prgBar.Value = frmBarra.prgBar.Value + Int((frmBarra.prgBar.max - frmBarra.prgBar.Value) / 2)
    
    strSql = "SELECT    RICETTE.*, " & _
            "           TIPOLOGIE_ESENZIONE.KEY AS TIPOLOGIE_ESENZIONEKEY, TIPOLOGIE_ESENZIONE.CODICE AS TIPOLOGIE_ESENZIONECODICE, " & _
            "           TIPI_EROGAZIONE.CODICE AS TIPI_EROGAZIONECODICE, Nazioni.Nome as NazioniNome, Nazioni.CodiceAlfa2 as NazioniCodiceAlfa2, TipiRicetta.Codice as TipiRicettaCodice, " & _
            "           PAZIENTI.CODICE_FISCALE_CIFRATO, PAZIENTI.DATA_NASCITA " & _
            "FROM       (((((RICETTE " & _
            "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
            "           INNER JOIN Nazioni ON Nazioni.KEY=Pazienti.NazioniID) " & _
            "           INNER JOIN TipiRicetta ON TipiRicetta.KEY=RICETTE.TipiRicettaID) " & _
            "           LEFT OUTER JOIN TIPOLOGIE_ESENZIONE ON TIPOLOGIE_ESENZIONE.KEY=RICETTE.CODICE_ESENZIONE) " & _
            "           INNER JOIN TIPI_EROGAZIONE ON TIPI_EROGAZIONE.KEY=RICETTE.CODICE_TIPO_EROGAZIONE) " & _
            "WHERE      (VALIDATA=FALSE AND " & _
            "           ANNO=" & cboAnno.Text & " AND " & _
            "           MESE=" & cboMese.ListIndex + 1 & " AND " & _
            "           NAZIONI.UnioneEuropea<>0) " & _
            "ORDER BY   FLAG DESC, PROGRESSIVO_ANNUALE ASC"
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        If UCase(rsDataset("NazioniNome")) = UCase("Italia") Then
            blnPazienteEstero = False
        Else
            blnPazienteEstero = True
        End If
        ' esplode ricetta per ricetta
        Set nodo1 = doc.createElement("Ricetta")
        nodo1.appendChild CreaNodo("FlagOperazione", Choose(rsDataset("FLAG"), "I", "V", "C"))
        nodo1.appendChild CreaNodo("RegStrutturaRic", "150")
        nodo1.appendChild CreaNodo("CodAslRic", codiceAsl)
        nodo1.appendChild CreaNodo("CodStrutturaRic", codiceSTS)
        nodo1.appendChild CreaNodo("CodRegione", Mid(rsDataset("NUMERO_RICETTA"), 1, 3))
        nodo1.appendChild CreaNodo("AnnoProduzione", Mid(rsDataset("NUMERO_RICETTA"), 4, 2))
        nodo1.appendChild CreaNodo("ProgRicettaRicettario", Mid(rsDataset("NUMERO_RICETTA"), 6, 9))
        nodo1.appendChild CreaNodo("CheckDigit", Mid(rsDataset("NUMERO_RICETTA"), 15, 1))
        nodo1.appendChild CreaNodo("CodiceAss", rsDataset("CODICE_FISCALE_CIFRATO"))
        nodo1.appendChild CreaNodo("ProgRicettaStruttura", "")
        nodo1.appendChild CreaNodo("SiglaProvincia", "")
        nodo1.appendChild CreaNodo("ASLAssistito", "")
        nodo1.appendChild CreaNodo("DispReg", "")
        nodo1.appendChild CreaNodo("Suggerita", IIf(rsDataset("TIPOLOGIA_RICETTA") = 1, "S", ""))
        nodo1.appendChild CreaNodo("Altro", "")
        nodo1.appendChild CreaNodo("AltroRic", IIf(rsDataset("TIPOLOGIA_RICETTA") = 3, "1", ""))
        nodo1.appendChild CreaNodo("DataCompilazione", "")
        nodo1.appendChild CreaNodo("DataSpedizione", "")
        nodo1.appendChild CreaNodo("TipoAccesso", "")
        nodo1.appendChild CreaNodo("GaranziaTempiMassimi", "")
        nodo1.appendChild CreaNodo("AnnoMeseFatt", "")
        nodo1.appendChild CreaNodo("TipoRic", IIf(blnPazienteEstero, rsDataset("TipiRicettaCodice"), ""))
        nodo1.appendChild CreaNodo("NonEsente", IIf(rsDataset("TIPOLOGIE_ESENZIONEKEY") = -1, 1, ""))
        If rsDataset("TIPOLOGIE_ESENZIONEKEY") = -1 Then
            nodo1.appendChild CreaNodo("CodEsenzione", "")
        Else
            nodo1.appendChild CreaNodo("CodEsenzione", rsDataset("TIPOLOGIE_ESENZIONECODICE"))
        End If
  '      nodo1.appendChild CreaNodo("Reddito", IIf(CBool(rsDataset("ESENTE_REDDITO")), "1", ""))
        nodo1.appendChild CreaNodo("Reddito", IIf(CBool(rsDataset("ESENTE_REDDITO")), "", ""))
        If CBool(rsDataset("STAMPATO_PC")) Then
            If CBool(rsDataset("PRESENZA_BARCODE")) Then
                tipoStampaPC = 1
            Else
                tipoStampaPC = 2
            End If
        Else
            tipoStampaPC = 0
        End If
        nodo1.appendChild CreaNodo("CodRaggrup", CStr(tipoStampaPC))
        nodo1.appendChild CreaNodo("ClassePriorita", "")
        nodo1.appendChild CreaNodo("TipoErogazione", rsDataset("TIPI_EROGAZIONECODICE"))
        nodo1.appendChild CreaNodo("CodiceDiagnosi", "")
        
        ' carica i totali prestazione, totale ricetta
        rsAppo.Open "SELECT COUNT(KEY) AS TOTALER, SUM(QUANTITA) AS TOTALEQ, SUM(QUANTITA*IMPORTO) AS TOTALE, SUM(QUANTITA*IMPORTO_SCONTATO) AS TOTALE_SCONTATO FROM PRESCRIZIONI WHERE CODICE_RICETTA=" & rsDataset("KEY"), cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        nodo1.appendChild CreaNodo("TotPrestazioni", rsAppo("TOTALEQ"))
        nodo1.appendChild CreaNodo("TotValoreRicetta", VirgolaOrPunto(Format(rsAppo("TOTALE"), "#####.00"), ","))
        totaleScontato = Format(rsAppo("TOTALE_SCONTATO"), "#####.00")
        rsAppo.Close
                
        ' imposta il coefficienteQuotaAggiuntiva per calcolare l'importo dell assistito
        rsAppo.Open "SELECT R.CODICE_ESENZIONE, ESENZIONE_QUOTA, ESENZIONE_DOPPIA, CODICE  FROM (RICETTE R INNER JOIN TIPOLOGIE_ESENZIONE T ON T.KEY=R.CODICE_ESENZIONE)  WHERE R.KEY=" & rsDataset("KEY"), cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        If CBool(rsAppo("ESENZIONE_DOPPIA")) Then
            coefficienteQuotaAggiuntiva = 0
        Else
            If rsAppo("CODICE_ESENZIONE") = -1 Then
                coefficienteQuotaAggiuntiva = 1
            ElseIf CBool(rsAppo("ESENZIONE_QUOTA")) Or UCase(rsAppo("CODICE")) = "E05" Then
                coefficienteQuotaAggiuntiva = 0
            Else
                coefficienteQuotaAggiuntiva = 1 / 2
            End If
        End If
        nodo1.appendChild CreaNodo("FranchigiaCaricoAss", VirgolaOrPunto(Format(IIf(rsAppo("CODICE_ESENZIONE") = -1 Or UCase(rsAppo("CODICE")) = "E05", ticket, 0), "#######0.00"), ","))
        nodo1.appendChild CreaNodo("QuotaCaricoAss", VirgolaOrPunto(Format(IIf(rsAppo("CODICE_ESENZIONE") = -1 Or UCase(rsAppo("CODICE")) = "E05", quotaNazionale, 0) + coefficienteQuotaAggiuntiva * quotaAggiuntiva, "#######0.00"), ","))
        totaleScontato = totaleScontato - Format(IIf(rsAppo("CODICE_ESENZIONE") = -1 Or UCase(rsAppo("CODICE")) = "E05", (ticket + quotaNazionale), 0) + coefficienteQuotaAggiuntiva * quotaAggiuntiva, "#######0.00")
        rsAppo.Close
        
        nodo1.appendChild CreaNodo("ImpCaricoSSN", VirgolaOrPunto(CStr(totaleScontato), ","))
        nodo1.appendChild CreaNodo("StatoEstero", IIf(blnPazienteEstero, rsDataset("NazioniCodiceAlfa2"), ""))
        nodo1.appendChild CreaNodo("IstituzCompetente", rsDataset("CodiceIstituzioneCompetente") & "")
        nodo1.appendChild CreaNodo("NumIdentPers", rsDataset("NumeroIdentificativoPersonale") & "")
        nodo1.appendChild CreaNodo("NumIdentTess", rsDataset("NumeroIdentificazioneTessera") & "")
        nodo1.appendChild CreaNodo("DataNascitaEstero", IIf(blnPazienteEstero, rsDataset("Data_Nascita") & "", ""))
        nodo1.appendChild CreaNodo("DataScadTessera", rsDataset("DataScadenzaTessera") & "")
        
        strSql = "SELECT    PRESCRIZIONI.*, NOMENCLATORE_TARIFFARIO.NOME , NOMENCLATORE_TARIFFARIO.CODICE " & _
                 "FROM      (PRESCRIZIONI " & _
                 "          INNER JOIN NOMENCLATORE_TARIFFARIO ON NOMENCLATORE_TARIFFARIO.KEY=PRESCRIZIONI.CODICE_PRESTAZIONE) " & _
                 "WHERE     CODICE_RICETTA=" & rsDataset("KEY")
        rsAppo.Open strSql, cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        Do While Not rsAppo.EOF
            ' prescrizione per prescrizione
            Set nodo2 = doc.createElement("Prestazione")
            nodo2.appendChild CreaNodo("CodicePresidio", "")
            nodo2.appendChild CreaNodo("CodicePrest", rsAppo("CODICE"))
            nodo2.appendChild CreaNodo("CodReparto", "")
            nodo2.appendChild CreaNodo("BrancaPrestazione", "13")
            nodo2.appendChild CreaNodo("DataPrenotazione", sistemaData(rsDataset("DATA_PRENOTAZIONE")))
            nodo2.appendChild CreaNodo("DataErogInizio", IIf(rsAppo("QUANTITA") = 1, "", sistemaData(rsAppo("DATA_INIZIO"))))
            nodo2.appendChild CreaNodo("DataErogFine", IIf(rsAppo("QUANTITA") = 1, "", sistemaData(rsAppo("DATA_FINE"))))
            nodo2.appendChild CreaNodo("DataErogazione", sistemaData(rsAppo("DATA_INIZIO")))
            nodo2.appendChild CreaNodo("TipologiaPrestazione", "")
            nodo2.appendChild CreaNodo("QtaPrest", rsAppo("QUANTITA"))
'            nodo2.appendChild CreaNodo("TariffaPrest", VirgolaOrPunto(Format(rsAppo("QUANTITA") * rsAppo("IMPORTO"), "#####.00"), ","))
'            nodo2.appendChild CreaNodo("TariffaPrestLab", VirgolaOrPunto(Format(rsAppo("QUANTITA") * rsAppo("IMPORTO_SCONTATO"), "#####.00"), ","))
            nodo2.appendChild CreaNodo("TariffaPrest", VirgolaOrPunto(Format(rsAppo("IMPORTO"), "#####.00"), ","))
            nodo2.appendChild CreaNodo("TariffaPrestLab", VirgolaOrPunto(Format(rsAppo("IMPORTO_SCONTATO"), "#####.00"), ","))
            
            nodo1.appendChild nodo2
            rsAppo.MoveNext
        Loop
        rsAppo.Close
        root.appendChild nodo1
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    doc.Save txtPercorso & "\" & structIntestazione.sCodiceSTS & "_" & Format(cboMese.ListIndex + 1, "00") & cboAnno.Text & Format(date, "ddmmyyyy") & ".xml"
    rsDataset.Close
    
    GeneraFileXML = True
    MsgBox "File XML generato correttamente", vbInformation, "Gestione file XML"
End Function

'' Genera i file C
'
' @remarks la creazione del file è gestita dalla classe CFileC
Private Sub GeneraFileC()
    Dim fileC As CFileC
    
    Set fileC = New CFileC
    fileC.anno = cboAnno.Text
    fileC.mese = cboMese.ListIndex + 1
    ret = fileC.CreaFiles(txtPercorso)
    Set fileC = Nothing

End Sub

'' Calcola i codici fiscali cifrati dei pazienti
Private Sub CalcolaCodiciCifrati()
    On Error GoTo gestione
    Dim rsDataset As New Recordset
    Dim ret As Long
    Dim codice As String
    Dim strSql As String
    
    strSql = "SELECT    DISTINCT PAZIENTI.KEY, CODICE_FISCALE, CODICE_FISCALE_CIFRATO " & _
            "FROM       (PAZIENTI INNER JOIN RICETTE ON RICETTE.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
            "WHERE      VALIDATA=FALSE AND " & _
            "           ANNO=" & cboAnno.Text & " AND " & _
            "           MESE=" & cboMese.ListIndex + 1
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    Call StartProgressBar(Int(4 / 3 * rsDataset.RecordCount + 1), 0, Me)

    Do While Not rsDataset.EOF
        frmBarra.prgBar.Value = frmBarra.prgBar.Value + 1
        Open structApri.pathExe & "\cf_chiaro.txt" For Output As #1
        Print #1, rsDataset("CODICE_FISCALE")
        Close #1
        'structApri.pathNomeCertificato = "C:\Programmi\Isodial\MEFpp.cer"
        ret = ExecCmd(structApri.pathExe & "\openssl.exe rsautl -encrypt -in " & structApri.pathExe & "\cf_chiaro.txt -out " & structApri.pathExe & "\cf_chiaro.enc -inkey " & structApri.pathNomeCertificato & " -certin -pkcs")
        ret = ExecCmd(structApri.pathExe & "\openssl.exe base64 -base64 -e  -out " & structApri.pathExe & "\cf_chiaro.enc.b64 -in " & structApri.pathExe & "\cf_chiaro.enc")

        Open structApri.pathExe & "\cf_chiaro.enc.b64" For Input As #1
        Input #1, codice
        Close #1
        
        rsDataset("CODICE_FISCALE_CIFRATO") = codice
        rsDataset.MoveNext
    Loop
    
    Kill structApri.pathExe & "\cf_chiaro.txt"
    Kill structApri.pathExe & "\cf_chiaro.enc"
    Kill structApri.pathExe & "\cf_chiaro.enc.b64"
    Exit Sub
gestione:
    MsgBox "Errore nella cifratura dei codici fiscali" & vbCrLf & "Impossibile generare il file", vbCritical, "Attenzione"
    Unload Me
End Sub

'' Elimina le ricette che sono cancellate "C" e le relative prescrizioni
Private Sub EliminaCancellate()
    Dim cmCommand As New Command
    
    cmCommand.ActiveConnection = cnPrinc
    cmCommand.CommandType = adCmdText
    
    cmCommand.CommandText = "DELETE * FROM PRESCRIZIONI WHERE CODICE_RICETTA IN (SELECT KEY FROM RICETTE WHERE (FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & "))"
    cmCommand.Execute
    
    cmCommand.CommandText = "DELETE * FROM RICETTE WHERE (FLAG=3 AND ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & ")"
    cmCommand.Execute
End Sub

'' Valida tutte le ricette del mese e dell'anno selezionato
Private Sub ValidaRicette()
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM RICETTE WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1, cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsDataset.EOF
        rsDataset("VALIDATA") = True
        rsDataset.MoveNext
    Loop
    rsDataset.Close
End Sub

Private Sub cmdGenera_Click()
    If structIntestazione.sCodiceSTS = CODICESTS_MOSCATI Or structIntestazione.sCodiceSTS = CODICESTS_GAMMADIAL Or structIntestazione.sCodiceSTS = CODICESTS_CGA Or structIntestazione.sCodiceSTS = CODICESTS_DIALIFE Or structIntestazione.sCodiceSTS = CODICESTS_CAMPANO Or structIntestazione.sCodiceSTS = CODICESTS_DIALGEST Or structIntestazione.sCodiceSTS = CODICESTS_SBIAGIO Or structIntestazione.sCodiceSTS = CODICESTS_NEPHRON Or structIntestazione.sCodiceSTS = CODICESTS_DELTA Or structIntestazione.sCodiceSTS = CODICESTS_POGGIOMARINO Or structIntestazione.sCodiceSTS = CODICESTS_EM_IRPINA Or structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Or structIntestazione.sCodiceSTS = CODICESTS_LA_PECCERELLA Or structIntestazione.sCodiceSTS = CODICESTS_SANNIOMEDICA Or structIntestazione.sCodiceSTS = CODICESTS_SANT_ANDREA Or structIntestazione.sCodiceSTS = CODICESTS_SODAV Or structIntestazione.sCodiceSTS = CODICESTS_HELIOS Then
    Else
        MsgBox "MODULO PER GENERARE I FILE 'C' OPZIONALE ATTIVABILE A RICHIESTA", vbInformation, "INFORMAZIONE"
        Exit Sub
    End If
    
    Dim testo As String
    
    If tFileRicette = tpFILEC Then
        Call GeneraFileC
        If ret Then
            Unload Me
        End If
    Else
        If GeneraFileXML Then
            If SuccessivoInvio Then
                testo = " L'INVIO "
            Else
                testo = " IL PRIMO INVIO "
            End If
            If MsgBox("SI CONFERMA IL FILE XML GENERATO PER" & testo & "AL MEF?", vbQuestion + vbYesNo + vbDefaultButton2, "Genera XML") = vbYes Then
                Call ValidaRicette
                Call EliminaCancellate
            End If
            Call StopProgressBar(Me)
            Unload Me
        Else
            Call StopProgressBar(Me)
        End If
    End If
End Sub


'' Verifica se c'è gia stato un invio per quel mese
Private Function SuccessivoInvio() As Boolean
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT VALIDATA, MESE, ANNO FROM RICETTE WHERE ANNO=" & cboAnno.Text & " AND MESE=" & cboMese.ListIndex + 1 & " AND VALIDATA = True", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    SuccessivoInvio = rsDataset.RecordCount
    rsDataset.Close
End Function

Private Sub cmdScegli_Click()
    Dim tmpPath As String
    tmpPath = txtPercorso
    If Len(tmpPath) > 0 Then
        If Not Right$(tmpPath, 1) <> "\" Then tmpPath = Left$(tmpPath, Len(tmpPath) - 1) ' Remove "\" if the user added
    Else
        tmpPath = ""
    End If
    If Right(tmpPath, 1) = ":" Then tmpPath = tmpPath & "\"
    txtPercorso = tmpPath
    tmpPath = BrowseForFolder(tmpPath)
    If tmpPath <> "" Then
        txtPercorso = tmpPath
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

