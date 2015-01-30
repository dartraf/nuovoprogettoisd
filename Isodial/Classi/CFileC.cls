VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo di classe - CFileC.cls
'
' <b>Descrizione</b>: Classe per la gestione dei file C
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.30
Option Explicit

Public mese As Integer
Public anno As Integer

Private mNodo() As CNodoFileC
Private mIntestazioneC1 As String
Private mIntestazioneC2 As String
Private mSTS As String

''
' Inizializza la pila
'
' @param
' @param
' @return
' @remarks
Private Sub Class_Initialize()
    ReDim mNodo(0)
    Set mNodo(0) = New CNodoFileC
End Sub

''
' Termina la pila
'
' @param
' @param
' @return
' @remarks
Private Sub Class_Terminate()
    ReDim mNodo(0)
    Set mNodo(0) = Nothing
End Sub

''
' Sistema le stringhe nel formato richiesto per i file C
'
' @param testo testo da analizzare
' @param lenTotale lunghezza richiesta
' @return testo formattato
' @remarks aggiunge spazi finali
Private Function CompletaTesto(testo As String, lenTotale As Integer) As String
    If Len(testo) < lenTotale Then
        CompletaTesto = testo & Space(lenTotale - Len(testo))
    Else
        CompletaTesto = testo
    End If
End Function

''
' Sistema i numeri nel formato richiesto per i file C
'
' @param numero numero da analizzare
' @param lenTotale lunghezza richiesta
' @return numero formattato
' @remarks aggiunge 0 iniziali
Private Function CompletaNumeri(numero As String, lenTotale As Integer) As String
    If Len(numero) < lenTotale Then
        CompletaNumeri = String(lenTotale - Len(numero), "0") & numero
    Else
        CompletaNumeri = numero
    End If
End Function

''
' Sistema la data nel formato richiesto per i file C
'
' @param data data da analizzare
' @param
' @return data formattata
' @remarks
Private Function sistemaData(data As Date) As String
    sistemaData = Format(Day(data), "00") & Format(Month(data), "00") & Year(data)
End Function

''
' Crea le intestazioni dei file C
'
' @param
' @param
' @return
' @remarks stringa identica in ogni riga dei file C
Private Sub CreaIntestazioni()
    Dim strSql As String
    strSql = "SELECT    ASL.CODICE AS ASLCODICE, DISTRETTI.CODICE AS DISTRETTICODICE " & _
            "FROM       (INTESTAZIONE_STAMPA " & _
            "           LEFT OUTER JOIN ASL ON ASL.KEY=INTESTAZIONE_STAMPA.CODICE_ASL) " & _
            "           LEFT OUTER JOIN DISTRETTI ON DISTRETTI.KEY=INTESTAZIONE_STAMPA.CODICE_DISTRETTO"
    Dim rsDataset As New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    mSTS = structIntestazione.sCodiceSTS
    mIntestazioneC1 = rsDataset("ASLCODICE") & CompletaTesto(structIntestazione.sCodiceSTS, 6) & Space(2) & "2" & CompletaNumeri(rsDataset("DISTRETTICODICE"), 5)
    mIntestazioneC2 = rsDataset("ASLCODICE") & CompletaTesto(structIntestazione.sCodiceSTS, 6) & "013    2" & CompletaNumeri(rsDataset("DISTRETTICODICE"), 5)
    Set rsDataset = Nothing
End Sub

''
' Crea i file come insieme di nodi
'
' @param percorso path nel quale salvare i file C
' @param
' @return
' @remarks
Public Function CreaFiles(percorso As String) As Boolean
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    
    Dim i As Integer
    Dim k As Integer
    
    Dim NuovaCartella As String
    Dim cronologiaPrescrizione As Integer
    Dim importoTotale As Single
    Dim ticket As Single
    Dim quotaAggiuntiva As Single
    Dim coefficienteQuotaAggiuntiva As Single
    Dim quotaNazionale As Single
    Dim blnPazienteEstero As Boolean
    
    Dim strSql As String
    
    ' crea l'intestazione uguale ad entrambi i files
    Call CreaIntestazioni
    
    rsDataset.Open "SELECT TICKET, QUOTA_AGGIUNTIVA, QUOTA_NAZIONALE FROM INTESTAZIONE_FATTURA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    ticket = VirgolaOrPunto(rsDataset("TICKET"), ".")
    quotaAggiuntiva = VirgolaOrPunto(rsDataset("QUOTA_AGGIUNTIVA"), ".")
    quotaNazionale = VirgolaOrPunto(rsDataset("QUOTA_NAZIONALE"), ".")
    rsDataset.Close
    
    i = 0
    strSql = "SELECT    RICETTE.*,  " & _
            "           PAZIENTI.NOME AS PAZIENTINOME, PAZIENTI.COGNOME AS PAZIENTICOGNOME, DATA_NASCITA, SESSO, CODICE_FISCALE, " & _
            "           COMUNI.CODICE AS COMUNICODICE, ASL.CODICE AS ASLCODICE, REGIONI.CODICE AS REGIONICODICE, " & _
            "           Nazioni.Nome as NazioniNome, Nazioni.CODICEALFA2 as NazioniCODICEALFA2, Nazioni.CODICE_ISTAT as NazioniCODICE_ISTAT, " & _
            "           TipiRicetta.Codice as TipiRicettaCodice, TIPOLOGIE_ESENZIONE.CODICE AS TIPOLOGIE_ESENZIONECODICE, " & _
            "           TIPOLOGIE_ESENZIONE.ESENZIONE_QUOTA " & _
            "FROM       (((((((RICETTE " & _
            "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=RICETTE.CODICE_PAZIENTE) " & _
            "           LEFT OUTER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) " & _
            "           LEFT OUTER JOIN ASL ON ASL.KEY=PAZIENTI.CODICE_ASL) " & _
            "           LEFT OUTER JOIN REGIONI ON REGIONI.KEY=PAZIENTI.CODICE_REGIONE) " & _
            "           INNER JOIN Nazioni ON Nazioni.KEY=PAZIENTI.NazioniID) " & _
            "           INNER JOIN TipiRicetta ON TipiRicetta.KEY=Ricette.TipiRicettaID) " & _
            "           INNER JOIN TIPOLOGIE_ESENZIONE ON TIPOLOGIE_ESENZIONE.KEY=RICETTE.CODICE_ESENZIONE) " & _
            "WHERE      (NOT FLAG=3 AND " & _
            "           ANNO=" & anno & " AND " & _
            "           MESE=" & mese & ") " & _
            "ORDER BY   PROGRESSIVO_ANNUALE"
    rsDataset.Open strSql, cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        MsgBox "Nessuna ricetta per il mese di " & MonthName(mese), vbInformation, "Genera file C"
        CreaFiles = False
        Exit Function
    End If
    Do While Not rsDataset.EOF
        i = i + 1
        ReDim Preserve mNodo(i)
        Set mNodo(i) = New CNodoFileC
        
        strSql = "SELECT    MEDICI_BASE.CODICE AS MEDICI_BASECODICE, TIPOLOGIE_MEDICO.CODICE AS TIPOLOGIE_MEDICOCODICE " & _
                 "FROM      (MEDICI_BASE " & _
                 "          LEFT OUTER JOIN TIPOLOGIE_MEDICO ON TIPOLOGIE_MEDICO.KEY=MEDICI_BASE.CODICE_TIPO_MEDICO) " & _
                 "WHERE     MEDICI_BASE.KEY=" & rsDataset("CODICE_MEDICO")
        rsAppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        ' stampa il tipo di medico solo a quelli di avellino
        mNodo(i).Medico = Space(8 + 16) & CompletaTesto(rsAppo("MEDICI_BASECODICE") & "", 7) & IIf(Mid(mIntestazioneC1, 1, 3) = "201", rsAppo("TIPOLOGIE_MEDICOCODICE"), Space(1))
        rsAppo.Close
        
        If UCase(rsDataset("NazioniNome")) = UCase("Italia") Then
            blnPazienteEstero = False
        Else
            blnPazienteEstero = True
        End If
        
        mNodo(i).Paziente = CompletaTesto(rsDataset("PAZIENTICOGNOME"), 40) & _
                            CompletaTesto(rsDataset("PAZIENTINOME"), 40) & _
                            rsDataset("CODICE_FISCALE") & _
                            IIf(rsDataset("SESSO") = "F", 2, 1) & _
                            sistemaData(rsDataset("DATA_NASCITA")) & "0"
        If blnPazienteEstero Then
           ' mNodo(i).Paziente = mNodo(i).Paziente & _
                            "999" & CompletaNumeri(rsDataset("NazioniCODICE_ISTAT"), 3) & _
                            CompletaNumeri(rsDataset("NazioniCODICE_ISTAT"), 3) & _
                            "   "
            mNodo(i).Paziente = mNodo(i).Paziente & _
                            "999" & CompletaNumeri(rsDataset("NazioniCODICE_ISTAT"), 3) & _
                            "   " & "   "
        Else
            mNodo(i).Paziente = mNodo(i).Paziente & _
                            rsDataset("COMUNICODICE") & _
                            rsDataset("REGIONICODICE") & _
                            CompletaNumeri(rsDataset("ASLCODICE"), 3)
        End If
                                    
        Dim strTipiRicettaCodice As String
        strTipiRicettaCodice = CompletaTesto(UCase(rsDataset("TipiRicettaCodice")), 2)
        mNodo(i).TipoRicetta = strTipiRicettaCodice
        If strTipiRicettaCodice = "UE" Or strTipiRicettaCodice = "NE" Or strTipiRicettaCodice = "EE" Or strTipiRicettaCodice = "NX" Then
            mNodo(i).CodiceIstituzioneCompetente = CompletaTesto(rsDataset("CodiceIstituzioneCompetente"), 28)
        Else
            mNodo(i).CodiceIstituzioneCompetente = Space(28)
        End If
        If strTipiRicettaCodice = "UE" Or strTipiRicettaCodice = "NE" Then
            mNodo(i).StatoEstero = IIf(blnPazienteEstero, CompletaTesto(rsDataset("NazioniCODICEALFA2"), 2), Space(2))
            mNodo(i).NumeroIdentificativoPersonale = CompletaTesto(rsDataset("NumeroIdentificativoPersonale"), 20)
            mNodo(i).NumeroIdentificativoTessera = CompletaTesto(rsDataset("NumeroIdentificazioneTessera"), 20)
        Else
            mNodo(i).StatoEstero = Space(2)
            mNodo(i).NumeroIdentificativoPersonale = Space(20)
            mNodo(i).NumeroIdentificativoTessera = Space(20)
        End If
        
        mNodo(i).StampaPC = IIf(CBool(rsDataset("STAMPATO_PC")), 1, 0)
        mNodo(i).Identificativi = anno & CompletaTesto(mSTS, 6) & "00" & CompletaNumeri(rsDataset("PROGRESSIVO_ANNUALE"), 8)
        
        ' prescrizioni
        k = 0
        importoTotale = 0
        
        strSql = "SELECT    PRESCRIZIONI.*, NOMENCLATORE_TARIFFARIO.CODICE  " & _
                "FROM       (PRESCRIZIONI " & _
                "           INNER JOIN NOMENCLATORE_TARIFFARIO ON NOMENCLATORE_TARIFFARIO.KEY=PRESCRIZIONI.CODICE_PRESTAZIONE) " & _
                "WHERE      CODICE_RICETTA=" & rsDataset("KEY")
        rsAppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsAppo.EOF
            k = k + 1
            If k = 1 Then
                If rsDataset("DATA_PRENOTAZIONE") = rsAppo("DATA_INIZIO") Then
                    cronologiaPrescrizione = 0
                Else
                    cronologiaPrescrizione = 1
                End If
            Else
                cronologiaPrescrizione = 1
            End If
            mNodo(i).AddPrescriozione (Format(k, "00") & "^" & CompletaTesto(rsAppo("CODICE"), 7) & "^" & Format("0", "000000.00") & "^" & Format("0", "000000.00") & "^" & Format(rsAppo("IMPORTO") * rsAppo("QUANTITA"), "000000.00") & "^" & sistemaData(rsAppo("DATA_FINE")) & "^" & CompletaNumeri(rsAppo("QUANTITA"), 3) & "^" & cronologiaPrescrizione)
            importoTotale = importoTotale + rsAppo("IMPORTO") * rsAppo("QUANTITA")
            rsAppo.MoveNext
        Loop
        ' calcola la quota regionale o aggiuntiva
        If rsDataset("ESENZIONE_DOPPIA") Then
            coefficienteQuotaAggiuntiva = 0
        Else
            If rsDataset("CODICE_ESENZIONE") = -1 Then
                coefficienteQuotaAggiuntiva = 1
            ElseIf CBool(rsDataset("ESENZIONE_QUOTA")) Or UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05" Then
                coefficienteQuotaAggiuntiva = 0
            Else
                coefficienteQuotaAggiuntiva = 1 / 2
            End If
        End If

        importoTotale = importoTotale - IIf(rsDataset("CODICE_ESENZIONE") = -1 Or UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05", ticket + quotaNazionale, 0) - coefficienteQuotaAggiuntiva * quotaAggiuntiva
        mNodo(i).AddPrescriozione ("99" & "^" & Space(7) & "^" & Format(IIf(rsDataset("CODICE_ESENZIONE") = -1 Or UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05", ticket, 0), "000000.00") & "^" & Format((coefficienteQuotaAggiuntiva * quotaAggiuntiva + IIf(rsDataset("CODICE_ESENZIONE") = -1 Or UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05", quotaNazionale, 0)), "000000.00") & "^" & Format(importoTotale, "000000.00") & "^" & Space(8) & "^" & String(3, "0") & "^" & cronologiaPrescrizione)
        rsAppo.Close
                
        mNodo(i).PosizioneContabile = IIf(UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05" And CBool(rsDataset("ESENZIONE_QUOTA")) = True, 1, Space(1))
        mNodo(i).TipologiaRicetta = Choose(rsDataset("TIPOLOGIA_RICETTA") + 1, " ", "S", "H", "A")
        mNodo(i).NumeriRicetta = Space(1) & rsDataset("NUMERO_RICETTA")
        mNodo(i).LeDate = sistemaData(rsDataset("DATA_RICETTA")) & sistemaData(rsDataset("DATA_PRENOTAZIONE"))
        rsAppo.Open "SELECT * FROM  TIPOLOGIE_ESENZIONE WHERE KEY=" & rsDataset("CODICE_ESENZIONE"), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        mNodo(i).Esenzione = IIf(UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05" Or rsDataset("CODICE_ESENZIONE") = -1, 1, 2) & CompletaTesto(IIf(rsAppo("KEY") = -1, "", rsAppo("CODICE")), 6) & IIf(CBool(rsDataset("ESENZIONE_DOPPIA")) Or CBool(rsDataset("ESENTE_REDDITO")), 1, 2)
        rsAppo.Close
        mNodo(i).Onere = IIf(rsDataset("CODICE_ESENZIONE") = -1 Or UCase(rsDataset("TIPOLOGIE_ESENZIONECODICE")) = "E05", 7, 1)
        mNodo(i).Mazzette = CompletaNumeri(rsDataset("MAZZETTA1"), 3) & CompletaNumeri(rsDataset("MAZZETTA2"), 5)
        mNodo(i).AnnoMese = anno & CompletaNumeri(Trim(Str(mese)), 2)
        mNodo(i).ProgressivoSuRicetta = CompletaNumeri(rsDataset("PROGRESSIVO_RICETTA"), 5)
    
        rsDataset.MoveNext
    Loop
    Set rsDataset = Nothing
    
    
    
    ' salva i file
    Dim vett() As String
    Dim rigaFileC1 As String * 265
    Dim rigaFileC2 As String * 246
    Dim nomeFile As String
    Dim nomeFile2 As String
    
    nomeFile = structIntestazione.sCodiceSTS & "_" & sistemaData(GetUltimoGiorno(mese, anno))
    nomeFile2 = "M" & sistemaData(GetUltimoGiorno(mese, anno)) & structIntestazione.sCodiceSTS & "_27"
    If Dir(nomeFile & ".c1") <> "" Then Kill nomeFile & ".c1"
    If Dir(nomeFile & ".c2") <> "" Then Kill nomeFile & ".c2"
    If Dir(nomeFile & "c1.txt") <> "" Then Kill nomeFile & "c1.txt"
    If Dir(nomeFile & "c2.txt") <> "" Then Kill nomeFile & "c2.txt"
    If Dir(nomeFile & "_ANA" & ".zip") <> "" Then Kill nomeFile & "_ANA" & ".zip"
    If Dir(nomeFile & "_SAN" & ".zip") <> "" Then Kill nomeFile & "_SAN" & ".zip"
    If Dir(nomeFile & ".zip") <> "" Then Kill nomeFile & ".zip"
    If Dir(nomeFile2 & ".c1") <> "" Then Kill nomeFile2 & ".c1"
    If Dir(nomeFile2 & ".c2") <> "" Then Kill nomeFile2 & ".c2"
    If Dir(nomeFile2 & ".zip") <> "" Then Kill nomeFile2 & ".zip"
    

    
    Open nomeFile & ".c1" For Output As #1
    Open nomeFile & ".c2" For Output As #2
    
    For i = 1 To UBound(mNodo)
        vett = Split(mNodo(i).CreaRigaC1(mIntestazioneC1), "^$^")
        For k = 0 To UBound(vett)
            rigaFileC1 = vett(k)
            If Len(rigaFileC1) <> 265 Then
                MsgBox "Errore di generazione dei file C", vbCritical, "Attenzione"
            Else
                Print #1, rigaFileC1
            End If
        Next k
        vett = Split(mNodo(i).CreaRigaC2(mIntestazioneC2), "^$^")
        For k = 0 To UBound(vett)
            rigaFileC2 = vett(k)
            If Len(rigaFileC2) <> 246 Then
                MsgBox "Errore di generazione dei file C", vbCritical, "Attenzione"
            Else
                Print #2, rigaFileC2
            End If
        Next k
    Next i
    
    Close #1
    Close #2
    
    FileCopy nomeFile & ".c1", nomeFile & "c1.txt"
    FileCopy nomeFile & ".c2", nomeFile & "c2.txt"
    
    ' zippa i file
    
    Dim MYUSER As ZIPUSERFUNCTIONS
    Dim retcode As Long

    MYUSER.DLLPrnt = Puntatore(AddressOf Stampa_messaggi_zip)
    MYUSER.DLLPASSWORD = 0&
    MYUSER.DLLCOMMENT = 0&
    MYUSER.DLLSERVICE = 0&
    retcode = ZpInit(MYUSER)

    Dim MYOPT As ZPOPT
    retcode = ZpSetOptions(MYOPT)

    Dim files As ZIPnames

    files.s(0) = nomeFile & ".c1"
    retcode = ZpArchive(1, nomeFile & "_ANA" & ".zip", files)
    
    files.s(0) = nomeFile & ".c2"
    retcode = ZpArchive(1, nomeFile & "_SAN" & ".zip", files)
    
    files.s(0) = nomeFile & ".c1"
    files.s(1) = nomeFile & ".c2"
    retcode = ZpArchive(2, nomeFile & ".zip", files)
    
    FileCopy nomeFile & ".c1", nomeFile2 & ".c1"
    FileCopy nomeFile & ".c2", nomeFile2 & ".c2"
    files.s(0) = nomeFile2 & ".c1"
    files.s(1) = nomeFile2 & ".c2"
    retcode = ZpArchive(2, nomeFile2 & ".zip", files)

    If retcode = 0 Then
        Kill nomeFile2 & ".c1"
        Kill nomeFile2 & ".c2"
        If percorso <> App.path Then
            'imposta cartella con il nome del mese e dell'anno
            NuovaCartella = percorso & "\FILE C " & UCase(MonthName(mese)) & " " & anno
            If Dir(NuovaCartella & "\" & nomeFile & ".c1") <> "" Then Kill NuovaCartella & "\" & nomeFile & ".c1"
            If Dir(NuovaCartella & "\" & nomeFile & ".c2") <> "" Then Kill NuovaCartella & "\" & nomeFile & ".c2"
            If Dir(NuovaCartella & "\" & nomeFile & "c1.txt") <> "" Then Kill NuovaCartella & "\" & nomeFile & "c1.txt"
            If Dir(NuovaCartella & "\" & nomeFile & "c2.txt") <> "" Then Kill NuovaCartella & "\" & nomeFile & "c2.txt"
            If Dir(NuovaCartella & "\" & nomeFile & "_ANA" & ".zip") <> "" Then Kill NuovaCartella & "\" & nomeFile & "_ANA" & ".zip"
            If Dir(NuovaCartella & "\" & nomeFile & "_SAN" & ".zip") <> "" Then Kill NuovaCartella & "\" & nomeFile & "_SAN" & ".zip"
            If Dir(NuovaCartella & "\" & nomeFile & ".zip") <> "" Then Kill NuovaCartella & "\" & nomeFile & ".zip"
            If Dir(NuovaCartella & "\" & nomeFile2 & ".zip") <> "" Then Kill NuovaCartella & "\" & nomeFile2 & ".zip"

            If DirExists(NuovaCartella) = False Then MkDir NuovaCartella
                        
            Name App.path & "\" & nomeFile & ".c1" As NuovaCartella & "\" & nomeFile & ".c1"
            Name App.path & "\" & nomeFile & ".c2" As NuovaCartella & "\" & nomeFile & ".c2"
            Name App.path & "\" & nomeFile & "c1.txt" As NuovaCartella & "\" & nomeFile & "c1.txt"
            Name App.path & "\" & nomeFile & "c2.txt" As NuovaCartella & "\" & nomeFile & "c2.txt"
            Name App.path & "\" & nomeFile & "_ANA" & ".zip" As NuovaCartella & "\" & nomeFile & "_ANA" & ".zip"
            Name App.path & "\" & nomeFile & "_SAN" & ".zip" As NuovaCartella & "\" & nomeFile & "_SAN" & ".zip"
            Name App.path & "\" & nomeFile & ".zip" As NuovaCartella & "\" & nomeFile & ".zip"
            Name App.path & "\" & nomeFile2 & ".zip" As NuovaCartella & "\" & nomeFile2 & ".zip"
        End If
        MsgBox "File C1 e C2 creati correttamente", vbInformation, "Gestione file C"
    End If
    
    CreaFiles = True
End Function

