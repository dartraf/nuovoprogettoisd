Attribute VB_Name = "modTipi"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modTipi.bas
'
' <b>Descrizione</b>: Tipi e relative variabili publiche
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.11
Option Explicit

'' struttura usata nella stampa degli esami di lab
Private Type riga
    nome As String
    codice As Integer
    unita As String
    minmax As String
    valori(1 To 12) As Variant
End Type
Private Type structEsami
    righe() As riga
End Type
Public Type tabEsami
    anno As Integer
    esami() As structEsami
End Type
'-----------------------------------------

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Type intestazione
    sPaziente As String                 ' var per il report della cartella clinica
    sDataPaziente As Date
    sRagione As String
    sTipoAmbulatorio As String
    sIndirizzo As String
    sCap As String
    sCitta As String
    sProv As String
    sTelefono As String
    sFax As String
    sCodiceFiscale As String
    sIva As String
    sMail As String
    sDirettoreSanitarioNome As String
    sDirettoreSanitarioCognome As String
    sCodiceSTS As String
    sCodiceAsl As Integer
    sCodiceDistretto As Integer
    sLogoISO As Boolean
    sLogoAziendale As Boolean
    sLogoQualita As Boolean
    sNomeLogoISO As String
    sNomeLogoQualita As String
    sNomeLogoAziendale As String
End Type
Public structIntestazione As intestazione
'--------------------------

Type apri
    pathVolume As String
    pathTrueCrypt As String
    pathDB As String                    ' il percorso del db sul computer
    pathExe As String
    server As Boolean
    nomeServer As String
    pathNomeCertificato As String
    F1abiliata As Boolean
    strFromModuliWord As String
End Type
Public structApri As apri
'-------------------------

Enum schedaGiornaliera
    tpCOMPILAZIONE = 1
    tpCONSULTAZIONE
    tpSTRAORDINARIA
End Enum
Public tScheda As schedaGiornaliera
'------------------------

Type structReni
    key As Integer
    postazione As String
    numero_apparato As String
    monitor As String
    Tipo As String
End Type
Public tReni As structReni
'------------------------

Enum tipoFileRicette
    tpFILEC = 1
    tpFILEXML = 2
End Enum
Public tFileRicette As tipoFileRicette
'------------------------

Enum orario
    tpMAT = 1
    tpPOM
    tpSER
    tpNULL
End Enum
Public tOrario As orario
'------------------------

Enum storico
    tpsPESO = 1
    tpsFILTRO
    tpsLINEE
End Enum
Private Type structStorico
    Tipo As storico                 ' tipo passato a storico
    condizione As String            ' condizione della ricerca di storico
End Type
Public tStorico As structStorico
'------------------------

Enum Stampa
    tpFOGLIOVIAGGIO = 1
    tpIMPEGNATIVE = 2
    tpMODULOFIRMEPAZIENTE = 3
    tpSCHEDADIALITICASETTIMANALE = 4
    tpKTVANNUALE
    tpTSATANNUALE
    tpPTHAnnuale
    tpCAPAnnuale
End Enum
Public tStampa As Stampa
'------------------------

Enum tipoInput
    tpISINGOLO = 1                       ' per il form tabelle
    tpICOMPOSTO
    
    tpIVOCI                           ' per i rispettivi form
    tpITERAPIADOMICILIARE
    tpITERAPIADIALITICA
    tpIPASSWORD
    
    tpITIPIESAMILAB                   ' per le descrizioni
    tpIESAMI                          ' per gli esami
    tpIRENI                           ' reni
    tpIRICOVERI                     ' importante che siano consecutivi
    tpIEPISODI
    tpITRASFUSIONI
    tpISIEROCONVERSIONI
    tpICOLTURE
    
    tpITERAPIESTRAORDINARIE           ' per la scheda dialitica straordinaria

    tpIPRESCRIZIONI
    tpIDISTRETTI
    tpINOMENCLATORE
    tpICOMUNI
    tpIASL
    tpITIPOLOGIEMEDICO
    tpIESENZIONE

End Enum
Private Type structInput
    Tipo As tipoInput                    ' tipo di input (form chiamante)
    mantieniDati As Boolean              ' indica se i dati devono essere mantenuti o cancellati dal frmInput
    v_valori(1 To 7) As String           ' valori restituiti dal form input
End Type
Public tInput As structInput
'------------------------

Enum accesso
    tpAMEDICO = 1
    tpAINFERMIERE
    tpACONTABILE
    tpAMASTER
End Enum
Private Type tipoAccesso
    cognome As String
    nome As String
    Tipo As accesso
    pass As String
    key As Integer
End Type
Public tAccesso As tipoAccesso
'------------------------

Enum tipoTrova
    tpPAZIENTE = 1
    tpMEDICOBASE
    tpMEDICOREFER
    tpINFERMIERE
    tpPSICOLOGI
    tpACCOMPAGNATORI
    tpPRODUTTORE_MANUTENTORE
End Enum
Private Type structTrova
    Tipo As tipoTrova                   ' tipo da trovare
    keyReturn As Integer                ' codice_XXX restituito dal frmTrova
    condizione As String                ' condizione di ricerca sql
    condStato As String                 ' condizione di caricamento dello stato nel cboStato
    NomeStriga As String                      ' nome_XXX restituito dal frmTrova in caso della tabella PRODUTTORE_MANUTENTORE
    isOpenFromInfoGenerali As Boolean
    isOpenFromEsamiPrescriz As Boolean
    End Type
Public tTrova As structTrova              '  il tipo passato ai form con molteplici funzioni frmTrova
'------------------------

Enum tipoTabella
    ' elementi costituiti da due informazioni
    tpRegioni
    tpTIPOLOGIEMEDICO
    tpESENZIONI
    tpEDTA
    ' elementi costituiti da tre informazioni
    tpCOMUNI
    tpasl
    tpDISTRETTI
    ' altri
    tpNOMENCLATORE
    tpESAME
    tpRENI
End Enum
Public tTabelle As tipoTabella
'------------------------

Enum tipoElenca
    tpACCESSO = 1
    tpDIARIO
    tpSCHEDEDIALITICHE
    tpREGISTRAZIONESAMI
    tpESAMISTRUMENTALI
    tpCOLTURE
    tpMON_ACC_VASCOLARE
    tpMON_VAL_PSICO
    tpMON_VACC_EPATITE
    tpMON_TRAT_ACQUE
    tpPRESCRIZIONI
    tpESPORTAESAMI                  ' KTV O TSAT
End Enum
Private Type structElenca
    Tipo As tipoElenca              ' tipo passato ad elenca
    condizione As String            ' condizione della ricerca di elenca
End Type
Public tElenca As structElenca
'------------------------

Enum pass
    tVERIFICA = 1
    tCAMBIA = 2
End Enum
Private Type structPass
    Tipo As pass
    password As String
    key As Integer
End Type
Public tipoPass As structPass
'------------------------

Enum tipoDisconnetti
    tpDCHIUDICONBACKUP = 1
    tpDLOGIN = 2
    tpDANNULLA = 3
End Enum
Public tDisconnetti As tipoDisconnetti
'------------------------

Enum tipoRete
    tpCONNETTI = 1
    tpDISCONNETTI = 2
End Enum
Public tRete As tipoRete
'------------------------

Enum tipoPeriodo
    tpMENSILE = 1
    tpTRIMESTRALE = 2
    tpSEMESTRALE = 3
    tpANNUALE = 4
    tpBIMESTRALE = 5
    tpPROBLEMI = 6
End Enum
Private Type structEsamiInput
    periodo As tipoPeriodo
    ' se interogruppo=0 conserva il key del singolo esame
    ' se interogruppo=1 conserva il key dell'associazione: valori negativi => codice associazione per esami di lab, valori pos => codice organo per esami strumentali
    codiceAssociazione As Integer
    interoGruppo As Integer           ' 0 singolo esame, -1 annulla, >0 intero gruppo
End Type
Public tEsamiPeriodici As structEsamiInput
'------------------------

Enum tipoPeriferica
    tpRIPRISTINA = 1
    tpBACKUP = 2
End Enum
Public tPeriferica As tipoPeriferica
'------------------------

Enum tipoStampeRiepilogo
    tpFATTURA = 1
    tpXPAZIENTE
    tpXTOTALIPERPRESTAZIONE
    tpXTOTALIPERASL
    TPXMAZZETTEMENSILI
    tpXMAZZETTASINGOLA
    tpXMAZZETTEDISTRETTI
    tpXASLDISTRETTI
    tpXIMPEGNATIVE
End Enum
Public tStampeRiepilogo As tipoStampeRiepilogo
'------------------------

' Per il form Apparati
Enum TipoTabellaManutenzione
    tpMANUTENZIONEORDINARIA
    tpMANUNTENZIONESTRAORDINARIA
End Enum
Public tTabellaManutenzione As TipoTabellaManutenzione
'------------------------


Enum tipoStato
    tpNoneStatoPaziente = -1
    tpDIALISI = 0
    TPDECEDUTO = 1
    TPTRASFERITO = 2
    TPTRAPIANTO = 3
    TPOSPITE = 4
    tpAMBULATORIALE = 5
End Enum
Private Type varStato                       ' variabili dello stato del paziente
    dataStato As String                   ' contiene data decesso, trapianto, trasferimanto
    statoPaz As tipoStato                 ' indica lo stato del paziente durante le varie modifiche
    donatore As Byte                      ' 0 cadavere   1 vivente   2 non immesso
    dataArrivi(1 To 3) As String          ' vettori utilizzati in frmStatoPaziente
    dataPartenza(1 To 3) As String
    centriProv(1 To 3) As Integer
End Type
Public statoPaziente As varStato
'-------------------------

Enum tipoDocumentoEsterno
    tpSCANESAMISTRUMENTALI = 0
    tpSCANMONITORAGGIO
    tpSCANTRAPIANTI
    tpSCANTRATTAMENTOACQUE
    tpSCANDOCPAZIENTI
End Enum
Public tDocumentiEsterni As tipoDocumentoEsterno
'-------------------------

Enum enumSelezionaDaCbo
    tpGRUPPI_ESAMI
End Enum
Private Type structSelezionaDaCbo
    tipoCampo As enumSelezionaDaCbo
    valoreDaEvitare As Integer
    valoreSelezionato As Integer
    nuovoInserimento As Boolean
End Type
Public tSelezionaDaCbo As structSelezionaDaCbo
'-------------------------

Enum enumSessioni
    tpNoneSession = -1
    tpPariMattina = 0
    tpPariPomeriggio = 1
    tpPariSera = 2
    tpDispariMattina = 3
    tpDispariPomeriggio = 4
    tpDispariSera = 5
End Enum
'-------------------------

Private Type structFiltroStato
    statoPaziente As tipoStato
    isTutteLeDate As Boolean
    dataDal As Date
    dataAl As Date
End Type
Public tFiltroStato As structFiltroStato
'-------------------------

Public Enum enumTipoApertura
    tpTipoAperturaNone = 0
    tpTrovaPaziente = 1
    tpCaricaPaziente = 2
End Enum

Public Enum enumTipoTabPersonale
    INFERMIERI
    MEDICI_DIALISI
    MEDICI_REFERTANTI
    PSICOLOGI
End Enum

Public Enum enumTipoTabSingolo
    filtro
    ANTICOAGULANTI
    ORGANO
    Medicinali
    TITOLIDIARIO
    LINEE
    AGO
End Enum

