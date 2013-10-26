Attribute VB_Name = "modVariabili"
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             ''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modVariabili.bas
'
' <b>Descrizione</b>: Variabili publiche
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.01
Option Explicit

' oggetti per il menu
Public objMenuEx As cMenuEx
Public glMenuStyle As EnumStyleMenu '/ style for menuXP
Public gbSubClassMenu As Boolean

Public Const TRACCIATO As Boolean = True        ' per evitare di fare la tracciatura
Public Const STESSO_PAZIENTE As Boolean = True  ' nuova gestione dei pazienti nei form del primo menu
Public Const RIDISPONI_FORMS As Boolean = True  ' nuovo sistema di sistemazione dei form nel padre
Public Const NUOVA_TOOLBAR As Boolean = False    ' ribbontab


' intestazione file scannerizzati
Public Const E_ST As String = "e_st"            ' esami strumentali
Public Const T_AC As String = "t_ac"            ' tratt acque
Public Const M_PS As String = "m_ps"            ' monit psicosociale
Public Const S_DP As String = "s_dp"            ' docum pazienti
Public Const M_TR As String = "m_tr"            ' monit trapianti

Public Const CODICESTS_GAMMADIAL As String = "AD0143"       'ASL NAPOLI 2 NORD
Public Const CODICESTS_CGA As String = "AD0145"
Public Const CODICESTS_DIALIFE As String = "AD0151"
Public Const CODICESTS_CAMPANO As String = "AD0152"
Public Const CODICESTS_DIALGEST As String = "AD0153"
Public Const CODICESTS_SBIAGIO As String = "AD0155"
'Public Const CODICESTS_SPIOX As String = "AD0156"

Public Const CODICESTS_NEPHRON As String = "AD0169"         'ASL NAPOLI 3 SUD
Public Const CODICESTS_DELTA As String = "AD0171"
Public Const CODICESTS_POGGIOMARINO As String = "AD0163"
Public Const CODICESTS_MOSCATI As String = "AD0112"

Public Const CODICESTS_EM_IRPINA As String = "AD0094"       'AVELLINO
Public Const CODICESTS_BARTOLI As String = "AD0099"

Public Const CODICESTS_LA_PECCERELLA As String = "AD0133"   'BENEVENTO
Public Const CODICESTS_SANNIOMEDICA As String = "AD0137"

Public Const CODICESTS_SANT_ANDREA As String = "AD0082"     'CASERTA
Public Const CODICESTS_HELIOS As String = "AD0089"
Public Const CODICESTS_SODAV As String = "AD0078"

Public Const CODICESTS_SM2 As String = "000SM2"             'POTENZA

Public Const colArancione As Long = &HC0FFFF

Public MenuEvents As CEvents

Public cnPrinc As Connection            ' connessione all'archivio dati
Public cnTrac As Connection             ' connessione all'archivio tracciatura

Public laData As String                 ' la data caricata con frmCalendario e altro
Public laOra As String                  ' l'orario caricato con frmOrario (per i turni)

Public SpegniPc As Boolean              ' true spegne il pc al termine del backup
Public isCorrotto As Boolean            ' true siginifica che il db all'apertura è risultato corrotto

Public oPazientiKey As New clsPazientiKey

Public strConnectionStringCentro As String
Public strConnectionStringTracciatura As String
Public strNomeTabella As String
Public intValore As Integer

Public dt_rott_rene As Date             ' data rottamazione del rene selezionato nella flex
Public cod_rene As Integer
Public sostituito As Boolean

Public ModificaProduttore As Boolean    ' variabile per il formApparatiInput per caricare il formTrova
Public ModificaManutentore As Boolean   ' variabile per il formApparatiInput per caricare il formTrova
Public StampaApparati As Boolean        ' variabile per il formStampaApparati per caricare il formTrova
Public MantieniKeyReturn As Integer     ' variabile per il form Apparati
Public KeyApparato As Integer           ' variabile apparato che viene passata per manutenzione CODICE_APPARATO
Public KeyReturnManutenzione As Integer ' variabile Key per la selezione della manutenzione
Public Selezionato As Boolean           ' variabile per il frmApparati
Public SelezionatoManutenzione As Boolean ' variabile per il frmApparati per caricare/inserire la manutenzione
Public numKey As Integer
Public mRagioneSociale As String         ' variabile per trasferire la stringa dal formProduttore al formTrova

