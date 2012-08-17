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

Public Const CODICESTS_BARTOLI As String = "AD0099"
Public Const CODICESTS_HELIOS As String = "AD0089"
Public Const CODICESTS_SODAV As String = "AD0078"

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
