VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNodoFileC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe che gestiche l'annullamento nella flexgird"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo di classe - CNodoFileC.cls
'
' <b>Descrizione</b>: Classe Nodo della pila FileC, a sua volta una pila
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.36
Option Explicit

Private Type CNodoNodoFileC
    ' intestazione rimane al di fuori
    ' per file c1 e c2
    progressivoRiga As String
    ' solo per file c2
    prescrizione As String
    importo As String
    ticket As String               ' in realta è l importo a carico dell'assistito
    dataErogazione As String
    Quantita As String
    cronologiaPrescrizione As String
    quotaAggiuntiva As String
End Type

Private mNodo() As CNodoNodoFileC
' intestazione rimane al di fuori
' solo per file c1
Public Paziente As String
Public Medico As String
Public TipoRicetta As String
Public CodiceIstituzioneCompetente As String
Public StatoEstero As String
Public NumeroIdentificativoPersonale As String
Public NumeroIdentificativoTessera As String
Public StampaPC As String
' per file c1 e c2
Public Identificativi As String
' solo per file c2
Public NumeriRicetta As String
Public LeDate As String
Public Esenzione As String
Public Mazzette As String
Public ProgressivoSuRicetta As String
Public AnnoMese As String
Public Onere As String
Public TipologiaRicetta As String
Public PosizioneContabile As String

''
' Inizializza la pila
'
' @param
' @param
' @return
' @remarks
Private Sub Class_Initialize()
    ReDim mNodo(0)
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
End Sub

''
' Aggiunge un nodo
'
' @param vParam parametri per la creazione del nodo da aggiungere
' @param
' @return
' @remarks
Public Sub AddPrescriozione(vParam As String)
    Dim vett() As String
    
    vett = Split(vParam, "^")
    ReDim Preserve mNodo(UBound(mNodo) + 1)
    mNodo(UBound(mNodo)).progressivoRiga = vett(0)
    mNodo(UBound(mNodo)).prescrizione = vett(1)
    mNodo(UBound(mNodo)).ticket = vett(2)
    mNodo(UBound(mNodo)).quotaAggiuntiva = vett(3)
    mNodo(UBound(mNodo)).importo = vett(4)
    mNodo(UBound(mNodo)).dataErogazione = vett(5)
    mNodo(UBound(mNodo)).Quantita = vett(6)
    mNodo(UBound(mNodo)).cronologiaPrescrizione = vett(7)
End Sub

''
' Crea le righe dello stesso paziente del file c1
'
' @param intestazione intestazione da aggiungere all'inizio di ogni riga
' @param
' @return
' @remarks
Public Function CreaRigaC1(intestazione As String) As String
    Dim testo As String
    Dim i As Integer
    
    For i = 1 To UBound(mNodo)
        testo = testo & "150" & intestazione & Medico & Paziente & TipoRicetta & CodiceIstituzioneCompetente & StatoEstero & NumeroIdentificativoPersonale & NumeroIdentificativoTessera & mNodo(i).progressivoRiga & StampaPC & Identificativi
        If i <> UBound(mNodo) Then
            testo = testo & "^$^"
        End If
    Next i
    CreaRigaC1 = testo
End Function

''
' Crea le righe dello stesso paziente del file c2
'
' @param intestazione intestazione da aggiungere all'inizio di ogni riga
' @param
' @return
' @remarks
Public Function CreaRigaC2(intestazione As String) As String
    Dim testo As String
    Dim i As Integer
    
    For i = 1 To UBound(mNodo)
        testo = testo & "150" & intestazione & TipologiaRicetta & "111  " & NumeriRicetta & mNodo(i).progressivoRiga & LeDate & mNodo(i).dataErogazione & "t" & mNodo(i).prescrizione & mNodo(i).Quantita & Esenzione & Onere & Mazzette & mNodo(i).ticket & mNodo(i).quotaAggiuntiva & mNodo(i).importo & Space(7) & String(4, "0") & Space(53) & mNodo(i).cronologiaPrescrizione & String(8, "0") & PosizioneContabile & Space(10) & Identificativi & AnnoMese & ProgressivoSuRicetta & "000"
        If i <> UBound(mNodo) Then
            testo = testo & "^$^"
        End If
    Next i
    CreaRigaC2 = testo
End Function
