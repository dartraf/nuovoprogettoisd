Attribute VB_Name = "modZip"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modZip.bas
'
' <b>Descrizione</b>: Insieme di funzioni e oggetti per zippare file
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.07
Option Explicit

Public Declare Function ZpInit Lib "zip32.dll" _
(ByRef Zipfun As ZIPUSERFUNCTIONS) As Long

Public Declare Function ZpSetOptions Lib "zip32.dll" _
(ByRef Opts As ZPOPT) As Long

Public Declare Function ZpArchive Lib "zip32.dll" _
(ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long



Public Type ZIPnames
    s(0 To 99) As String
End Type

Public Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 byte
    szRootDir As String ' fino a 256 byte
End Type

Public Type ZIPUSERFUNCTIONS
    DLLPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type

Public Type CBChar
    ch(4096) As Byte
End Type

Public Function Stampa_messaggi_zip(ByRef fname As CBChar, ByVal lenght As Long) As Long
    Dim messaggio As String ' conterrà il messaggio
    Dim i As Long   ' lunghezza in byte del messaggio
    
    On Error Resume Next    'sempre necessario nelle funzioni di callback
    ' ricostruisco il messaggio a partire dalla stringa di byte
    For i = 0 To lenght
        If fname.ch(i) = 0 Then Exit For Else messaggio = messaggio + Chr(fname.ch(i))
    Next i
    
    'Debug.Print "" & messaggio
     DoEvents
    Stampa_messaggi_zip = 0
    
End Function

Public Function Puntatore(ByVal lp As Long) As Long
    Puntatore = lp
End Function
