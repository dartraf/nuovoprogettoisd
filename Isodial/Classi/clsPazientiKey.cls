VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPazientiKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public intPazientiKey As Integer
Public intNumeroFormAperti As Integer

Public listFormAperti As New clsLista
Private strPazienteNome As String
Private strPazienteCognome As String

Private Sub Class_Initialize()
    intPazientiKey = 0
    intNumeroFormAperti = 0
    strPazienteCognome = ""
    strPazienteNome = ""
    listFormAperti.Refresh
End Sub

Public Sub OnClosingForm(inMeCaption As String)
    If intNumeroFormAperti > 0 Then
        intNumeroFormAperti = intNumeroFormAperti - 1
        listFormAperti.Remove (inMeCaption)
    End If
    If intNumeroFormAperti = 0 Then
        listFormAperti.Refresh
        intPazientiKey = 0
        strPazienteCognome = ""
        strPazienteNome = ""
    End If
End Sub

Public Function GetPazienteInfo() As String
    GetPazienteInfo = strPazienteCognome & " " & strPazienteNome
End Function

Public Sub ImpostaPazientiKey(inPazientiKey As Integer, inMeCaption As String)
    If intPazientiKey <> inPazientiKey Then
        intNumeroFormAperti = 1
        listFormAperti.Refresh
        listFormAperti.Add (inMeCaption)
        intPazientiKey = inPazientiKey
        Call CaricaInfoPaziente
    Else
        If Not listFormAperti.IsPresent(inMeCaption) Then
            intNumeroFormAperti = intNumeroFormAperti + 1
            listFormAperti.Add (inMeCaption)
        End If
    End If
End Sub

Private Sub CaricaInfoPaziente()
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT NOME, COGNOME FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        strPazienteCognome = rsDataset("COGNOME")
        strPazienteNome = rsDataset("NOME")
    Else
        strPazienteCognome = ""
        strPazienteNome = ""
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

