VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNodo() As CNodo
Private mVuoto As Boolean

Private Sub Class_Initialize()
    ReDim mNodo(0)
    Set mNodo(0) = New CNodo
    mVuoto = True
End Sub

Private Sub Class_Terminate()
    ReDim mNodo(0)
    Set mNodo(0) = Nothing
End Sub

Public Sub Add(ByVal inDato As String)
    Dim dimen As Integer
    Dim i As Integer
    mVuoto = False
    dimen = UBound(mNodo)

    ReDim Preserve mNodo(dimen + 1)
    Set mNodo(dimen + 1) = New CNodo
    dimen = dimen + 1
    mNodo(dimen).Dato = inDato
End Sub

Public Function IsPresent(inDato As String) As Boolean
    Dim i As Integer
    
    IsPresent = False
    For i = 1 To UBound(mNodo)
        If UCase(mNodo(i).Dato) = UCase(inDato) Then
            IsPresent = True
            Exit For
        End If
    Next i
End Function

Public Sub RemoveLast()
    Dim dimen As Integer
    
    dimen = UBound(mNodo)
    If dimen <> 0 Then
        ReDim Preserve mNodo(dimen - 1)
        If dimen - 1 = 0 Then
            mVuoto = True
        End If
    End If
End Sub

Public Sub Remove(inDato As String)
    Dim oNodoAppo() As New CNodo
    Dim i As Integer
    Dim intCont As Integer
    
    If UBound(mNodo) = 0 Then Exit Sub
    ReDim oNodoAppo(UBound(mNodo))
    For i = 1 To UBound(mNodo)
        oNodoAppo(i).Dato = mNodo(i).Dato
    Next i
    
    ReDim mNodo(UBound(mNodo) - 1)
    intCont = 1
    For i = 1 To UBound(oNodoAppo)
        If UCase(oNodoAppo(i).Dato) <> UCase(inDato) Then
            Set mNodo(intCont) = New CNodo
            mNodo(intCont).Dato = oNodoAppo(i).Dato
            intCont = intCont + 1
        End If
    Next i

End Sub

Public Sub Refresh()
    ReDim Preserve mNodo(0)
    mVuoto = True
End Sub

Public Property Get Dato() As String
    Dato = mNodo(UBound(mNodo)).Dato
End Property

Public Property Get Vuoto() As Boolean
    Vuoto = mVuoto
End Property


