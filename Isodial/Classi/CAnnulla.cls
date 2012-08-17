VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnnulla"
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
' <b>Tipo e nome modulo</b>:        Modulo di classe - CAnnulla.cls
'
' <b>Descrizione</b>: Classe per la gestione della pila Annulla
'
' @remarks utilizzata solo per le grigle
'
' @author
'
' @date 28/01/2011 18.27
Option Explicit

Private mNodo() As CNodo
Private mVuoto As Boolean
Private Const maxNodi As Byte = 5

''
' Inizializza la pila
'
' @param
' @param
' @return
' @remarks
Private Sub Class_Initialize()
    ReDim mNodo(0)
    Set mNodo(0) = New CNodo
    mVuoto = True
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
' Aggiunge un nuovo nodo alla pila
'
' @param vDato dato da aggiungere
' @param vCol colonna dove è presente il dato
' @param vRow riga dove è presente il dato
' @return
' @remarks
Public Sub Add(ByVal vDato As String, ByVal vCol As Integer, ByVal vRow As Integer)
    Dim dimen As Integer
    Dim i As Integer
    mVuoto = False
    dimen = UBound(mNodo)
    If dimen >= maxNodi Then
        ' richiama lo shift
        For i = 1 To maxNodi - 1
            mNodo(i).Dato = mNodo(i + 1).Dato
            mNodo(i).Col = mNodo(i + 1).Col
            mNodo(i).Row = mNodo(i + 1).Row
            'Debug.Print "       shift  in pos= " & i & "  dato= " & mNodo(i).Dato & "  col= " & mNodo(i).Col & "  row= " & mNodo(i).Row
        Next i
        ' elimina l'ultimo posto che è uguale al penultimo
        ReDim Preserve mNodo(dimen - 1)
        dimen = dimen - 1
    End If
    ' lo aggiunge
    ReDim Preserve mNodo(dimen + 1)
    Set mNodo(dimen + 1) = New CNodo
    dimen = dimen + 1
    mNodo(dimen).Dato = vDato
    mNodo(dimen).Col = vCol
    mNodo(dimen).Row = vRow
    ' Debug.Print "add  in dimen= " & dimen & "  dato= " & vDato & "  col= " & vCol & "  row= " & vRow
End Sub

''
' Rimuove un nodo dalla pila
'
' @param
' @param
' @return
' @remarks
Public Sub Remove()
    Dim dimen As Integer
    
    dimen = UBound(mNodo)
    If dimen <> 0 Then
        ReDim Preserve mNodo(dimen - 1)
        If dimen - 1 = 0 Then
            mVuoto = True
        End If
    End If
End Sub

''
' Pulisce la pila
'
' @param
' @param
' @return
' @remarks
Public Sub Refresh()
    ReDim Preserve mNodo(0)
    mVuoto = True
End Sub

Public Property Get Col() As Integer
    Col = mNodo(UBound(mNodo)).Col
End Property

Public Property Get Row() As Integer
    Row = mNodo(UBound(mNodo)).Row
End Property

Public Property Get Dato() As String
    Dato = mNodo(UBound(mNodo)).Dato
End Property

Public Property Get Vuoto() As Boolean
    Vuoto = mVuoto
End Property

