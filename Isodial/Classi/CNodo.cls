VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNodo"
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
' <b>Tipo e nome modulo</b>:        Modulo di classe - CNodo.cls
'
' <b>Descrizione</b>: Classe Nodo della pila Annulla
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.29
Option Explicit

Private mDato As String
Private mCol As Integer
Private mRow As Integer

Public Property Let Dato(ByVal vdata As String)
    mDato = vdata
End Property

Public Property Get Dato() As String
    Dato = mDato
End Property

Public Property Let Col(ByVal vdata As Integer)
    mCol = vdata
End Property

Public Property Get Col() As Integer
    Col = mCol
End Property

Public Property Let Row(ByVal vdata As Integer)
    mRow = vdata
End Property

Public Property Get Row() As Integer
    Row = mRow
End Property
