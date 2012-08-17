VERSION 5.00
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmPannelloFiltroStato 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa pazienti"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cboStato 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmPannelloFiltroStato.frx":0000
         Left            =   1800
         List            =   "frmPannelloFiltroStato.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stato paziente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   760
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   5
         Top             =   765
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5295
      Begin VB.CommandButton cmdAnnulla 
         Cancel          =   -1  'True
         Caption         =   "&Annulla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmPannelloFiltroStato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmFiltro.frm
'
' <b>Descrizione</b>: Pannello utilizzato per filtrare i paziente da stampare in Stampa Lista Pazienti
'
' @remarks
'
' @author
'
' @date 07/02/2011 17.12
Option Explicit

Private Sub Form_Load()
    Call RicaricaComboBox("TIPO_STATO", "NOME", cboStato)
    cboStato.ListIndex = 0
    oData(0).data = "01/01/" & Year(Now)
    oData(1).data = date
End Sub

'' Verifica prima di mandare in stampa se i dati sono corretti
Private Function Completo() As Boolean
    Completo = False
    If oData(0).data <> "" Then
        If oData(1).data <> "" Then
            If CDate(oData(0).data) > CDate(oData(1).data) Then
                MsgBox "Inserire le date correttamente", vbCritical, "Attenzione"
                Exit Function
            End If
        Else
            MsgBox "Inserire entrambe le date", vbCritical, "Attenzione"
            Exit Function
        End If
    Else
        If oData(1).data <> "" Then
            MsgBox "Inserire entrambe le date", vbCritical, "Attenzione"
            Exit Function
        End If
    End If
    Completo = True
End Function

Private Sub cmdStampa_Click()
    If Completo Then
        If oData(0).data <> "" And oData(1).data <> "" Then
            tFiltroStato.isTutteLeDate = False
            tFiltroStato.dataDal = DateValue(Month(oData(0).data) & "/" & Day(oData(0).data) & "/" & Year(oData(0).data))
            tFiltroStato.dataAl = DateValue(Month(oData(1).data) & "/" & Day(oData(1).data) & "/" & Year(oData(1).data))
        Else
            tFiltroStato.isTutteLeDate = True
        End If
        tFiltroStato.statoPaziente = cboStato.ItemData(cboStato.ListIndex)
        Unload Me
    End If
End Sub

Private Sub cmdAnnulla_Click()
    tFiltroStato.statoPaziente = tpNoneStatoPaziente
    Unload Me
End Sub

Private Sub oData_LostFocus(Index As Integer)
    If oData(Index).txtBox = "" Then
        oData(Index).txtBox = "Tutte le date"
    End If
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

