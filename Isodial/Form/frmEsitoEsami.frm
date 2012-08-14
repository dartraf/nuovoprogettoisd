VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{AAFB789A-EB36-45DC-A196-1802D8AA28C9}#3.0#0"; "DataTimeBox.ocx"
Begin VB.Form frmEsitoEsami 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CONSULTAZIONE ESAMI DI LABORATORIO"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar hscrBarra 
      Height          =   255
      LargeChange     =   3
      Left            =   280
      Max             =   6
      TabIndex        =   3
      Top             =   7400
      Visible         =   0   'False
      Width           =   12300
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmEsitoEsami.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblCognome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   25
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblEta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11880
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Index           =   1
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anni"
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
         Index           =   3
         Left            =   11280
         TabIndex        =   12
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   12855
      Begin VB.CheckBox chkFiltra 
         Height          =   270
         Left            =   12000
         Picture         =   "frmEsitoEsami.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Filtra esami effettuati"
         Top             =   580
         Width           =   375
      End
      Begin VB.ComboBox cboEsami 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   9615
      End
      Begin DataTimeBox.uDataTimeBox oData 
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   27
         Top             =   150
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
         Left            =   10200
         TabIndex        =   28
         Top             =   150
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   661
         DataBox         =   -1  'True
         TimeBox         =   0   'False
         VisibleElenca   =   0   'False
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
         Left            =   9600
         TabIndex        =   9
         Top             =   195
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gruppo di Esami"
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
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1740
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
         Left            =   360
         TabIndex        =   7
         Top             =   195
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   12855
      Begin VB.TextBox txtAppo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   10186
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         MousePointer    =   15
         FormatString    =   "|| Descrizioni esame           | PN | Unità Misura    | Min    | Max    "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmEsitoEsami.frx":05A3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   5055
      Begin VB.Label Label2 
         Caption         =   "Valori nella norma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   260
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3240
         Shape           =   1  'Square
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1560
         Shape           =   1  'Square
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   120
         Shape           =   1  'Square
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Al di sotto del valore minimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   260
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Al di sopra del valore massimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   16
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   5160
      TabIndex        =   4
      Top             =   7680
      Width           =   7815
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Annulla Digitazione"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "&Elimina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   19
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmEsitoEsami"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' rs della scheda
Dim rsEsami As Recordset
Dim stoPulendo As Boolean
Dim stoCaricando As Boolean
Dim lettera As String
Dim vCol As Integer
Dim vRow As Integer
'' oggetto che gestisce l'annullamento dei dati nelle flx
Dim objAnnulla As CAnnulla
'' rs per la tracciatura
Dim rsDisco As Recordset
Dim intPazientiKey As Integer

Const icsPN As String = "  X"

'' Ricarica la cbo
Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    stoCaricando = True
    Call Filtra
    stoCaricando = False
    
    Select Case CaricaPazienteInAperturaForm(Me.Caption, False, intPazientiKey)
        Case tpTrovaPaziente
            Call TrovaPaziente
        Case tpCaricaPaziente
            Call CaricaPaziente
    End Select
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
    
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    stoPulendo = True
    oData(0).data = date - 155
    oData(1).data = date
    stoPulendo = False
    With flxGriglia
        ' flxGriglia.TextMatrix(0,1) contiene il key del gruppo esame
        ' flxGriglia.TextMatrix(i,1) contiene il key della tab VOCI_ESAMI
        ' la colonna 0 non viene utilizzata
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .Rows = 1
        .ColAlignment(2) = vbLeftJustify
        .Row = 0
        For i = 2 To 6
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexDefault
    End With
    For i = 0 To 1
        oData(i).ConnectionString = strConnectionStringCentro
    Next i
    Call ApriRsDisconnesso
    Set objAnnulla = New CAnnulla
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oPazientiKey.OnClosingForm (Me.Caption)
    intPazientiKey = 0
End Sub

Private Sub TrovaPaziente()
    cmdTrova_Click
    If tTrova.keyReturn = 0 Then
        Unload Me
    End If
End Sub

'' Apre il recordset disconnesso per la tracciatura
Private Sub ApriRsDisconnesso()
    Dim i As Integer
    Dim strSql As String
    
    strSql = "SELECT     ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, CODICE_ESAME, VALORE " & _
            "FROM       (ANAMNESI_ESAMI " & _
            "           INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY)"
    Dim rsDataset As New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set rsDisco = New ADODB.Recordset
    For i = 0 To rsDataset.Fields.count - 1
        rsDisco.Fields.Append rsDataset.Fields(i).Name, rsDataset.Fields(i).Type, rsDataset.Fields(i).DefinedSize, rsDataset.Fields(i).Attributes
    Next i
    rsDisco.CursorLocation = adUseClient
    rsDisco.Open , , adOpenDynamic, adLockOptimistic
    Set rsDataset = Nothing
End Sub

'' Confronta i campi per rilevare le eventuali modifiche
' e le salva nella relativa tabella delle modifiche
Private Sub Confronta()
    Dim rsDataset As Recordset
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    Dim valori As String
    Dim trovato As Boolean
    
    ' filtra per la presenza di piu record
    rsDisco.Filter = " (ESAMI_LABKEY=" & rsEsami("ESAMI_LABKEY") & ")"
    If rsDisco("VALORE") <> rsEsami("VALORE") Then
        trovato = True
    Else
        trovato = False
    End If
    If trovato Then
        valori = VirgolaOrPunto(rsDisco("VALORE"), ",")
        ' aggiorna il rsDisco
        rsDisco("VALORE") = rsEsami("VALORE")
        v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE", "CODICE_RECORD", "DATA_RECORD", "CODICE_ESAME", "CODICE_VOCE", "VALORE")
        v_Val = Array(tAccesso.key, date, Time, intPazientiKey, rsEsami("ESAMI_LABKEY"), flxGriglia.TextMatrix(0, vCol), flxGriglia.TextMatrix(0, 1), flxGriglia.TextMatrix(vRow, 1), valori)
        Set rsDataset = New Recordset
        rsDataset.Open "M_ESAMI_LAB", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.AddNew v_Nomi, v_Val
        rsDataset.Update
        Set rsDataset = Nothing
    End If
End Sub

'' Permette il funzionamento della rotellina del mouse nella flx
'Public Sub MouseWheel(flx As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'    Dim NewValue As Long
'    Dim Lstep As Single

'    On Error Resume Next
'    With flx
'        Lstep = .Height / .RowHeight(0)
'        Lstep = Int(Lstep)
'        If Lstep < 10 Then
'            Lstep = 10
'        End If
'        If Rotation > 0 Then
'            NewValue = .TopRow - Int(Lstep / 3)
'            If NewValue < 1 Then
'                NewValue = 1
'            End If
'        Else
'            NewValue = .TopRow + Int(Lstep / 3)
'            If NewValue > .Rows - 1 Then
'                NewValue = .Rows - 1
'            End If
'        End If
'        .TopRow = NewValue
'    End With
'End Sub
'-------------------------------


'' Pulisce l'intera scheda
Private Sub PulisciTutto()
    stoPulendo = True
    intPazientiKey = 0
    oData(0).Pulisci
    oData(1).Pulisci
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    flxGriglia.Cols = 7
    lblCognome = ""
    lblNome = ""
    lblEta = ""
    cboEsami.ListIndex = -1
    stoPulendo = False
    hscrBarra.Visible = False
    flxGriglia.Height = 5775
    cmdTrova.SetFocus
End Sub

'' Verifica la presenza di tutti i dati necessati per caricare la scheda
Private Function Completo(Optional Stampa As Boolean = False) As Boolean
    Completo = False
    If intPazientiKey = 0 Then
        MsgBox "Selezionere il paziente", vbCritical, "Attenzione"
        Exit Function
    End If
    If cboEsami.ListIndex = -1 And Not Stampa Then
        MsgBox "Selezionare il gruppo di esami di laboratorio", vbCritical, "Attenzione"
        Exit Function
    End If
    If oData(0).data = "" Or oData(1).data = "" Then
        MsgBox "Inserire le date", vbCritical, "Attenzione"
        Exit Function
    End If
    Completo = True
End Function

'' Restituisce il numero di riga dove è presente il codiceEsame
Private Function getRigaEsame(codiceEsame As Integer) As Integer
    Dim i As Integer
    For i = 1 To flxGriglia.Rows - 1
        If flxGriglia.TextMatrix(i, 1) = codiceEsame Then
            getRigaEsame = i
            Exit Function
        End If
    Next i
End Function

'' Filtra i soli gruppi di esami per il quale il paziente ha esami per le date selezionate
Private Sub Filtra()
    Dim data_min As Date
    Dim data_max As Date
    Dim filtroData As String
    
    If chkFiltra.Value = Checked Then
        If oData(0).data <> "" And oData(1).data <> "" Then
            data_min = oData(0).DataAmericana
            data_max = oData(1).DataAmericana
            filtroData = " AND DATA BETWEEN #" & data_min & "# AND #" & data_max & "#"
        End If
        
        Call RicaricaComboBox("SELECT DISTINCT G.NOME, G.KEY FROM (ANAMNESI_ESAMI A INNER JOIN GRUPPI_ESAMI G ON G.KEY=A.CODICE_GRUPPO) WHERE CODICE_PAZIENTE=" & intPazientiKey & filtroData, "NOME", cboEsami)
    Else
        Call RicaricaComboBox("GRUPPI_ESAMI", "NOME", cboEsami)
    End If
End Sub

'' Carica la scheda con gli esami del gruppo di esami e tra le due date selezionate
Private Sub CaricaScheda()
    Dim i As Integer
    Dim data_min As Date
    Dim data_max As Date
    Dim dataAppo As Date
    Dim rigaEsame As Integer
    Dim strSql As String
    
    If cboEsami.ListIndex = -1 Then Exit Sub
        
    With flxGriglia
        ' pulisce
        .Rows = 1
        .Cols = 7
        vRow = 0
        vCol = 0
        
        flxGriglia.TextMatrix(0, 1) = cboEsami.ItemData(cboEsami.ListIndex)
        ' carica le voci
        Set rsEsami = New Recordset
        rsEsami.Open "SELECT V.NOME, V.KEY, UNITA, MIN, MAX, PN FROM (ASSOCIAZIONE_ESAMI_LAB A INNER JOIN VOCI_ESAMI V ON V.KEY=A.CODICE_ESAME) WHERE CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 1) & " ORDER BY ORDINE_VISUALIZZAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsEsami.EOF
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = rsEsami("KEY")
                .TextMatrix(.Rows - 1, 2) = rsEsami("NOME")
                If CBool(rsEsami("PN")) = True Then
                    .TextMatrix(.Rows - 1, 3) = icsPN
                Else
                    .TextMatrix(.Rows - 1, 4) = rsEsami("UNITA")
                    .TextMatrix(.Rows - 1, 5) = rsEsami("MIN")
                    .TextMatrix(.Rows - 1, 6) = rsEsami("MAX")
                End If
            End With
            rsEsami.MoveNext
        Loop
        rsEsami.Close
    End With
    
    ' carica le date
    data_max = oData(1).DataAmericana
    data_min = oData(0).DataAmericana
        
    strSql = "SELECT    ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, VALORE, CODICE_ESAME " & _
             "FROM      (ANAMNESI_ESAMI " & _
             "          INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) " & _
             "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
             "          CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 1) & " AND " & _
             "          DATA BETWEEN #" & data_min & "# AND #" & data_max & "# " & _
             "ORDER BY  DATA DESC"
    rsEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsEsami.EOF And rsEsami.BOF) Then
        flxGriglia.Redraw = False
        
        ' pulisce rsDisco
        Do While Not rsDisco.EOF
            rsDisco.Delete
            rsDisco.MoveNext
        Loop
        dataAppo = rsEsami("DATA")
        With flxGriglia
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 1150
            .TextMatrix(0, .Cols - 1) = Format(rsEsami("DATA"), "dd/mm/yy")
            Do While Not rsEsami.EOF
                If rsEsami("DATA") <> dataAppo Then
                    ' aggiunge una nuova colonna
                    .Cols = .Cols + 1
                    .TextMatrix(0, .Cols - 1) = Format(rsEsami("DATA"), "dd/mm/yy")
                    .ColWidth(.Cols - 1) = 1150
                End If
                dataAppo = rsEsami("DATA")
                rigaEsame = getRigaEsame(rsEsami("CODICE_ESAME"))
                .TextMatrix(rigaEsame, 0) = rsEsami("ESAMI_LABKEY")
                ' aggiunge i valori
                Select Case rsEsami("VALORE")
                    Case -3
                        .TextMatrix(rigaEsame, .Cols - 1) = ""
                    Case -2
                        .Col = .Cols - 1
                        .Row = rigaEsame
                        .CellForeColor = vbRed
                        .TextMatrix(rigaEsame, .Cols - 1) = "NEGATIVO"
                    Case -1
                        .Col = .Cols - 1
                        .Row = rigaEsame
                        .CellForeColor = vbRed
                        .TextMatrix(rigaEsame, .Cols - 1) = "POSITIVO"
                    Case Else
                        .TextMatrix(rigaEsame, .Cols - 1) = VirgolaOrPunto(rsEsami("VALORE"), ",") & ""
                        ' imposta il colore di avvertimento
                        If rsEsami("VALORE") <> "" Then
                            Call ColoreDiAvviso(flxGriglia, rigaEsame, .Cols - 1, VirgolaOrPunto(rsEsami("VALORE"), "."), .TextMatrix(rigaEsame, 6), .TextMatrix(rigaEsame, 5))
                        End If
                End Select
                
                ' aggiorna i dati nel rsDisco
                rsDisco.AddNew
                For i = 0 To rsDisco.Fields.count - 1
                    rsDisco.Fields(i) = rsEsami.Fields(i)
                Next i
                rsDisco.Update
                
                rsEsami.MoveNext
            Loop
            ' imposta il grassetto a tutta la prima riga
            .Row = 0
            For i = 7 To .Cols - 1
                .Col = i
                .CellFontBold = True
            Next i
            ' verifica se attivare la barra orizzontale
            If .Cols > 11 Then
                hscrBarra.Visible = True
                flxGriglia.Height = 5535
                hscrBarra.max = .Cols - 8 - 1
                hscrBarra.min = 0
                hscrBarra.Value = 0
            Else
                hscrBarra.Visible = False
                flxGriglia.Height = 5775
            End If
            ' azzera
            .Col = 0
        End With
        
        flxGriglia.Redraw = True
    Else
        flxGriglia.Rows = 1
        hscrBarra.Visible = False
        flxGriglia.Height = 5775
    End If
    rsEsami.Close
    Set rsEsami = Nothing
End Sub

'' Carica la scheda o pulisce la flx
Private Sub Cerca()
    If Not (intPazientiKey = 0 Or cboEsami.ListIndex = -1 Or oData(0).data = "" Or oData(1).data = "") Then
        ' pulisce l'oggetto
        objAnnulla.Refresh
        cmdAnnulla.Enabled = False
        Call CaricaScheda
    Else
        ' pulisce le flx
        flxGriglia.Rows = 1
        flxGriglia.Cols = 7
    End If
End Sub

'' Carica i dati del paziente
Private Sub CaricaPaziente()
    Dim rsDataset As Recordset
    If intPazientiKey = 0 Then
        ' pulisce la griglia
        ' pulisce la flx azzerando le righe
        flxGriglia.Rows = 1
        ' e le colonne
        flxGriglia.Cols = 7
    Else
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA,CODICE_ID FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        lblCognome = rsDataset("COGNOME")
        lblNome = rsDataset("NOME")
        Dim somma As Integer
        If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
            somma = -1
        ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
            somma = -1
        Else
            somma = 0
        End If
        lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
        Set rsDataset = Nothing
    
        Call oPazientiKey.ImpostaPazientiKey(intPazientiKey, Me.Caption)
    
        oData(0).data = date - 155
        oData(1).data = date
        Call CaricaScheda
    End If
End Sub

'' Salva le modifiche o inserisce nuovi record
Private Sub SalvaModifiche(ByRef outRefresh As Boolean)
    Dim i As Integer
    Dim v_Nomi(1 To 5) As Variant
    Dim v_Val(1 To 5) As Variant
    Dim valore As Variant
    Dim numKey As Integer
    Dim data As Date
    Dim data_max As Date
    Dim data_min As Date
    Dim strSql As String
    
    With flxGriglia
        Set rsEsami = New Recordset
        data = DateValue(Month(.TextMatrix(0, vCol)) & "/" & Day(.TextMatrix(0, vCol)) & "/" & Year(.TextMatrix(0, vCol)))
        If .TextMatrix(vRow, vCol) <> "" Then
            data_max = oData(1).DataAmericana
            data_min = oData(0).DataAmericana
               
            strSql = "SELECT    ANAMNESI_ESAMI.* " & _
                     "FROM      ANAMNESI_ESAMI " & _
                     "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                     "          DATA=#" & data & "# AND " & _
                     "          CODICE_GRUPPO=" & .TextMatrix(0, 1)
            rsEsami.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsEsami("UTENTE_MODIFICATORE") = tAccesso.key
            numKey = rsEsami("KEY")
            rsEsami.Update
            rsEsami.Close

            Select Case .TextMatrix(vRow, vCol)
                Case Is = "POSITIVO"
                    valore = -1
                Case Is = "NEGATIVO"
                    valore = -2
                Case Else
                    valore = .TextMatrix(vRow, vCol)
            End Select
            
            strSql = "SELECT    ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, VALORE, CODICE_ESAME " & _
                     "FROM      (ANAMNESI_ESAMI " & _
                     "          INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                     "WHERE     ANAMNESI_ESAMI.KEY=" & numKey & " AND " & _
                     "          CODICE_ESAME=" & .TextMatrix(vRow, 1)
            rsEsami.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If Not (rsEsami.EOF And rsEsami.BOF) Then
                rsEsami.Update "VALORE", valore
                rsEsami.Close
                
                strSql = "SELECT    ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, VALORE, CODICE_ESAME " & _
                         "FROM      (ANAMNESI_ESAMI " & _
                         "          INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) " & _
                         "WHERE     ANAMNESI_ESAMI.KEY=" & numKey & " AND " & _
                         "          ESAMI_LAB.CODICE_ESAME=" & .TextMatrix(vRow, 1)
                rsEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockOptimistic, adCmdText
                If TRACCIATO Then
                    Call Confronta
                End If
                rsEsami.Close
            Else
                v_Nomi(1) = "KEY"
                v_Nomi(2) = "CODICE_ESAME"
                v_Nomi(3) = "VALORE"
                v_Nomi(4) = "CODICE_ANAMNESI_ESAMI"
                v_Val(1) = GetNumero("ESAMI_LAB")
                v_Val(2) = flxGriglia.TextMatrix(vRow, 1)
                v_Val(3) = valore
                v_Val(4) = numKey
                rsEsami.Close
                rsEsami.Open "ESAMI_LAB", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
                rsEsami.AddNew
                For i = 1 To 4
                    rsEsami(v_Nomi(i)) = v_Val(i)
                Next
                rsEsami.Update
                rsEsami.Close

                ' pulisce rsDisco
                Do While Not rsDisco.EOF
                    rsDisco.Delete
                    rsDisco.MoveNext
                Loop
                ' aggiorna i dati nel rsDisco
                strSql = "SELECT    ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, VALORE, CODICE_ESAME " & _
                         "FROM      (ANAMNESI_ESAMI " & _
                         "          INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) " & _
                         "WHERE     CODICE_PAZIENTE=" & intPazientiKey & " AND " & _
                         "          CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 1) & " AND " & _
                         "          DATA BETWEEN #" & data_min & "# AND #" & data_max & "#"
                rsEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                Do While Not rsEsami.EOF
                    ' aggiorna i dati nel rsDisco
                    rsDisco.AddNew
                    For i = 0 To rsDisco.Fields.count - 1
                        rsDisco.Fields(i) = rsEsami.Fields(i)
                    Next i
                    rsDisco.Update
                    rsEsami.MoveNext
                Loop
                rsEsami.Close
            End If
        Else
            Dim intEsamiLabKey As Integer
            Dim intAnamnesiEsamiKey As Integer
            Dim blnEliminaAnamnesi As Boolean
            
            ' verifica se esisteva un record con valore
            strSql = "SELECT    ANAMNESI_ESAMI.*, ESAMI_LAB.KEY AS ESAMI_LABKEY, VALORE, CODICE_ESAME " & _
                     "FROM      (ANAMNESI_ESAMI " & _
                     "          INNER JOIN ESAMI_LAB ON ESAMI_LAB.CODICE_ANAMNESI_ESAMI=ANAMNESI_ESAMI.KEY) " & _
                     "WHERE     (CODICE_PAZIENTE=" & intPazientiKey & ") AND " & _
                     "          (DATA=#" & data & "#) AND " & _
                     "          (CODICE_GRUPPO=" & .TextMatrix(0, 1) & ") AND " & _
                     "          (ESAMI_LAB.CODICE_ESAME=" & .TextMatrix(vRow, 1) & ")"
            rsEsami.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            If Not (rsEsami.EOF And rsEsami.BOF) Then
                intEsamiLabKey = rsEsami("ESAMI_LABKEY")
                intAnamnesiEsamiKey = rsEsami("ANAMNESI_ESAMI.KEY")
            End If
            rsEsami.Close
            If intEsamiLabKey <> 0 Then
                rsEsami.Open "Select * From ESAMI_LAB where Key=" & intEsamiLabKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                If Not (rsEsami.EOF And rsEsami.BOF) Then
                    rsEsami.Delete
                End If
                rsEsami.Close
                
                ' verifica se l'anamnesi esami è rimasta senza esamilab
                rsEsami.Open "Select * From ESAMI_LAB where CODICE_ANAMNESI_ESAMI=" & intAnamnesiEsamiKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                If (rsEsami.EOF And rsEsami.BOF) Then
                    blnEliminaAnamnesi = True
                End If
                rsEsami.Close
                
                If blnEliminaAnamnesi Then
                    rsEsami.Open "Select * From ANAMNESI_ESAMI where Key=" & intAnamnesiEsamiKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
                    If Not (rsEsami.EOF And rsEsami.BOF) Then
                        rsEsami.Delete
                        outRefresh = True
                    End If
                    rsEsami.Close
                End If
            End If
        End If
        Set rsEsami = Nothing
    End With
End Sub

'' Salva l'eliminazione nel db di tracciature
Private Sub SalvaEliminazione()
    Dim v_nome As Variant
    Dim v_Val As Variant
    Dim massimo As Integer
    Dim i As Integer
    
    Dim rsDataset As New Recordset
    v_nome = Array("CODICE_UTENTE", "DATA", "ORA", "CODICE_PAZIENTE")
    v_Val = Array(tAccesso.key, date, Time, intPazientiKey)
    rsDataset.Open "E_ESAMI_LAB", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_nome, v_Val
    rsDataset.Update
    rsDataset.Close
    rsDataset.Open "SELECT MAX(KEY) AS MASSIMO FROM E_ESAMI_LAB", cnTrac, adOpenKeyset, adLockReadOnly, adCmdText
    massimo = rsDataset("MASSIMO")
    rsDataset.Close
    
    v_nome = Array("DATA_ESAME", "CODICE_ESAME", "CODICE_ELIMINAZIONE")
    rsDataset.Open "INFO_ESAMI_LAB", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    For i = 1 To flxGriglia.Cols - 7
        v_Val = Array(flxGriglia.TextMatrix(0, 6 + i), flxGriglia.TextMatrix(0, 1), massimo)
        rsDataset.AddNew v_nome, v_Val
        rsDataset.Update
    Next i
    Set rsDataset = Nothing
End Sub

Private Sub cmdAnnulla_Click()
    Dim Dato As String
    Dim Col As Integer
    Dim Row As Integer
    Dato = objAnnulla.Dato
    Col = objAnnulla.Col
    Row = objAnnulla.Row
    ' cerca la riga con il key memorizzato in rowkey
    With flxGriglia
        ' annulla
        .TextMatrix(Row, Col) = Dato
        vRow = Row
        vCol = Col
        objAnnulla.Remove
        ' cambia colore
        ' nel caso il dato sia cancellato
        Dato = IIf(Dato = "", -1, Dato)
        If Dato = "POSITIVO" Or Dato = "NEGATIVO" Then Dato = -1
        Call ColoreDiAvviso(flxGriglia, Row, Col, VirgolaOrPunto(Dato, "."), IIf(.TextMatrix(vRow, 3) = icsPN, "0", .TextMatrix(Row, 6)), IIf(.TextMatrix(vRow, 3) = icsPN, "0", .TextMatrix(Row, 5)))
        ' modifica anche il db
        Call SalvaModifiche(0)
        If objAnnulla.Vuoto = True Then
            cmdAnnulla.Enabled = False
        End If
    End With
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

'' Elimina gli esami di tutto il periodo selezionato
Private Sub cmdElimina_Click()
    Dim data_max As Date
    Dim data_min As Date
    Dim eliminato As Boolean
    Dim numKey As Integer
    Dim cmCommand As New Command
    
    If Not Completo Then Exit Sub
    If MsgBox("Sei sicuro di voler eliminare gli esami di laboratorio del paziente: " & UCase(lblCognome) & " " & UCase(lblNome) & " ?" & vbCrLf & "Dal " & oData(0).data & " al " & oData(1).data, vbQuestion + vbYesNo, "Eliminazione") = vbYes Then
        Set rsEsami = New Recordset
        ' carica le date
        data_max = oData(1).DataAmericana
        data_min = oData(0).DataAmericana
        cmCommand.ActiveConnection = cnPrinc
        cmCommand.CommandType = adCmdText
    
        rsEsami.Open "SELECT * FROM ANAMNESI_ESAMI WHERE CODICE_PAZIENTE=" & intPazientiKey & " AND CODICE_GRUPPO=" & flxGriglia.TextMatrix(0, 1) & " AND DATA BETWEEN #" & data_min & "# AND #" & data_max & "#", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
        Do While Not rsEsami.EOF
            numKey = rsEsami("KEY")
            rsEsami.Delete
        
            cmCommand.CommandText = "DELETE * FROM ESAMI_LAB WHERE CODICE_ANAMNESI_ESAMI=" & numKey
            cmCommand.Execute
            
            rsEsami.MoveNext
            eliminato = True
        Loop
        rsEsami.Close
        Set rsEsami = Nothing
        
        If eliminato And TRACCIATO Then
            Call SalvaEliminazione
        End If
        Call PulisciTutto
        MsgBox "Cancellazione effettuata", vbInformation, "Eliminazione"
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim condizione As String
    Dim data_min As Date
    Dim data_max As Date
    
    data_min = oData(0).DataAmericana
    data_max = oData(1).DataAmericana
    condizione = " AND ANAMNESI_ESAMI.DATA BETWEEN #" & data_min & "# AND #" & data_max & "# "
    
    If intPazientiKey = 0 Then
        MsgBox "Selezionare il paziente", vbInformation, "Impossibile stampare"
        Exit Sub
    End If
      
    Set rsEsami = New Recordset
    rsEsami.Open "SELECT COGNOME, NOME, DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    structIntestazione.sPaziente = rsEsami("COGNOME") & " " & rsEsami("NOME")
    structIntestazione.sDataPaziente = rsEsami("DATA_NASCITA")
    Set rsEsami = Nothing
    
    Call StampaSestaParte(False, intPazientiKey, condizione)
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    Call PulisciTutto
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn = 0 Then
        Unload Me
    Else
        intPazientiKey = tTrova.keyReturn
        Call CaricaPaziente
    End If
End Sub

Private Sub chkFiltra_Click()
    If intPazientiKey <> 0 Then
        ' pulisce
        flxGriglia.Rows = 1
        flxGriglia.Cols = 7
        stoCaricando = True
        Call Filtra
        stoCaricando = False
    End If
End Sub

Private Sub cboEsami_Click()
    If stoPulendo Or stoCaricando Then Exit Sub
    Call Cerca
End Sub

Private Sub hscrBarra_Change()
    Dim i As Integer
    With flxGriglia
        For i = 1 To hscrBarra.Value
            .ColWidth(i + 7 - 1) = 0
        Next i
        For i = hscrBarra.Value + 1 To hscrBarra.max
            .ColWidth(i + 7 - 1) = .ColWidth(.Cols - 1)
        Next i
        .SetFocus
    End With
End Sub

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If flxGriglia.Rows = 1 Then Exit Sub
    If flxGriglia.Row = flxGriglia.Rows - 1 Then
        i = 1
    Else
        i = flxGriglia.Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia.TextMatrix(i, 2), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia.Row = i
            If i >= 16 Or flxGriglia.TopRow > 16 Then
                flxGriglia.TopRow = i
            End If
            Call ColoraFlx(flxGriglia, 6)
            Exit Do
        End If
        If i = flxGriglia.Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxGriglia.Row
End Sub

'Private Sub flxGriglia_GotFocus()
'    Call WheelHook(Me, flxGriglia)
'End Sub

'Private Sub flxGriglia_LostFocus()
'    Call WheelUnHook
'End Sub
'---------------------------


Private Sub flxGriglia_Click()
    vCol = flxGriglia.Col
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, 6, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, 6)
        vRow = flxGriglia.Row
    End If
End Sub

Private Sub flxGriglia_DblClick()
    ' permette anche la modifica dei dati
    With flxGriglia
        .SetFocus
        If VerificaClickFlx(flxGriglia) = False Then Exit Sub
        If .Col <= 5 Then Exit Sub
        ' verifica se la voce accetta valori pos e neg
        If AccettaPN(.TextMatrix(.Row, 2)) Then
            Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, vRow)
            cmdAnnulla.Enabled = True
            Call GestisciPN(flxGriglia, .Col, True)
            Call SalvaModifiche(0)
        Else
            txtAppo.Left = .colPos(.Col) + .Left + 45
            txtAppo.Top = .rowPos(.Row) + .Top + 45
            txtAppo.Text = .TextMatrix(.Row, .Col)
            txtAppo.Width = .ColWidth(.Col)
            txtAppo.Visible = True
            txtAppo.SetFocus
        End If
    End With
End Sub

Private Sub flxGriglia_Scroll()
    If txtAppo.Visible Then
        txtAppo.Top = flxGriglia.rowPos(flxGriglia.Row) + flxGriglia.Top + 45
    End If
End Sub

Private Sub oData_OnDataClick(Index As Integer)
    oData(Index).Pulisci
End Sub

Private Sub oData_OnDataChange(Index As Integer)
    If stoPulendo Then Exit Sub
    If oData(Index).data <> "" Then
        stoCaricando = True
        Call Filtra
        stoCaricando = False
        Call Cerca
    Else
        ' pulisce la griglia
        ' pulisce la flx azzerando le righe
        flxGriglia.Rows = 1
        ' e le colonne
        flxGriglia.Cols = 7
    End If
End Sub

Private Sub txtAppo_Change()
    If Not (lettera = "." Or lettera = "") Then
        Call OnlyNumber(txtAppo, lettera)
    End If
End Sub

Private Sub txtAppo_GotFocus()
    txtAppo.SelStart = 0
    txtAppo.SelLength = Len(txtAppo)
End Sub

Private Sub txtAppo_KeyPress(KeyAscii As Integer)
    ' quando inserisce la virgola(44) cambia con il punto(46)
    If KeyAscii = 44 Then KeyAscii = 46
    lettera = Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        flxGriglia.SetFocus
    End If
End Sub

Private Sub txtAppo_LostFocus()
    Dim valPassato As Single
    Dim blnRefresh As Boolean
    
    txtAppo.Visible = False
    If UCase(flxGriglia.TextMatrix(vRow, vCol)) <> UCase(txtAppo) Then
        Call objAnnulla.Add(flxGriglia.TextMatrix(vRow, vCol), vCol, vRow)
        cmdAnnulla.Enabled = True
        flxGriglia.TextMatrix(vRow, vCol) = txtAppo.Text
        Call SalvaModifiche(blnRefresh)
        If blnRefresh Then
            Call CaricaScheda
        Else
            ' imposta il colore di sfondo
            With flxGriglia
                If .TextMatrix(vRow, vCol) <> "" Then
                    valPassato = VirgolaOrPunto(.TextMatrix(vRow, vCol), ".")
                Else
                    valPassato = -1
                End If
                Call ColoreDiAvviso(flxGriglia, vRow, vCol, valPassato, .TextMatrix(vRow, 6), .TextMatrix(vRow, 5))
            End With
        End If
    End If
End Sub

Private Sub txtAppo_Validate(Cancel As Boolean)
    If txtAppo = "" Then
        Cancel = False
    Else
        Cancel = ControlloNumerico(txtAppo.Text)
    End If
End Sub

