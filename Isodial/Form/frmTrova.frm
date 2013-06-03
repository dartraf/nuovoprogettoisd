VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmTrova 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleziona il "
   ClientHeight    =   5460
   ClientLeft      =   750
   ClientTop       =   1710
   ClientWidth     =   7545
   Icon            =   "frmTrova.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox cboStato 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmTrova.frx":000C
         Left            =   4800
         List            =   "frmTrova.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdCerca 
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   120
         Picture         =   "frmTrova.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox txtCerca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   7335
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   "|  Cognome                            |  Nome                                 |  Data di nascita          "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   7335
      Begin VB.CommandButton cmdCambiaData 
         Caption         =   "&Cambio Data/Turno"
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
         Left            =   2420
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1340
      End
      Begin VB.CommandButton cmdNuovo 
         Caption         =   "&Nuovo"
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
         Left            =   3760
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAvanti 
         Caption         =   "&Seleziona"
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
         Left            =   4900
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdIndietro 
         Caption         =   "&Chiudi"
         CausesValidation=   0   'False
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
         Left            =   6140
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblVoci 
         AutoSize        =   -1  'True
         Caption         =   "Infermieri in elenco:  "
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2160
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCognome 
         Caption         =   "COGNOME"
      End
      Begin VB.Menu mnuNome 
         Caption         =   "NOME"
      End
   End
End
Attribute VB_Name = "frmTrova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDatasetCerca As Recordset
Dim testoVoce As String

Private Sub cmdCambiaData_Click()

 ' scelta = False
 ' caricato = False
' Call Select_Data
    Unload Me
 '   Exit Sub
    
 '   Unload Me
 '   frmPannelloPeriodo.Show 1
 '   periodo = frmPannelloPeriodo.GetPeriodo
 '   laData = frmPannelloPeriodo.getData
 '   Unload frmPannelloPeriodo
 '   If periodo = -1 Then
 '       Unload Me
 '   Else
 '       tTrova.Tipo = tpPAZIENTE
 '       tTrova.condizione = CreaCondizione
 '       tTrova.condStato = "(-1)"
 '       frmTrova.Show 1
 '       intPazientiKey = tTrova.keyReturn'

'        If tTrova.keyReturn = 0 Then
'           Unload Me
'        End If
'    End If
End Sub

Private Sub cmdNuovo_Click()
    tTrova.keyReturn = -1
    Unload Me
End Sub

Private Sub Form_Activate()
    Call Cerca
End Sub

'' Permette il funzionamento della rotellina del mouse nella flx
'Public Sub MouseWheel(flx As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
 '   Dim NewValue As Long
 '   Dim Lstep As Single

 '   On Error Resume Next
 '   With flx
 '       Lstep = .Height / .RowHeight(0)
 '       Lstep = Int(Lstep)
 '       If Lstep < 10 Then
 '           Lstep = 10
 '       End If
 '       If Rotation > 0 Then
 '           NewValue = .TopRow - Int(Lstep / 3)
 '           If NewValue < 1 Then
 '               NewValue = 1
 '           End If
 '       Else
 '           NewValue = .TopRow + Int(Lstep / 3)
 '           If NewValue > .Rows - 1 Then
 '               NewValue = .Rows - 1
 '           End If
 '       End If
 '       .TopRow = NewValue
 '   End With
'End Sub
'---------------------------------

Private Function isDialisi() As Boolean
    If (InStr(1, tTrova.condizione, "IN") > 0 Or InStr(1, tTrova.condizione, "KEY=-1")) And InStr(1, tTrova.condStato, "(-1)") Then
        isDialisi = True
    Else
        isDialisi = False
    End If
End Function

Private Sub Form_Load()
    Dim i As Integer
    
    tTrova.keyReturn = 0
    Select Case tTrova.Tipo
        Case tpMEDICOBASE
            flxGriglia.ColWidth(3) = 0
            Me.Caption = Me.Caption & "medico di base"
            testoVoce = "Medici in elenco: "
        Case tpPAZIENTE
            Me.Caption = Me.Caption & "paziente"
            If Not isDialisi Then
                testoVoce = "Pazienti in elenco: "
            Else
                testoVoce = "Pazienti in turno: "
            End If
        Case tpMEDICOREFER
            Me.Caption = Me.Caption & "medico refertante"
            flxGriglia.ColWidth(3) = 0
            testoVoce = "Medici in elenco: "
        Case tpINFERMIERE
            Me.Caption = Me.Caption & "infermiere"
            flxGriglia.ColWidth(3) = 0
            testoVoce = "Infermieri in elenco: "
        Case tpPSICOLOGI
            Me.Caption = Me.Caption & "psicologo"
            flxGriglia.ColWidth(3) = 0
            testoVoce = "Psicologi in elenco: "
        Case tpACCOMPAGNATORI
            Me.Caption = Me.Caption & "accompagnatori"
            flxGriglia.ColWidth(3) = 0
            testoVoce = "Accompagnatori in elenco: "
    End Select
    lblVoci = testoVoce
    With flxGriglia
        .ColWidth(0) = 0
        .ColAlignment(3) = vbLeftJustify
        .Row = 0
        For i = 1 To 3
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
    
    If tTrova.condStato = "" Then tTrova.condStato = "(-1) OR TRUE"
    
    Call RicaricaComboBox("SELECT KEY, NOME FROM TIPO_STATO WHERE KEY IN " & tTrova.condStato & " ORDER BY KEY", "NOME", cboStato)
    
    If cboStato.ListCount <> 0 And tTrova.Tipo = tpPAZIENTE Then
        cboStato.Visible = True
        cboStato.ListIndex = 0
        If InStr(1, tTrova.condStato, "-1") <> 0 And Not isDialisi Then
            cboStato.AddItem "Tutti"
            cboStato.ItemData(cboStato.NewIndex) = 0
        End If
    End If
    
    If tTrova.isOpenFromInfoGenerali Then
        cmdNuovo.Visible = True
    End If
    
    If tTrova.isOpenFromEsamiPrescriz Then
        cmdCambiaData.Visible = True
        cmdIndietro.Visible = False
    End If
End Sub

Private Function nomeTabella() As String
    Select Case tTrova.Tipo
        Case tpPAZIENTE: nomeTabella = "PAZIENTI"
        Case tpMEDICOBASE: nomeTabella = "MEDICI_BASE"
        Case tpMEDICOREFER: nomeTabella = "MEDICI_REFERTANTI"
        Case tpINFERMIERE: nomeTabella = "INFERMIERI"
        Case tpPSICOLOGI: nomeTabella = "PSICOLOGI"
        Case tpACCOMPAGNATORI: nomeTabella = "ACCOMPAGNATORI"
    End Select
End Function

Private Sub Cerca()
    ' cerca il paziente
    Dim chiaveRic As String
    Dim strSql As String
    Dim condizione As String
    
    ' pulisce la flx azzerando le righe
    flxGriglia.Rows = 1
    chiaveRic = UCase(txtCerca.Text)
    'tTrova.condizione = "(KEY=1 OR KEY=2 OR KEY=3)"
    
    condizione = IIf(tTrova.condizione <> "", " AND ", "") & tTrova.condizione
    If tTrova.Tipo = tpPAZIENTE And cboStato.Visible = True And cboStato.Text <> "Tutti" Then
        If Not isDialisi Then
             condizione = condizione & " AND STATO=" & cboStato.ItemData(cboStato.ListIndex)
        Else
            If cboStato.ListIndex <> 0 Then
                condizione = " AND STATO=" & cboStato.ItemData(cboStato.ListIndex)
            End If
        End If
    End If
    If tTrova.Tipo = tpINFERMIERE Then
        condizione = condizione & " AND ELIMINATO=FALSE "
    End If
    
    strSql = "SELECT * FROM " & nomeTabella & " WHERE COGNOME LIKE '" & Apostrophe(chiaveRic) & "%' " & condizione & " ORDER BY COGNOME, NOME"
    Set rsDatasetCerca = New Recordset
    rsDatasetCerca.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDatasetCerca.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDatasetCerca("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDatasetCerca("COGNOME")
            .TextMatrix(.Rows - 1, 2) = "" & rsDatasetCerca("NOME")
            If tTrova.Tipo = tpPAZIENTE Then
                .TextMatrix(.Rows - 1, 3) = rsDatasetCerca("DATA_NASCITA")
            End If
            rsDatasetCerca.MoveNext
        End With
    Loop
    lblVoci = testoVoce & rsDatasetCerca.RecordCount
    flxGriglia.Row = 0
    Set rsDatasetCerca = Nothing
End Sub

Private Sub cboStato_Click()
    Call Cerca
End Sub

Private Sub cmdAvanti_Click()
    ' il caricamento dei dati avviene in CaricaPaziente di ogni form
    If flxGriglia.Row <> 0 Then
        tTrova.keyReturn = flxGriglia.TextMatrix(flxGriglia.Row, 0)
    Else
        tTrova.keyReturn = 0
        MsgBox Me.Caption, vbInformation, "Attenzione"
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdCerca_Click()
    Call Cerca
End Sub

Private Sub cmdIndietro_Click()
    tTrova.keyReturn = 0
    Unload Me
End Sub
'--Attiva la rotellina del mouse----
'Private Sub flxGriglia_GotFocus()
'    Call WheelHook(Me, flxGriglia)
'End Sub

'Private Sub flxGriglia_LostFocus()
'    Call WheelUnHook
'End Sub
'------------------------------------

Private Sub flxGriglia_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If flxGriglia.Rows = 1 Then Exit Sub
    If flxGriglia.Row = flxGriglia.Rows - 1 Then
        i = 1
    Else
        i = flxGriglia.Row + 1
    End If
    Do
        If UCase(Mid(flxGriglia.TextMatrix(i, 1), 1, 1)) = UCase(Chr(KeyAscii)) Then
            flxGriglia.Row = i
            If i > 16 Then
                flxGriglia.TopRow = i
            End If
            Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
            Exit Do
        End If
        If i = flxGriglia.Rows - 1 Then
            i = 1
        Else
            i = i + 1
        End If
    Loop Until i = flxGriglia.Row
End Sub

Private Sub flxGriglia_Click()
    On Error GoTo gestione
    Dim numCol As Integer
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
        Exit Sub
    Else
        numCol = IIf(tTrova.Tipo = tpMEDICOBASE, 2, 3)
        Call ColoraFlx(flxGriglia, numCol)
    End If
    Exit Sub
gestione:
    If Err.Number = 13 Then
        tTrova.keyReturn = 0
    Else
        MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
    End If
End Sub

Private Sub flxGriglia_DblClick()
   If VerificaClickFlx(flxGriglia) = False Then Exit Sub
   cmdAvanti_Click
End Sub

Private Sub txtCerca_Change()
    Call Cerca
End Sub

Private Sub txtCerca_GotFocus()
    txtCerca.BackColor = colArancione
End Sub

Private Sub txtCerca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdIndietro_Click
    End If
End Sub

Private Sub txtCerca_LostFocus()
    txtCerca.BackColor = vbWhite
End Sub

Private Sub WheelCatcher1_WheelRotation(Rotation As Long, X As Long, Y As Long, CtrlHwnd As Long)
On Error GoTo gestione
' se NON è stata selezionata una riga esce e NON attiva lo scroll
'    If flxGriglia.Row = 0 Then
'       Exit Sub
'    End If

    Select Case CtrlHwnd

        Case flxGriglia.hWnd
            If flxGriglia.TopRow - Rotation > 0 Then
               flxGriglia.TopRow = flxGriglia.TopRow - Rotation
            End If
    
        End Select
' Evita crash in caso di griglia non completa
gestione:
End Sub

