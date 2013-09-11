VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{892E8F6D-4FB0-4046-9D7A-C6882F0F0CEB}#2.0#0"; "WheelCatcher.ocx"
Begin VB.Form frmAlertApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prossime Revisioni Apparati"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VB.Label lblTesto 
         AutoSize        =   -1  'True
         Caption         =   "APPARATI PROSSIMI ALLA MANUTENZIONE"
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
         Left            =   5640
         TabIndex        =   2
         Top             =   315
         Width           =   4740
      End
      Begin VB.Label lblAttenzione 
         AutoSize        =   -1  'True
         Caption         =   "ATTENZIONE!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   210
         Width           =   2790
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   12855
      Begin WheelCatch.WheelCatcher WheelCatcher1 
         Height          =   480
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   3255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   99
         FormatString    =   "| Tabella                                                                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmAlertApparati.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   12855
      Begin VB.CommandButton cmdDisTutAlert 
         Caption         =   "Disattiva tutti gli Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   1660
      End
      Begin VB.CommandButton cmdDisSingAlert 
         Caption         =   "Disattiva Singolo Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7680
         TabIndex        =   7
         Top             =   240
         Width           =   1660
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "Stampa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9480
         TabIndex        =   6
         Top             =   220
         Width           =   1660
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11280
         TabIndex        =   5
         Top             =   220
         Width           =   1400
      End
   End
End
Attribute VB_Name = "frmAlertApparati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmAlertApparati.frm
'
' <b>Descrizione</b>: Pannello Prossime Revisioni Apparati mostra gli apparati che sono da revisionare
'
' @remarks
'
' @author
'
' @date 03/06/2011 18.22

Option Explicit

Dim rsDataset As Recordset
Dim vRow As Integer
Dim vCol As Integer

Private Sub cmdDisSingAlert_Click()
    Dim data As Date
    If KeyApparato = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Si conferma la disabilitazione dell'alert per l'apparato selezionato ?", vbQuestion + vbYesNo + vbDefaultButton2, "Disattivazione ALERT") = vbYes Then
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM APPARATI WHERE KEY =" & KeyApparato, cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
        rsDataset("LETTO") = True
        rsDataset.Update
        rsDataset.Close
        Set rsDataset = Nothing
        MsgBox "ALERT DISATTIVATO per l'apparato selezionato", vbInformation, "Disattivazione ALERT"
    Else
        KeyApparato = 0
        Exit Sub
    End If

End Sub

Private Sub cmdDisTutAlert_Click()
    Dim data As Date

    If MsgBox("Si conferma la disabilitazione dell'alert per gli apparati visualizzati ?", vbQuestion + vbYesNo + vbDefaultButton2, "Disattivazione ALERT") = vbYes Then
        data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
        Set rsDataset = New Recordset
        'se si cambia la select cambiarla nella sub->Caricaflx e nel form LOGIN->Sub->ControllaAlertAppa
        rsDataset.Open "SELECT * FROM APPARATI WHERE (PROXREVFUN<#" & data & "# or PROXREVSIC<#" & data & "#) AND SOSTITUITO=FALSE AND LETTO=FALSE ORDER BY TIPO_APPARATO,PROXREVFUN,PROXREVSIC", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText
        Do While Not rsDataset.EOF
            rsDataset("LETTO") = True
            rsDataset.Update
            rsDataset.MoveNext
        Loop
        rsDataset.Close
        Set rsDataset = Nothing
        MsgBox "ALERT DISATTIVATI!!!", vbInformation, "Disattivazione ALERT"
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim data As Date
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim TotaleReni As Integer
    
    data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(4) AS NUMERO_APPARATO, " & _
                "       NEW adVarChar(4) AS POSTAZIONE, " & _
                "       NEW adVarChar(50) AS TIPO_APPARATO, " & _
                "       NEW adVarChar(50) AS MODELLO, " & _
                "       NEW adVarChar(50) AS MATRICOLA, " & _
                "       NEW adVarChar(50) AS PRODUTTORE, " & _
                "       NEW adDate AS PROXREVFUN, " & _
                "       NEW adDate AS PROXREVSIC"
                
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    'se si cambia la select cambiarla nella sub->Caricaflx, nella sub->cmdDisTutAlert_Click e nel form LOGIN->Sub->ControllaAlertAppa
    rsDataset.Open "SELECT * FROM APPARATI WHERE (PROXREVFUN<#" & data & "# or PROXREVSIC<#" & data & "#) AND SOSTITUITO=FALSE AND LETTO=FALSE ORDER BY TIPO_APPARATO,PROXREVFUN,PROXREVSIC", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            Do While Not rsDataset.EOF
                .AddNew
                .Fields("NUMERO_APPARATO") = rsDataset("NUMERO_APPARATO")
                .Fields("POSTAZIONE") = rsDataset("POSTAZIONE")
                .Fields("TIPO_APPARATO") = rsDataset("TIPO_APPARATO")
                .Fields("MODELLO") = rsDataset("MODELLO")
                .Fields("MATRICOLA") = rsDataset("MATRICOLA")
                .Fields("PRODUTTORE") = rsDataset("PRODUTTORE")
                .Fields("PROXREVFUN") = rsDataset("PROXREVFUN")
                .Fields("PROXREVSIC") = rsDataset("PROXREVSIC")
                rsDataset.MoveNext
            Loop
        End With
    End If
    
    If rsDataset.RecordCount > 0 Then
        TotaleReni = rsDataset.RecordCount
    End If
    
    Set rsDataset = Nothing
    
    Set rptApparatiAlert.DataSource = rsMain
    rptApparatiAlert.Orientation = rptOrientLandscape
    rptApparatiAlert.TopMargin = 0
    rptApparatiAlert.RightMargin = 0
    rptApparatiAlert.LeftMargin = 0
    rptApparatiAlert.Sections("Intestazione").Controls("lblElenco").Caption = "Apparati Prossimi alla Revisione"
    rptApparatiAlert.Sections("Section5").Controls.Item("lblTotaleReni").Caption = TotaleReni
    rptApparatiAlert.PrintReport True, rptRangeAllPages
End Sub

Private Sub Form_Load()
    Dim i As Integer

    flxGriglia.Rows = 1
    
    With flxGriglia
        .Cols = 9
        .ColWidth(1) = .ColWidth(1) * 0.2
        .ColWidth(2) = .ColWidth(2) * 0.8
        .ColWidth(3) = .ColWidth(3) * 2.8
        .ColWidth(4) = .ColWidth(4) * 2
        .ColWidth(5) = .ColWidth(5) * 1
        .ColWidth(6) = .ColWidth(6) * 3.02
        .ColWidth(7) = .ColWidth(7) * 1.3
        .ColWidth(8) = .ColWidth(8) * 1.3
                                     
        .TextMatrix(0, 1) = "N°App."
        .TextMatrix(0, 2) = "Postaz."
        .TextMatrix(0, 3) = "Categoria Apparato"
        .TextMatrix(0, 4) = "Modello"
        .TextMatrix(0, 5) = "Matricola"
        .TextMatrix(0, 6) = "Produttore"
        .TextMatrix(0, 7) = "Pros.Rev.Fun."
        .TextMatrix(0, 8) = "Pros.Rev.Sic."
    
    End With
    
    Call CaricaFlx
    
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        .MousePointer = flexCustom
        For i = 0 To flxGriglia.Cols - 1
            .Col = i
            .ColAlignment(i) = vbLeftJustify
            .CellFontBold = True
        Next i
    End With
End Sub

Private Sub CaricaFlx()
    Dim data As Date
    
    data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
    flxGriglia.Rows = 1
    vCol = 0
    vRow = 0
      
    Set rsDataset = New Recordset
    'se si cambia la select cambiarla nella sub->cmdDisTutAlert_Click e nel form LOGIN->Sub->ControllaAlertAppa
    rsDataset.Open "SELECT * FROM APPARATI WHERE (PROXREVFUN<#" & data & "# or PROXREVSIC<#" & data & "#) AND SOSTITUITO=FALSE AND LETTO=FALSE ORDER BY TIPO_APPARATO,PROXREVFUN,PROXREVSIC", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsDataset.EOF
        With flxGriglia
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
        .TextMatrix(.Rows - 1, 1) = rsDataset("NUMERO_APPARATO")
        .TextMatrix(.Rows - 1, 2) = rsDataset("POSTAZIONE")
        .TextMatrix(.Rows - 1, 3) = rsDataset("TIPO_APPARATO")
        .TextMatrix(.Rows - 1, 4) = rsDataset("MODELLO")
        .TextMatrix(.Rows - 1, 5) = rsDataset("MATRICOLA")
        .TextMatrix(.Rows - 1, 6) = rsDataset("PRODUTTORE")
        .TextMatrix(.Rows - 1, 7) = rsDataset("PROXREVFUN")
        .TextMatrix(.Rows - 1, 8) = rsDataset("PROXREVSIC")
        End With
    rsDataset.MoveNext
    Loop
    rsDataset.Close
    flxGriglia.Row = 0
    Set rsDataset = Nothing

End Sub

Private Sub cmdChiudi_Click()
  Unload Me
End Sub

Private Sub flxGriglia_Click()
    Dim keyappa As Integer
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        vRow = flxGriglia.Row
        vCol = flxGriglia.Col
        KeyApparato = flxGriglia.TextMatrix(vRow, 0)
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
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

Private Sub Timer1_Timer()
    If lblAttenzione.ForeColor = vbRed Then
        lblAttenzione.ForeColor = vbBlack
    Else
        lblAttenzione.ForeColor = vbRed
    End If
End Sub

