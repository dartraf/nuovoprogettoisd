VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchedeSorveglianzaIAV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schede Sorveglianza IAV"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmSchedeSorveglianzaIAV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   450
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
         TabIndex        =   7
         Top             =   360
         Width           =   465
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
         TabIndex        =   6
         Top             =   360
         Width           =   630
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
         TabIndex        =   5
         Top             =   360
         Width           =   1005
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
         TabIndex        =   4
         Top             =   360
         Width           =   615
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
         Left            =   7200
         TabIndex        =   3
         Top             =   360
         Width           =   3135
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
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6495
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   12855
      Begin VB.TextBox txtAppo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
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
         Left            =   1200
         MaxLength       =   35
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   3120
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrigliaSintomi 
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         MousePointer    =   15
         FormatString    =   "| Segni e Sintomi locali                         | Valore                       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmSchedeSorveglianzaIAV.frx":0459
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   7080
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
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSchedeSorveglianzaIAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDataset As Recordset
Dim rsSintomi As Recordset
Dim intPazientiKey As Integer
Dim vRow As Integer
Dim vCol As Integer

Private Sub cmdChiudi_Click()
    Unload frmSchedeSorveglianzaIAV
End Sub

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If intPazientiKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTop As Single
    Dim intLeft As Single
   
    Call GetCenterForm(Me.Height, Me.Width, intTop, intLeft)
    Me.Top = intTop
    Me.Left = intLeft
    
    With flxGrigliaSintomi
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 2
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
        .MousePointer = flexCustom
    End With
    
    flxGrigliaSintomi.ColAlignment(1) = vbLeftJustify
    flxGrigliaSintomi.Rows = 1
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    flxGrigliaSintomi.Rows = 1
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        If intPazientiKey = tTrova.keyReturn Then
            intPazientiKey = 0
            Call CaricaPaziente
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        Else
            intPazientiKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub CaricaPaziente()
    
    If intPazientiKey = 0 Then
        ' pulisce la griglia
        ' pulisce la flx azzerando le righe
        flxGrigliaSintomi.Rows = 1
    Else
        ' carica i dati del paziente
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & intPazientiKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
       
        ' cerca i riferimenti al paziente
        Call CaricaFlxSintomi
    
    End If
End Sub

Private Sub CaricaFlxSintomi()
    
    ' Carico i Sintomi all' interno della colonna 1
    With flxGrigliaSintomi
        .Rows = .Rows + 1
            .TextMatrix(1, 1) = "Eritema"
        .Rows = .Rows + 1
            .TextMatrix(2, 1) = "Dolore"
        .Rows = .Rows + 1
            .TextMatrix(3, 1) = "Gonfiore"
        .Rows = .Rows + 1
            .TextMatrix(4, 1) = "Infiltrazione"
        .Rows = .Rows + 1
            .TextMatrix(5, 1) = "Presenza Fremito"
    End With
                
    Set rsSintomi = New Recordset
    
    rsSintomi.Open "SELECT * FROM SCHEDA_SORV_IAV WHERE KEY_PAZIENTE= " & intPazientiKey & " ORDER BY KEY DESC ", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not (rsSintomi.BOF And rsSintomi.EOF) Then
        Do While Not rsSintomi.EOF
            With flxGrigliaSintomi
                .TextMatrix(1, 0) = rsSintomi("KEY")
                .TextMatrix(1, 2) = rsSintomi("ERITEMA") & ""
                .TextMatrix(2, 2) = rsSintomi("DOLORE") & ""
                .TextMatrix(3, 2) = rsSintomi("GONFIORE") & ""
                .TextMatrix(4, 2) = rsSintomi("INFILTRAZIONE") & ""
                .TextMatrix(5, 2) = rsSintomi("PRESENZA_FREMITO") & ""
                rsSintomi.MoveNext
            End With
        Loop
        Set rsSintomi = Nothing
        flxGrigliaSintomi.Row = 0
    End If
        
End Sub

Private Sub flxGrigliaSintomi_Click()
    flxGrigliaSintomi.SetFocus
    If VerificaClickFlx(flxGrigliaSintomi) = False Then
        ' discolora
        Call ColoraFlx(flxGrigliaSintomi, flxGrigliaSintomi.Cols - 1, True)
        ' annulla le row e col
        flxGrigliaSintomi.Row = 0
        flxGrigliaSintomi.Col = 0
    Else
        vCol = flxGrigliaSintomi.Col
        vRow = flxGrigliaSintomi.Row
        Call ColoraFlx(flxGrigliaSintomi, flxGrigliaSintomi.Cols - 1)
    End If
End Sub

Private Sub flxGrigliaSintomi_DblClick()
    If VerificaClickFlx(flxGrigliaSintomi) = False Then Exit Sub
     
    With flxGrigliaSintomi
        .SetFocus
        
        ' Se la colonna è quella dei valori
        If .Col = 2 Then
        
        ' Scrive i valori in rosso
        flxGrigliaSintomi.CellForeColor = vbRed
       '     Call objAnnulla.Add(.TextMatrix(.Row, .Col), .Col, .TextMatrix(.Row, 0))
       '     cmdAnnulla.Enabled = True
       ' Con la pressione del muose mi cambia tutti i valori
                If .TextMatrix(.Row, 2) = "" Then
                    .TextMatrix(.Row, 2) = "SI"
                ElseIf .TextMatrix(.Row, 2) = "SI" Then
                    .TextMatrix(.Row, 2) = "NO"
                ElseIf .TextMatrix(.Row, 2) = "NO" Then
                    .TextMatrix(.Row, 2) = "LIEVE"
                ElseIf .TextMatrix(.Row, 2) = "LIEVE" Then
                    .TextMatrix(.Row, 2) = "MENO"
                ElseIf .TextMatrix(.Row, 2) = "MENO" Then
                    .TextMatrix(.Row, 2) = "GRAVE"
                ElseIf .TextMatrix(.Row, 2) = "GRAVE" Then
                    .TextMatrix(.Row, 2) = ""
                End If
                Call SalvaModificheSintomi
        End If
    End With
End Sub

Private Sub SalvaModificheSintomi()
    Dim keyId As Integer
    Dim v_Nomi(1 To 7) As Variant
    Dim v_Val(1 To 7) As Variant
    Dim rsSalvaModificheSintomi As Recordset
    
    v_Nomi(1) = "KEY"
    v_Nomi(2) = "KEY_PAZIENTE"
    v_Nomi(3) = "ERITEMA"
    v_Nomi(4) = "DOLORE"
    v_Nomi(5) = "GONFIORE"
    v_Nomi(6) = "INFILTRAZIONE"
    v_Nomi(7) = "PRESENZA_FREMITO"
    
    With flxGrigliaSintomi
        keyId = .TextMatrix(1, 0)
        v_Val(1) = .TextMatrix(1, 0)
        v_Val(2) = intPazientiKey
        v_Val(3) = .TextMatrix(1, 2)
        v_Val(4) = .TextMatrix(2, 2)
        v_Val(5) = .TextMatrix(3, 2)
        v_Val(6) = .TextMatrix(4, 2)
        v_Val(7) = .TextMatrix(5, 2)
        
        Set rsSalvaModificheSintomi = New Recordset
        rsSalvaModificheSintomi.Open "SELECT * FROM SCHEDA_SORV_IAV WHERE KEY=" & keyId, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        rsSalvaModificheSintomi.Update v_Nomi, v_Val
        Set rsSalvaModificheSintomi = Nothing
        
    End With

End Sub


