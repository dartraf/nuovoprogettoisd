VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReniDaRottamare 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elenco reni da rottamare"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Label lblTesto 
         AutoSize        =   -1  'True
         Caption         =   "Reni prossimi alla rottamazione. Provvedere alla sostituzione entro la data indicata"
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
         Left            =   310
         TabIndex        =   2
         Top             =   720
         Width           =   8610
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
         Left            =   3120
         TabIndex        =   1
         Top             =   180
         Width           =   2790
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         FormatString    =   $"frmReniDaRottamare.frx":0000
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
      TabIndex        =   5
      Top             =   3240
      Width           =   9255
      Begin VB.CommandButton cmdSostituisci 
         Caption         =   "Sostituisci"
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
         Left            =   6480
         TabIndex        =   7
         Top             =   240
         Width           =   1335
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
         Height          =   495
         Left            =   7920
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReniDaRottamare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmReniDaRottamare.frm
'
' <b>Descrizione</b>: Pannello Reni Da Rottamare mostra i reni che sono da rottamare
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

Private Sub Form_Load()
    Dim i As Integer
    
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 0 To 6
            .Col = i
            .CellFontBold = True
        Next i
        .MousePointer = flexCustom
    End With
    Call CaricaFlx
End Sub

Private Sub CaricaFlx()
    Dim data As Date
    
    data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
    flxGriglia.Rows = 1
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM RENI WHERE DATA_ROTTAMAZIONE<#" & data & "# AND SOSTITUITO=FALSE ORDER BY POSTAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("POSTAZIONE")
            .TextMatrix(.Rows - 1, 2) = rsDataset("NUMERO_RENE") & ""
            .TextMatrix(.Rows - 1, 3) = rsDataset("TIPO_RENE")
            .TextMatrix(.Rows - 1, 4) = rsDataset("MATRICOLA")
            .TextMatrix(.Rows - 1, 6) = rsDataset("DATA_ROTTAMAZIONE") & ""
            If rsDataset("TIPO") = 0 Then
                .TextMatrix(.Rows - 1, 5) = "NEG"
            ElseIf rsDataset("TIPO") = 1 Then
                .TextMatrix(.Rows - 1, 5) = "HCV POS"
            Else
                .TextMatrix(.Rows - 1, 5) = "HBV POS"
            End If
        End With
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    flxGriglia.Row = 0
    Set rsDataset = Nothing
End Sub

'' Chiude Isodial se il rene non � stato sostituito
Private Sub cmdChiudi_Click()
    Dim data As Date
    Dim ret As Long
    Dim rsDataset As New Recordset
    Dim numero As Integer
    Dim trovato As Boolean
    
    data = DateValue(Month(date) & "/" & Day(date) & "/" & Year(date))
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM RENI WHERE DATA_ROTTAMAZIONE<#" & data & "# AND SOSTITUITO=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        trovato = True
    Else
        trovato = False
    End If
    rsDataset.Close
    Set rsDataset = Nothing
    
    If trovato Then
        MsgBox "IMPOSSIBILE PROSEGUIRE!!! SOSTITUZIONE OBBLIGATORIA DEI RENI IN ROTTAMAZIONE", vbInformation, "Reni in rottamazione"
        'On Error Resume Next

        '/ release object
        Call objMenuEx.Uninstall(frmMain.hWnd, frmMain.ImageList1, MenuEvents)
        Set MenuEvents = Nothing
        Set objMenuEx = Nothing
    
        If Not structApri.server Then
            ' esce dalla lista dei client collegati
            rsDataset.Open "CLIENT", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.Update "NUMERO", rsDataset("NUMERO") - 1
            Set rsDataset = Nothing
            Set cnPrinc = Nothing
            Set cnTrac = Nothing
            tRete = tpDISCONNETTI
            frmAttendi.Show 1
        Else
            rsDataset.Open "CLIENT", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            numero = rsDataset("NUMERO")
            Set rsDataset = Nothing
            If numero = 0 Then
                ' chiude la connessione
                Set cnPrinc = Nothing
                Set cnTrac = Nothing
                ' chiude la condivisione
                Call Shell("NET SHARE RISORSA /DELETE", vbHide)
                ' smonta il volume
                ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)
            Else
                Call PulisciTabCLIENTI
                ' chiude la connessione
                Set cnPrinc = Nothing
                Set cnTrac = Nothing
                ' chiude la condivisione
                Call Shell("NET SHARE RISORSA /DELETE", vbHide)
                ' smonta il volume
                ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)
            End If
        End If
        End
    Else
        Unload Me
    End If
End Sub

'' Effettua la sostituzione del rene da rottamare
Private Sub cmdSostituisci_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim num As Integer
    
    If flxGriglia.Row <> 0 Then
        tInput.Tipo = tpIRENI
        tInput.v_valori(1) = flxGriglia.TextMatrix(vRow, 1)
        tInput.mantieniDati = True
        frmInput.Show 1
        If Not (tInput.v_valori(1) = "" And tInput.v_valori(2) = "") Then
            num = GetNumero("RENI")
            v_Nomi = Array("KEY", "POSTAZIONE", "TIPO_RENE", "MATRICOLA", "TIPO", "DATA_ROTTAMAZIONE", "SOSTITUITO", "NUMERO_RENE")
            v_Val = Array(num, tInput.v_valori(1), tInput.v_valori(2), tInput.v_valori(3), tInput.v_valori(4), IIf(tInput.v_valori(5) = "", Null, tInput.v_valori(5)), False, tInput.v_valori(6))
            Set rsDataset = New Recordset
            rsDataset.Open "RENI", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdTable
            rsDataset.AddNew v_Nomi, v_Val
            rsDataset.Update
            rsDataset.Close
            
            rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_RENE=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            Do While Not rsDataset.EOF
                rsDataset("CODICE_RENE") = num
                rsDataset.MoveNext
            Loop
            rsDataset.Close
            
            rsDataset.Open "SELECT * FROM RENI WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            Do While Not rsDataset.EOF
                rsDataset("SOSTITUITO") = True
                rsDataset.MoveNext
            Loop
            rsDataset.Close
            
            If flxGriglia.Rows = 2 Then
                Unload Me
            Else
                flxGriglia.RemoveItem vRow
            End If
            flxGriglia.Row = 0
        End If
    Else
        MsgBox "Selezionare il rene da sostituire", vbCritical, "Attenzione"
    End If
End Sub

Private Sub flxGriglia_Click()
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        vRow = flxGriglia.Row
        vCol = flxGriglia.Col
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

'Private Sub wheelMouse_MouseScroll(MouseKeys As Long, Rotation As Long, X As Long, Y As Long, ControlHWnd As Long)
'    If ControlHWnd = flxGriglia.hWnd Then
'        If flxGriglia.TopRow - Rotation > 0 Then
'            If flxGriglia.TopRow - Rotation < flxGriglia.Rows Then
'                flxGriglia.TopRow = flxGriglia.TopRow - Rotation
'            End If
'        End If
'    End If
'End Sub
'-----------------------------------------


Private Sub Timer1_Timer()
    If lblAttenzione.ForeColor = vbRed Then
        lblAttenzione.ForeColor = vbBlack
    Else
        lblAttenzione.ForeColor = vbRed
    End If
End Sub

