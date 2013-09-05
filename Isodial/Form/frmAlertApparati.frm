VERSION 5.00
Begin VB.Form frmAlertApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prossime Revisioni Apparati"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9225
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
      Width           =   9015
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
      Width           =   9015
      Begin VB.PictureBox flxGriglia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   8715
         TabIndex        =   4
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   9015
      Begin VB.CommandButton cmdSostParco 
         Caption         =   "Sostituisci da Parco Reni"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   220
         Width           =   1660
      End
      Begin VB.CommandButton cmdSostituisci 
         Caption         =   "Sostituisci con Rene Nuovo"
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
         Left            =   5640
         TabIndex        =   7
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
         Left            =   7440
         TabIndex        =   6
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
    rsDataset.Open "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE<#" & data & "# AND SOSTITUITO=FALSE ORDER BY POSTAZIONE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsDataset("KEY")
            .TextMatrix(.Rows - 1, 1) = rsDataset("POSTAZIONE")
            .TextMatrix(.Rows - 1, 2) = rsDataset("NUMERO_APPARATO") & ""
            .TextMatrix(.Rows - 1, 3) = rsDataset("MODELLO")
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

'' Chiude Isodial se il rene non è stato sostituito
Private Sub cmdChiudi_Click()
    Dim data As Date
    Dim ret As Long
    Dim rsDataset As New Recordset
    Dim numero As Integer
    Dim trovato As Boolean
    
    data = DateValue(Month(date) & "/" & Day(date) & "/" & Year(date))
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE<#" & data & "# AND SOSTITUITO=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
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
Private Sub cmdSostParco_Click()
    If flxGriglia.Row <> 0 Then
        sostituito = False
        frmTabellaReni.Show 1
        If sostituito And flxGriglia.Rows > 2 Then
           flxGriglia.RemoveItem vRow
           flxGriglia.Row = 0
        ElseIf sostituito And flxGriglia.Rows = 2 Then
           flxGriglia.Row = 0
           Unload Me
        End If
     Else
        MsgBox "Selezionare il rene da sostituire", vbCritical, "Attenzione"
    End If
End Sub

'' Effettua la sostituzione del rene da rottamare
Private Sub cmdSostituisci_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    Dim num As Integer
    
    If flxGriglia.Row <> 0 Then
 '       tInput.Tipo = tpIRENI
        tInput.v_valori(1) = flxGriglia.TextMatrix(vRow, 1)
        tInput.mantieniDati = True
        frmApparatiInput.Show 1
        If Not (tInput.v_valori(1) = "" And tInput.v_valori(2) = "") Then
            num = numKey
 '         cboTipoApparato(0) = cboTipoApparatoPrec
 '       txtpostazione = PostazionePrec
 '       cboTipoRene.Text = cboTipoRenePrec
          
            Set rsDataset = New Recordset
           
            rsDataset.Open "SELECT * FROM TURNI WHERE CODICE_RENE=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
            Do While Not rsDataset.EOF
                rsDataset("CODICE_RENE") = num
                rsDataset.MoveNext
            Loop
            rsDataset.Close
            
            rsDataset.Open "SELECT * FROM APPARATI WHERE KEY=" & flxGriglia.TextMatrix(vRow, 0), cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
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
        dt_rott_rene = flxGriglia.TextMatrix(vRow, 6)
        cod_rene = flxGriglia.TextMatrix(vRow, 0)
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

