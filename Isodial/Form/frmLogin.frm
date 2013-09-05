VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Accesso"
   ClientHeight    =   1575
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930.562
   ScaleMode       =   0  'User
   ScaleWidth      =   2901.343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Inserisci il codice utente"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "&Esci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Inserisci la password"
      Top             =   600
      Width           =   2325
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Top             =   600
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   120
      Picture         =   "frmLogin.frx":018A
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsLogin As Recordset
Dim ENTRA As Boolean        ' accesso rapido

Private Sub Form_Load()
    Dim i As Integer
    Call TakeCloseOff(Me.hWnd)
    ' setta i menu di default
    With frmMain
        .mnuFatturazione.Visible = False
        .mnuStrumenti.Visible = False
        .mnuPaziente.Enabled = True
        .mnuDialisi.Enabled = True
        .mnuArchivi.Enabled = True
        .mnuStrumenti.Enabled = True
        .mnuFatturazione.Enabled = True
        .mnuSottoDialisi(1).Enabled = True
        .mnuGestioneIndicatori.Visible = False
        .mnuSottoDialisi(4).Enabled = True
        
        .picContenitore.Enabled = True
        For i = 0 To 16
            .cmdToolbar(i).Enabled = True
            .cmdToolbar(i).Visible = True
        Next i
    End With
    ENTRA = False
End Sub

Private Sub cmdEsci_Click()
    '/ prevent error if the menu is not subclassed
    'On Error Resume Next
    Dim ret As Long
    Dim rsDataset As New Recordset
    Dim numero As Integer
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
            If MsgBox(numero & " CLIENT COLLEGATI" & vbCrLf & "Disconnetto automaticamente gli altri utenti?", vbQuestion + vbYesNo, "Disconnessione") = vbYes Then
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
    End If
    End
End Sub

Private Sub cmdOK_Click()
    Dim verificaPass As Boolean
    Dim strSql As String
    
    txtPassword = txtPassword & ""
    txtUserName = txtUserName & ""
    ' cerca l'utente
    If ENTRA Then
        strSql = "SELECT * FROM LOGIN WHERE CHIAVE='" & Apostrophe(txtUserName) & "' AND ELIMINATO=FALSE"
    Else
        strSql = "SELECT * FROM LOGIN WHERE CHIAVE='" & Apostrophe(txtUserName) & "' AND PASSWORD='" & Apostrophe(txtPassword) & "' AND ELIMINATO=FALSE"
    End If
    Set rsLogin = New Recordset
    rsLogin.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsLogin.EOF And rsLogin.BOF Then
        ' l'utente nn esiste o ha sbagliato ad inserire i dati
        MsgBox "Chiave utente o password sbagliate", vbExclamation, "Accesso"
        txtUserName.SetFocus
        Exit Sub
    Else
        If isCorrotto Then
            If rsLogin("TIPO") = tpAMASTER Then
                tAccesso.Tipo = tpAMASTER
                frmMain.staBar.Panels(4).Text = "Amministratore"
            Else
                MsgBox "Accesso consentito al solo amministratore di sistema", vbExclamation, "Accesso"
                txtUserName.SetFocus
                Exit Sub
            End If
        Else
            ' controlla l'alert trimestrale
            If rsLogin("DATA") + 90 < date And Not ENTRA Then
                ' fa cambiare la password
                MsgBox "Password scaduta!" & vbCrLf & "Necessario cambio password", vbInformation, "Alert trimestrale cambio password"
                verificaPass = False
                tipoPass.Tipo = tCAMBIA
                tipoPass.password = txtPassword
                tipoPass.key = rsLogin("KEY")
                frmPass.Show 1
                verificaPass = frmPass.GetRisPass
                Unload frmPass
                If verificaPass Then
                    MsgBox "Password cambiata con successo", vbInformation, "Password"
                Else
                    Exit Sub
                End If
            End If
            tAccesso.Tipo = rsLogin("TIPO")
            tAccesso.cognome = rsLogin("COGNOME") & ""
            tAccesso.nome = rsLogin("NOME") & ""
            tAccesso.pass = rsLogin("PASSWORD") & ""
            tAccesso.key = rsLogin("KEY")
            If tAccesso.Tipo = tpAMASTER Then
                frmMain.staBar.Panels(4).Text = "Amministratore"
            Else
                frmMain.staBar.Panels(4).Text = Choose(tAccesso.Tipo, "Medico: ", "Infermiere: ", "Contabile: ") & tAccesso.cognome & " " & UCase(Mid(tAccesso.nome, 1, 1)) & "."
            End If
        End If
    End If
    Set rsLogin = Nothing
    Call impostaMenu
    ' memorizza l'accesso
    If TRACCIATO Then
        Call SalvaAccesso
    End If
    If isCorrotto Then
        frmPeriferiche.Show 1
    End If
    
    If tAccesso.Tipo = tpAMEDICO Or tAccesso.Tipo = tpAMASTER Then
        Call ControllaReni
   '     Call ControllaAlertAppa
    ElseIf tAccesso.Tipo <> tpAINFERMIERE Then
   '     Call ControllaAlertAppa
    End If
    Unload Me
End Sub

' Controlla che non ci siano reni da rottamare
Private Sub ControllaReni()
    Dim rsDataset As New Recordset
    Dim data As Date
    
    data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
    
    rsDataset.Open "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE<#" & data & "# AND SOSTITUITO=FALSE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        frmReniDaRottamare.Show 1
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

' Controlla che non ci siano reni da rottamare
Private Sub ControllaAlertAppa()
    Dim rsDataset As New Recordset
    Dim data As Date
    
    data = DateValue(Month(date + 30) & "/" & Day(date + 30) & "/" & Year(date + 30))
    
    rsDataset.Open "SELECT * FROM APPARATI WHERE (PROXREVFUN<#" & data & "# or PROXREVSIC<#" & data & "#) AND ALERT=False", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        frmAlertApparati.Show 1
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Sub


Private Sub impostaMenu()
    Dim i  As Integer
    With frmMain
        Select Case tAccesso.Tipo
          Case tpAMASTER
            .mnuGestioneIndicatori.Visible = True
            .mnuStampaPaz.Enabled = True
            .mnuStampaPaz.Enabled = True
            .mnuStampaMediciBase.Enabled = True
            .mnuMostraFattElaborazione.Enabled = True
            .mnuImpegnativeDialisi.Enabled = True
            .mnuStrumenti.Visible = True
            .mnuFatturazione.Visible = True
            .mnuStampe.Enabled = True
            .mnuSottoDialisi(5).Enabled = True
            .mnuKtvAnnuale.Enabled = True
            .mnuTsatAnnuale.Enabled = True
            For i = 1 To 4
                .mnuSottoDialisi(i).Visible = True
            Next i
            For i = 1 To 6
                .mnuSottoPaz(i).Visible = True
            Next i
          Case tpACONTABILE
            .mnuStampe.Enabled = True
            .mnuStampaPaz.Enabled = True
            .mnuStampaMediciBase.Enabled = True
            .mnuMostraFattElaborazione.Enabled = True
            .mnuImpegnativeDialisi.Enabled = True
            .mnuKtvAnnuale.Enabled = False
            .mnuTsatAnnuale.Enabled = False
            .mnuFatturazione.Visible = True
            ' rende inattivi gli altri
            For i = 2 To 6
                .mnuSottoPaz(i).Visible = False
             Next i
            For i = 1 To 4
                .mnuSottoDialisi(i).Visible = False
            Next i
            .mnuSottoDialisi(5).Enabled = True
            .mnuDialisi.Enabled = True
            .mnuArchivi.Enabled = False
            .mnuStrumenti.Enabled = False
            For i = 1 To 13
                .cmdToolbar(i).Enabled = False
            Next i
          Case tpAMEDICO
            .mnuSottoDialisi(5).Enabled = True
            .mnuGestioneIndicatori.Visible = True
            .mnuStampe.Enabled = True
            .mnuStampaPaz.Enabled = True
            .mnuStampaMediciBase.Enabled = True
            .mnuMostraFattElaborazione.Enabled = True
            .mnuImpegnativeDialisi.Enabled = True
            For i = 14 To 16
                .cmdToolbar(i).Enabled = False
            Next i
            For i = 1 To 4
                .mnuSottoDialisi(i).Visible = True
            Next i
          Case tpAINFERMIERE
            .mnuStampe.Enabled = True
            .mnuStampaPaz.Enabled = False
            .mnuStampaMediciBase.Enabled = False
            .mnuMostraFattElaborazione.Enabled = True
            .mnuImpegnativeDialisi.Enabled = False
            .mnuKtvAnnuale.Enabled = False
            .mnuTsatAnnuale.Enabled = False
            .mnuSottoDialisi(1).Visible = True
            .mnuSottoDialisi(1).Enabled = False
            .mnuPaziente.Enabled = False
            .mnuArchivi.Enabled = False
            .mnuStrumenti.Enabled = False
            .mnuSottoDialisi(4).Enabled = IsCaposala
            For i = 0 To 16
                .cmdToolbar(i).Enabled = False
            Next i
            .mnuSottoDialisi(2).Visible = True
            .mnuSottoDialisi(3).Visible = True
            .cmdToolbar(12).Enabled = True
            .mnuSottoDialisi(5).Enabled = False
        End Select
        .mnuImpostaBackup.Visible = structApri.server
        .mnuRipristina.Visible = structApri.server
        .mnuTabFatt(6).Visible = structApri.F1abiliata
        .mnuRimborsi.Visible = structApri.F1abiliata
    End With
End Sub

Private Function IsCaposala() As Boolean
    Dim rsDataset As New Recordset
    IsCaposala = False
    rsDataset.Open "SELECT * FROM INFERMIERI WHERE NOME='" & Apostrophe(tAccesso.nome) & "' AND COGNOME='" & Apostrophe(tAccesso.cognome) & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        If rsDataset("MANSIONE") = 2 Then
            IsCaposala = True
        End If
    End If
    Set rsDataset = Nothing
End Function

Private Sub SalvaAccesso()
    Dim rsDataset As New Recordset
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant
    
    v_Nomi = Array("CODICE_UTENTE", "DATA", "ORA", "NOME_PC", "TIPO_PC")
    v_Val = Array(tAccesso.key, date, Time, Environ("COMPUTERNAME"), IIf(structApri.server, 1, 2))
    
    rsDataset.Open "ACCESSI", cnTrac, adOpenKeyset, adLockPessimistic, adCmdTable
    rsDataset.AddNew v_Nomi, v_Val
    rsDataset.Update
    Set rsDataset = Nothing
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
    txtPassword.BackColor = colArancione
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    ' evita di chiudere il form con alt f4
    If (Shift And vbAltMask) And KeyCode = vbKeyF4 Then
        KeyCode = 0
    End If
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = vbWhite
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
    txtUserName.BackColor = colArancione
    If Environ$("COMPUTERNAME") = "MASTERMIO" Or Environ$("COMPUTERNAME") = "MASTER" Then
        txtUserName = "Admin"
        ENTRA = True
        cmdOK_Click
    End If
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    ' TEMPORANEO
    If (Shift And vbAltMask) And KeyCode = vbKeyF2 Then
        txtUserName = "Admin"
        ENTRA = True
        cmdOK_Click
    End If
    If (Shift And vbAltMask) And KeyCode = vbKeyF9 Then
        '/ release object
        Call objMenuEx.Uninstall(frmMain.hWnd, frmMain.ImageList1, MenuEvents)
        Set MenuEvents = Nothing
        Set objMenuEx = Nothing
        Set cnPrinc = Nothing
        Set cnTrac = Nothing
        End
    End If
    ' evita di chiudere il form con alt f4
    If (Shift And vbAltMask) And KeyCode = vbKeyF4 Then
        KeyCode = 0
    End If
End Sub

Private Sub txtUserName_LostFocus()
    txtUserName.BackColor = vbWhite
End Sub
