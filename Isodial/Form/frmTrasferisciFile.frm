VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTrasferisciFile 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation anmAvi 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      _Version        =   393216
      Center          =   -1  'True
      BackColor       =   -2147483637
      FullWidth       =   297
      FullHeight      =   49
   End
   Begin VB.Label lblScritta 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "L'operazione potrebbe richiedere alcuni minuti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4440
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAttendi 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Attendere prego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2850
   End
End
Attribute VB_Name = "frmTrasferisciFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Megabyte = 1048576
Dim finito As Boolean
Dim numClient As Integer

Private Type structFile
    data As Date
    num As Integer
End Type
Dim records() As structFile
Dim MAX_BACKUP As Integer

Private Sub Form_Load()
    Call TakeCloseOff(Me.hWnd)
    finito = False
End Sub

Private Sub Form_Activate()
    Dim lettera As String
    If VerificaDiscoRimovibile(lettera) Then
        If tPeriferica = tpBACKUP Then
            If nessunClient(numClient) Then
                If nonCorrotto Then
                    lblScritta = "Database integro. Backup in corso" & vbCrLf & "L'operazione potrebbe richiedere alcuni minuti"
                    Call Copia(lettera)
                Else
                    MsgBox "Impossibile procedere al backup" & vbCrLf & "Ripristinare un precedente backup o richiedere l'intervento tecnico", vbCritical, "Database corrotto"
                    tDisconnetti = tpDANNULLA
                    Unload Me
                End If
            Else
                'MsgBox "Impossibile effettuare il backup" & vbCrLf & "Disconnettere prima gli altri utenti client" & vbCrLf & "Connessi " & numClient & " client", vbCritical, "Attenzione"
                'tDisconnetti = tpDANNULLA
                'Unload Me
                If MsgBox(numClient & " CLIENT COLLEGATI" & vbCrLf & "Disconnetto automaticamente gli altri utenti?", vbQuestion + vbYesNo, "Disconnessione") = vbYes Then
                    If nonCorrotto Then
                        lblScritta = "Database integro. Backup in corso" & vbCrLf & "L'operazione potrebbe richiedere alcuni minuti"
                        Call PulisciTabCLIENTI
                        Call Copia(lettera)
                    Else
                        MsgBox "Impossibile procedere al backup" & vbCrLf & "Ripristinare un precedente backup o richiedere l'intervento tecnico", vbCritical, "Database corrotto"
                        tDisconnetti = tpDANNULLA
                        Unload Me
                    End If
                Else
                    tDisconnetti = tpDANNULLA
                    Unload Me
                End If
            End If
        Else
            Call RipristinaArchivio(lettera)
        End If
    Else
        MsgBox "Impossibile effettuare il backup del database", vbCritical, "Disco rimovibile non presente"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If finito Then
        Unload Me
    End If
End Sub

Private Sub CaricaMaxBackup()
    ' carica il numeridi backup impostato dall'utente
    Dim rsDataset As New Recordset
    rsDataset.Open "IMPOSTAZIONI_BACKUP", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    MAX_BACKUP = rsDataset("NUMERO")
    Set rsDataset = Nothing
End Sub

Private Sub Copia(lettera As String)
    Dim tempo As Single
    Dim ret As Double
    On Error GoTo gestioneerror
    
    Screen.MousePointer = cc2Hourglass
    anmAvi.Open App.Path & "\clip.avi"
    anmAvi.Play
    tempo = Timer
    Do
        DoEvents
    Loop Until tempo + 1 <= Timer
    
    Call CaricaMaxBackup
    ' chiude la connessione
    Set cnPrinc = Nothing
    Set cnTrac = Nothing
    ' chiude la condivisione
    Call Shell("NET SHARE RISORSA /DELETE", vbHide)
    ' smonta il volume
    ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)
    
    Call BackupArchivio(lettera)
    
    anmAvi.Stop
    anmAvi.Visible = False
    Screen.MousePointer = 0
    Me.Height = 1485
    lblAttendi = "Premere un tasto per chiudere"
    lblScritta = "Backup eseguito correttamente"
    Call BloccoCentri
    Me.SetFocus
    If SpegniPc Then
        tempo = Timer
        Do
            DoEvents
        Loop Until tempo + 5 <= Timer
        Call Spegni
    End If
    finito = True
    Exit Sub
    
gestioneerror:
    Call GestioneErrore
End Sub

Private Sub BloccoCentri()
'    If structIntestazione.sCodiceSTS = "480205" And date >= DateValue("03/05/2011") Then
'        Kill structApri.pathExe & "\tabelle.xml"
'    End If
'    If structIntestazione.sCodiceSTS = "AD0163" And date >= DateValue("28/05/2011") Then
'        Kill structApri.pathExe & "\tabelle.xml"
'    End If
End Sub

Private Function GestisciFile(lettera As String) As Integer
    ' gestisce i file Dati.dat e i vari backup
    On Error GoTo gestione
    Dim numRecord As Integer
    Dim stessoGiorno As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim dataAppo As Date
    Dim numAppo As Integer
    
    stessoGiorno = False
    ReDim records(0)

    If Dir(lettera & ":\Dati.dat") <> "" Then
        ' legge il file
        Open lettera & ":\Dati.dat" For Random As 1
        i = 0
        Do While Not EOF(1)
            Get 1, i + 1, records(i)
            ReDim Preserve records(UBound(records) + 1)
            i = i + 1
        Loop
        Close 1
        ReDim Preserve records(UBound(records) - 1)

        ' ordina il vettore per data
        numRecord = UBound(records) + 1
        For j = 1 To numRecord - 1
            dataAppo = records(j).data
            numAppo = records(j).num
            i = j - 1
            Do While i >= 0
                If dataAppo < records(i).data Then
                    records(i + 1).data = records(i).data
                    records(i + 1).num = records(i).num
                    i = i - 1
                Else
                    Exit Do
                End If
            Loop
            records(i + 1).data = dataAppo
            records(i + 1).num = numAppo
        Next j
        
        ' cerca un backup di oggi
        j = 0
        Do While j <= numRecord - 1 And Not stessoGiorno
            If CDate(records(j).data) = date Then
                stessoGiorno = True
            End If
            j = j + 1
        Loop
        
        If stessoGiorno Then
            ' sovrascrive il backup dello stesso giorno senza modificare niente
            If Dir(lettera & ":\" & nomeVolume & records(j - 1).num) <> "" Then
                Kill lettera & ":\" & nomeVolume & records(j - 1).num
            End If
            GestisciFile = records(j - 1).num
        ElseIf numRecord = MAX_BACKUP Then
            ' elimina il meno recente
            If Dir(lettera & ":\" & nomeVolume & records(0).num) <> "" Then
                Kill lettera & ":\" & nomeVolume & records(0).num
            End If
            ' aggiorna i dati
            records(0).data = date
            GestisciFile = records(0).num
        Else
            ' aggiunge un nuovo backup
            ReDim Preserve records(UBound(records) + 1)
            numRecord = numRecord + 1
            records(numRecord - 1).data = date
            records(numRecord - 1).num = numRecord - 1
            GestisciFile = numRecord - 1
        End If
        
        ' salva
        Open lettera & ":\Dati.dat" For Random As 1
        For i = 0 To numRecord - 1
           Put 1, i + 1, records(i)
           'Debug.Print "data " & records(i).data & "    num " & records(i).num
        Next i
        Close 1
    Else
        records(0).num = 0
        records(0).data = date
        Open lettera & ":\Dati.dat" For Random As 1
        Put 1, 1, records(0)
        Close 1
        GestisciFile = 0
    End If
    Exit Function
gestione:
    MsgBox "Errore n° 3 - " & Err.Description, vbCritical, "Attenzione"
End Function

Private Sub RipristinaArchivio(lettera As String)
    Dim tempo As Single
    Dim numFile As Integer
    
    Screen.MousePointer = cc2Hourglass
    anmAvi.Open App.Path & "\clip.avi"
    anmAvi.Play
    tempo = Timer
    Do
        DoEvents
    Loop Until tempo + 1 <= Timer
    
    numFile = frmPeriferiche!flxGriglia.TextMatrix(frmPeriferiche!flxGriglia.Row, 0)
    If Dir(lettera & ":\" & nomeVolume & numFile) <> "" Then
        ' ripristina il file
        FileCopy lettera & ":\" & nomeVolume & numFile, structApri.pathVolume & "\" & nomeVolume & numFile
        ' elimina il vecchio database
        Kill structApri.pathVolume & "\" & nomeVolume
        Name structApri.pathVolume & "\" & nomeVolume & numFile As structApri.pathVolume & "\" & nomeVolume
        lblScritta = "Ripristino eseguito correttamente"
    Else
        MsgBox "Impossibile ripristinare. File inesistente", vbCritical, "Attenzione"
        lblScritta = "Ripristino non eseguito correttamente"
    End If
    anmAvi.Stop
    anmAvi.Visible = False
    Screen.MousePointer = 0
    Me.Height = 1485
    lblAttendi = "Premere un tasto per continuare"
    
    Me.SetFocus
    finito = True
End Sub

Private Sub BackupArchivio(lettera As String)
    ' effettua la copia del volume sul disco rimovibile
    On Error GoTo gestione
    Dim numFile As Integer
    numFile = GestisciFile(lettera)
    If SpazioSufficiente(lettera, FileLen(structApri.pathVolume & "\" & nomeVolume) / Megabyte) Then
        FileCopy structApri.pathVolume & "\" & nomeVolume, lettera & ":\" & nomeVolume & numFile
    Else
        MsgBox "Impossibile continuare" & vbCrLf & "Spazio insufficiente sull'unita' di backup", vbCritical, "Backup archivio"
    End If
    Exit Sub
gestione:
    If Err.Number = 70 Then
        MsgBox "Impossibile effettuare il backup" & vbCrLf & "Disconnettere prima gli altri utenti client", vbCritical, "Attenzione"
    Else
        MsgBox "Errore n°: 2 - " & Err.Description, vbCritical, "Attenzione"
    End If
    tDisconnetti = tpDANNULLA
    Unload Me
End Sub

Private Sub GestioneErrore()
    Dim strMsg As String
    Select Case Err.Number
        Case 68
            strMsg = "L'unità non è disponibile"
        Case 71
            strMsg = "Inserire un dichetto nel dispositivo "
        Case 57
            strMsg = "Errore interno del disco"
        Case 61
            strMsg = "Disco pieno"
        Case 76
            strMsg = "Percorso inesistente"
        Case 54
            strMsg = "Impossibile spostare l'archivio"
        Case 53
            strMsg = "File inesistente"
        Case 62
            strMsg = "Il file risulta danneggiato"
        Case Else
            strMsg = Err.Description
    End Select
    Screen.MousePointer = 0
    anmAvi.Stop
    If Err.Number <> 5 Then
        MsgBox "Errore n° 1 - " & strMsg, vbCritical, "Attenzione"
    End If
    Unload Me
End Sub
