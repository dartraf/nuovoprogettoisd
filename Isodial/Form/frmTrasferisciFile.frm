VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrasferisciFile 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   4330
      _ExtentX        =   7646
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmTrasferisciFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Megabyte = 1048576
Dim numClient As Integer

Private Type structFile
    data As Date
    num As Integer
End Type
Dim records() As structFile
Dim MAX_BACKUP As Integer

Private Sub Form_Load()
    Call TakeCloseOff(Me.hWnd)
End Sub

Private Sub Form_Activate()
   Dim lettera As String
   Me.Left = 5420
   Me.Top = 6900
    
    'ATTENZIONE - NON CAMBIARE L'ORDINE DEGLI ELSEIF
    If VerificaDiscoRimovibile(lettera) = False Then
           MsgBox "Impossibile effettuare il backup del database - CONNETTERE L'UNITA'", vbCritical, "UNITA' DI BACKUP NON PRESENTE"
           tDisconnetti = tpDANNULLA
 '   ElseIf tPeriferica = tpBACKUP = False Then
 '          ProgressBar1.Width = 3200
 '          Me.Left = 6000
 '          Me.Top = 6800
 '          Me.Width = 3390
 '          Me.BackColor = &H8000000A
 '          Call RipristinaArchivio(lettera)
    ElseIf SpazioSufficiente(lettera, FileLen(structApri.pathVolume & "\" & nomeVolume) / Megabyte) = False Then
           MsgBox "Impossibile continuare" & vbCrLf & "Spazio insufficiente sull'unita' di backup", vbCritical, "Backup Database"
           tDisconnetti = tpDANNULLA
    ElseIf nonCorrotto = False Then
           MsgBox "Impossibile procedere al backup" & vbCrLf & "Ripristinare un precedente backup o contattare l'autore", vbCritical, "ATTENZIONE!!! DATABASE CORROTTO"
           tDisconnetti = tpDANNULLA
    ElseIf nessunClient(numClient) = False Then
           If MsgBox("ATTENZIONE!!! Altri utenti sono connessi ad ISODIAL - Li disconnetto automaticamente?", vbQuestion + vbYesNo, "CONTROLLO UTENTI") = vbYes Then
      '       lblScritta = "Database integro. Backup in corso" & vbCrLf & "L'operazione potrebbe richiedere alcuni minuti"
              Call PulisciTabCLIENTI
              Call Copia(lettera)
           Else
              MsgBox "Disconnettere TUTTI gli utenti e riavviare il backup", vbCritical, "BACKUP ARCHIVIO"
              tDisconnetti = tpDANNULLA
           End If
    Else
      Call Copia(lettera)
    End If
    Unload Me
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'   If finito Then
'      Unload Me
'   End If
'End Sub

Private Sub CaricaMaxBackup()
  ' carica il numero di backup impostato dall'utente
    Dim rsDataset As New Recordset
    rsDataset.Open "IMPOSTAZIONI_BACKUP", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    MAX_BACKUP = rsDataset("NUMERO")
    Set rsDataset = Nothing
End Sub

Private Sub Copia(lettera As String)

    Dim tempo As Single

    On Error GoTo gestioneerror
    
   ' anmAvi.Open App.Path & "\clip.avi"
   ' anmAvi.Play
    
    tempo = Timer
    Do
     DoEvents
    Loop Until tempo + 1 <= Timer
    
    Call CaricaMaxBackup
    Call BackupArchivio(lettera)
    
'    anmAvi.Stop
'    anmAvi.Visible = False
    
    Call BloccoCentri
    Me.SetFocus
    If SpegniPc Then
        tempo = Timer
        Do
            DoEvents
        Loop Until tempo + 5 <= Timer
        Call Spegni
    End If
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

Private Sub BackupArchivio(lettera As String)
  ' effettua la copia del volume sull'unità di backup (disco rimovibile)
    On Error GoTo gestione
    Dim ret As Double
    Dim numFile As Integer
    
    numFile = GestisciFile(lettera)
    Screen.MousePointer = cc2Hourglass
        
  ' chiude la connessione
    Set cnPrinc = Nothing
    Set cnTrac = Nothing
  ' chiude la condivisione
    Call Shell("NET SHARE RISORSA /DELETE", vbHide)
  ' smonta il volume
    ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)

    Dim fileorigine As String
    Dim filedestinazione As String
    fileorigine = structApri.pathVolume & "\" & nomeVolume
    filedestinazione = lettera & ":\" & nomeVolume & numFile

    Call CopiaFile(fileorigine, filedestinazione, ProgressBar1)
    Screen.MousePointer = 0
    Exit Sub
gestione:
    If Err.Number = 70 Then
        MsgBox "IMPOSSIBILE EFFETTUARE IL BACKUP - Altri utenti sono connessi ad ISODIAL" & vbCrLf & "Disconnetterli TUTTI e riavviare il backup", vbCritical, "ATTENZIONE!!!"
    Else
        MsgBox "Errore n°: 2 - " & Err.Description, vbCritical, "ATTENZIONE!!!"
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
    If Err.Number <> 5 Then
        MsgBox "Errore n° 1 - " & strMsg, vbCritical, "Attenzione"
    End If
    Unload Me
End Sub
