Attribute VB_Name = "modPrincipale"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modPrincipale.bas
'
' <b>Descrizione</b>: Funzioni di avvio e controllo
'
' @remarks
'
' @author
'
' @date 28/01/2011 18.09
Option Explicit

Const strKeyDb As String = "ciao"                               '  accesso al db Centro
Const strKeyDbTrac As String = "limone"                         '  accesso al db Connessioni
Public Const strKeyVolume As String = "!&^_`}~804FGHUJK!'^"            '  volume

Public Const appName As String = "IsoDial"             ' nomi delle chiavi del registro
Const sezione As String = "Impostazioni"

Const nomeUsb As String = "U_BACKUP"                     ' nome disco rimovibile
Public Const nomeVolume As String = "SysCrypt.tc"

Const Megabyte = 1048576

Private Type structFile
    data As Date
    num As Integer
End Type

''
' Funzione di avvio
'
' @param
' @param
' @return
' @remarks
Sub Main()
    Dim ret As Long
    Dim datadb As Date
    
    Call CaricaPercorso
    Call VerificaErrori
  '  Call ControlloFileEsterni    controlla la versione delle librerie
    Call MontaVolume
    Call CaricaDati
    Call VerificaFunzionalita
    ' verifica che il db non sia corrotto
    If Not nonCorrotto Then
        MsgBox "Impossibile procedere" & vbCrLf & "Ripristinare un precedente backup o richiedere l'intervento tecnico" & vbCrLf & "Accesso consentito al solo amministratore di sistema", vbCritical, "Database corrotto"
        isCorrotto = True
    Else
        isCorrotto = False
    End If
    ' controlla che la data di sistema sia quella corrente
    datadb = CDate(Left(FileDateTime(structApri.pathDB + "\centro.mdb"), 10))
'    datadb = DateValue(Month(datadb) & "/" & Day(datadb) & "/" & Year(datadb))
    If datadb > date Then
        MsgBox ("IMPOSSIBILE AVVIARE ISODIAL - La data di sistema non è corretta"), vbCritical, "ATTENZIONE!!!"
        ' chiude la connessione
        Set cnPrinc = Nothing
        Set cnTrac = Nothing
        ' chiude la condivisione
        Call Shell("NET SHARE RISORSA /DELETE", vbHide)
        ' smonta il volume
        ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)
        End
    End If

    frmMain.Show
    frmLogin.Show 1
End Sub

''
' Carica i path dal registro di sistema
'
' @param
' @param
' @return
' @remarks
Private Sub CaricaPercorso()
    On Error GoTo gestione
    structApri.pathVolume = (GetSetting(appName, sezione, "percorsoVolume"))
    structApri.pathTrueCrypt = (GetSetting(appName, sezione, "percorsoTrueCrypt"))
    structApri.pathExe = (GetSetting(appName, sezione, "percorsoExe"))
    structApri.server = CBool(GetSetting(appName, sezione, "Server"))
    structApri.nomeServer = GetSetting(appName, sezione, "nomeServer")
    structApri.pathNomeCertificato = GetSetting(appName, sezione, "nomeCertificato")
    structApri.strFromModuliWord = GetSetting(appName, sezione, "strFromModuliWord")
    Exit Sub
gestione:
    ' identificativo errore 1-
    MsgBox "Errore n° 1-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
    End
End Sub

''
' Verifica gli errori all'avvio
'
' @param
' @param
' @return
' @remarks controlla se il software è gia aperto
Public Sub VerificaErrori()
    On Error GoTo gestione
    Dim lettera As String
    Dim dimVolume As Double
    
    If App.PrevInstance Then
        MsgBox "Il programma è già in esecuzione" & vbCrLf & _
               "(Situato nella barra in basso a destra dello schermo vicino all'orologio.)", vbCritical, "Attenzione"
        End
    End If
    ' verifica la presenza del file
    If Dir((structApri.pathVolume) & "\" & nomeVolume) = "" Then
        MsgBox "Archivio inesistente", vbCritical, "Apertura archivio"
        End
    End If
    dimVolume = FileLen(structApri.pathVolume & "\" & nomeVolume) / Megabyte
    If structApri.server Then
        ' verifica la presenza di TrueCrypt
        If Dir((structApri.pathTrueCrypt) & "\TrueCrypt.exe") = "" Then
            MsgBox "Programma di criptaggio non istallato", vbCritical, "Apertura archivio"
            End
        End If
        ' verifica la presenza della penna
        If Not (Environ$("COMPUTERNAME") = "MASTER" Or Environ$("COMPUTERNAME") = "MASTERMIO") Then
            If Not VerificaDiscoRimovibile(lettera) Then
                MsgBox "Impossibile continuare" & vbCrLf & "Unita' di backup mancante", vbCritical, "Apertura archivio"
                End
            End If
            If Not SpazioSufficiente(lettera, dimVolume) Then
                MsgBox "Impossibile continuare" & vbCrLf & "Spazio insufficiente sull'unita' di backup", vbCritical, "Apertura archivio"
                End
            End If
            If Not backupValidi(lettera) Then
                MsgBox "I file di backup presenti nell'unità esterna non hanno superato il controllo di coerenza" & vbCrLf & "Contattare l'amministratore di sistema", vbCritical, "Attenzione"
            End If
        End If
    End If
    Exit Sub
gestione:
    If Err.Number = 55 Or Err.Number = 53 Then
        Exit Sub
    ElseIf Err.Number = 52 Then
        MsgBox "Impossibile avviare Isodial" & vbCrLf & "Verificare la connessione al server", vbCritical, "Attenzione"
        End
    Else
        ' identificativo errore 2-
        MsgBox "Errore n° 2-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
        End
    End If
End Sub

''
' Monta il volume dal disco cryptato e lo condivide se PC server
'
' @param
' @param
' @return
' @remarks PC client si connette a condividi.exe
Public Sub MontaVolume()
    Dim ret As Double
    On Error GoTo gestione
    If structApri.server Then
        ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe" & " /v " & structApri.pathVolume & "\" & nomeVolume & " /l X  /p " & strKeyVolume & " /a /q /s", vbHide)
        ' deve condividere la risorsa
        Shell ("NET SHARE RISORSA=X: /UNLIMITED")
        structApri.pathDB = "X:"
    Else
        tRete = tpCONNETTI
        frmAttendi.Show 1
        ' il volume e gia montato
        structApri.pathDB = structApri.nomeServer & "\RISORSA"
    End If
    Exit Sub
gestione:
    ' identificativo errore 3-
    MsgBox "Errore n° 3-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
    End
End Sub

''
' Compatta il db
'
' @param nomeDB nome del db
' @param strPercorsoDB path del db
' @param strKeyDb password del db
' @return
' @remarks non utilizzata perche corrompe il db
Public Sub CompattaDB(nomeDB As String, strPercorsoDB As String, strKeyDb As String)
     On Error GoTo ErrorHandler

     Dim strFileTemporaneo As String
     Dim oJet As JRO.JetEngine

     Set oJet = New JRO.JetEngine

     ' Determina il nome di un file temporaneo nella dir dell'applicazione
     strFileTemporaneo = strPercorsoDB & "\temp.mdb"

 oJet.CompactDatabase _
    "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" & strPercorsoDB & "\" & nomeDB & ";Jet OLEDB:Database Password=" & strKeyDb, _
    "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" & strFileTemporaneo & ";" _
    & "Jet OLEDB:Engine Type = 5;Jet OLEDB:Database Password=" & strKeyDb

     Kill strPercorsoDB & "\" & nomeDB

     Name strFileTemporaneo As strPercorsoDB & "\" & nomeDB
     Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
        Exit Sub
    Else
        ' identificativo errore 4-
        MsgBox "Errore n° 4-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
    End If
End Sub

''
' Aprele connessioni e carica i dati iniziali
'
' @param
' @param
' @return
' @remarks
Public Sub CaricaDati()
    On Error GoTo gestione
    Dim rsDataset As Recordset
    ' connessione principale su Centro
    Set cnPrinc = New ADODB.Connection
    cnPrinc.CursorLocation = adUseClient 'adUseServer
    strConnectionStringCentro = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (structApri.pathDB) & "\Centro.mdb;Jet OLEDB:Database Password=" & strKeyDb
    cnPrinc.Open strConnectionStringCentro
    If TRACCIATO Then
        ' connessione su Connessioni
        Set cnTrac = New ADODB.Connection
        cnTrac.CursorLocation = adUseClient 'adUseServer
        strConnectionStringTracciatura = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (structApri.pathDB) & "\Connessioni.mdb;Jet OLEDB:Database Password=" & strKeyDbTrac
        cnTrac.Open strConnectionStringTracciatura
    End If
    
    If Not structApri.server Then
        ' si aggiunge alla lista dei client collegati
        Set rsDataset = New Recordset
        rsDataset.Open "CLIENT", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
        rsDataset.Update "NUMERO", rsDataset("NUMERO") + 1
        Set rsDataset = Nothing
    End If
    intValore = 10
    
    ' carica le var publiche
    Call CaricaVarPublic
    
    Exit Sub
gestione:
    ' identificativo errore 5-
    MsgBox "Errore n° 5-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
    End
End Sub

''
' Abilita la gestione dei rimborsi
'
' @param
' @param
' @return
' @remarks
Private Sub VerificaFunzionalita()
'If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Or structIntestazione.sCodiceSTS = CODICESTS_BARTOLI Then
   structApri.F1abiliata = True
'Else
'   structApri.F1abiliata = False
'End If
    
'    On Error GoTo gestione
'    Dim rsDataset As New Recordset
'    Dim appo As String
'    Dim codiceSTS As String

'    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
'    codiceSTS = rsDataset("CODICE_STS")
'    rsDataset.Close
    
'    If Dir(structApri.pathExe & "\impostazioni.dat") <> "" Then
'        Open structApri.pathExe & "\impostazioni.dat" For Input As #1
'        If Not EOF(1) Then
'            Line Input #1, appo
'        End If
'        Close #1
'    End If
    
'    If appo = "" Then
'        structApri.F1abiliata = False
'    Else
'        If CInt(Mid(appo, 11, 4)) = CInt(Mid(codiceSTS, 3, 4) + 1111) Then
'            structApri.F1abiliata = True
'        Else
'            structApri.F1abiliata = False
'        End If
'    End If
    
'    Exit Sub
'gestione:
    ' identificativo errore 7-
'    MsgBox "Errore n° 7-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
'    End

End Sub

''
' Controlla che ci sia coerenza tra le info del file Dati.dat e i file di backup presenti nella penna
'
' @param lettera lettera dove è presente la penna (es. E)
' @param
' @return true se i backup sono validi, altrimenti false
' @remarks
Private Function backupValidi(lettera As String) As Boolean
    On Error GoTo gestione

    Dim i As Integer
    Dim records() As structFile
    
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
        
        For i = 0 To UBound(records)
            If Dir(lettera & ":\" & nomeVolume & records(i).num) = "" Then
                backupValidi = False
                Exit Function
            End If
        Next i

        backupValidi = True
    Else
        backupValidi = True
    End If
    Exit Function
gestione:
    MsgBox "Errore n° 6 - " & Err.Description, vbCritical, "Attenzione"
End Function

'' Confronta le versioni richieste con quelle dei file presenti sul pc
Private Function ConfrontoVersione(inLibreria As String, inVersioneRichiesta As String, ByRef outTesto As String) As Boolean
    Dim strVersioneAttuale As String

    ConfrontoVersione = True
    If Not IsCorrectVersion(inVersioneRichiesta, inLibreria, strVersioneAttuale) Then
        outTesto = outTesto & _
                            "La libreria " & inLibreria & " non è aggiornata." & vbCrLf & _
                            "Versione richiesta: " & inVersioneRichiesta & Space(5) & "Versione attuale: " & strVersioneAttuale & vbCrLf
        ConfrontoVersione = False
    End If
End Function

'' Controlla la coerenza dei file esterni (dll, ocx)
Private Sub ControlloFileEsterni()
    Dim blnBloccaProgramma As Boolean
    Dim strTesto As String
    Dim strVersioneAttuale As String
    Dim strVersioneRichiesta As String
    Dim strLibreria As String
    
    strLibreria = "DataTimeBox.ocx"
    strVersioneRichiesta = "1.03.0007"
    blnBloccaProgramma = Not ConfrontoVersione(strLibreria, strVersioneRichiesta, strTesto)
    
    strLibreria = "SuperTextBox.ocx"
    strVersioneRichiesta = "1.01.0003"
    blnBloccaProgramma = blnBloccaProgramma Or Not ConfrontoVersione(strLibreria, strVersioneRichiesta, strTesto)
    
    strLibreria = "ACPRibbon.ocx"
    strVersioneRichiesta = "1.00.0001"
    blnBloccaProgramma = blnBloccaProgramma Or Not ConfrontoVersione(strLibreria, strVersioneRichiesta, strTesto)
    
    strLibreria = "DataComboBox.ocx"
    strVersioneRichiesta = "1.00.0001"
    blnBloccaProgramma = blnBloccaProgramma Or Not ConfrontoVersione(strLibreria, strVersioneRichiesta, strTesto)
    
    
    If blnBloccaProgramma Then
        Beep
        Load frmControlloFileEsterni
        strTesto = "Impossibile avviare Isodial. " & vbCrLf & "Si prega di contattare l'autore." & vbCrLf & vbCrLf & strTesto
        frmControlloFileEsterni.lblTesto.Caption = strTesto
        frmControlloFileEsterni.Show 1
        Unload frmControlloFileEsterni
    End If
End Sub

''
' Verifica se sulla penna "lettera" c'è spazio disponibile>spazio
'
' @param lettera lettera dove è presente la penna (es. E)
' @param spazio spazio minimo richiesto
' @return true se c'è lo spazio, altrimenti false
' @remarks
Public Function SpazioSufficiente(lettera As String, spazio As Double) As Boolean
    SpazioSufficiente = CLng(GetDriveSize(lettera & ":")) > spazio
End Function

''
' Analizza lo spazio disponibile sulla penna
'
' @param DriveName etichetta della penna
' @param
' @return spazio disponibile
' @remarks
Public Function GetDriveSize(DriveName As String) As String
    Dim FB As Currency, BT As Currency, FBT As Currency
    Dim RetVal As Long
    RetVal = GetDiskFreeSpace_FAT32(Left(DriveName, 2), FB, BT, FBT)
    FBT = FBT * 10000 'convert result To actual size In bytes
    GetDriveSize = Format(FBT / Megabyte, "####,###,###")
End Function

''
' Verifica la presenza della penna
'
' @param lettera lettera dove è presente la penna (es. E)
' @param
' @return
' @remarks
Public Function VerificaDiscoRimovibile(ByRef lettera As String) As Boolean
    Dim ret As Long
    Dim allDrives As String
    Dim v_drives() As String
    Dim i As Integer
    
    Dim volName As String
    Dim serial As Long
    Dim f As String
    Dim g As Long
    
    'get the list of all available drives
    allDrives = VBGetLogicalDriveStrings()
    v_drives = Split(allDrives, Chr(0))
    
    For i = 0 To UBound(v_drives)
        ret = GetDriveType(v_drives(i))
        If ret = DRIVE_REMOVABLE Then
            If Left(v_drives(i), 1) <> "A" And Left(v_drives(i), 1) <> "B" Then
                Call GetDriveInfo(Left(v_drives(i), 1) & ":", volName, serial, f, g)
                'MsgBox volName & vbTab & serial & vbTab & f & vbTab & g
                If volName = nomeUsb Then
                    lettera = Left(v_drives(i), 1)
                    VerificaDiscoRimovibile = True
                    Exit Function
                End If
            End If
        End If
    Next i
    VerificaDiscoRimovibile = False
End Function

''
' Trova la lettera associata alla penna
'
' @param
' @param
' @return stringa di lettere associate ai disco rimovibili separati da chr$(0)
' @remarks
Private Function VBGetLogicalDriveStrings() As String
    Dim r As Long
    Dim tmp As String
    
    tmp$ = Space$(64)
    
    r& = GetLogicalDriveStrings(Len(tmp$), tmp$)
    
    VBGetLogicalDriveStrings = Trim$(tmp$)
End Function

Private Function GetDriveInfo(ByVal DriveName As String, Optional VolumeName As String, _
        Optional SerialNumber As Long, Optional FileSystem As String, _
        Optional FileSystemFlags As Long) As Boolean
    
    Dim ignore As Long
    
    ' if it isn't a UNC path, enforce the correct format
    If InStr(DriveName, "\\") = 0 Then
        DriveName = Left$(DriveName, 1) & ":\"
    End If
    
    ' prepare receiving buffers
    SerialNumber = 0
    FileSystemFlags = 0
    VolumeName = String$(MAX_PATH, 0)
    FileSystem = String$(MAX_PATH, 0)
    
    ' The API function return a non-zero value if successful
    GetDriveInfo = GetVolumeInformation(DriveName, VolumeName, Len(VolumeName), _
        SerialNumber, ignore, FileSystemFlags, FileSystem, Len(FileSystem))
    ' drop characters in excess
    VolumeName = Left$(VolumeName, InStr(VolumeName, vbNullChar) - 1)
    FileSystem = Left$(FileSystem, InStr(FileSystem, vbNullChar) - 1)
    
End Function

