Attribute VB_Name = "modFunctions"
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Modulo - modFunctions.bas
'
' <b>Descrizione</b>: Funzioni publiche
'
' @remarks
'
' @author
'
' @date 27/01/2011 19.06

Option Explicit

''
' Verifica se la voce passata accetta i valori pos e neg
'
' @param voce voce su cui ricercare i valori pn
' @param
' @return true se la voce vale positivo-negativo, altrimenti false
' @remarks
Public Function AccettaPN(voce As String) As Boolean
    Dim rsDataset As Recordset
    If voce = "" Then
        AccettaPN = False
        Exit Function
    End If
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM VOCI_ESAMI WHERE NOME='" & Apostrophe(voce) & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        MsgBox "Errore nel caricamento dei dati", vbCritical, "Impossibile aggiornare"
        AccettaPN = False
    Else
        AccettaPN = CBool(rsDataset("PN"))
    End If
    Set rsDataset = Nothing
End Function

''
' Annulla le var di stato del paziente
'
' @param
' @param
' @return
' @remarks le variabili sono publiche
Public Sub AnnullaVarStato()
    Dim i As Integer
    For i = 1 To 3
        statoPaziente.dataArrivi(i) = ""
        statoPaziente.dataPartenza(i) = ""
        ' nessun valore nella cbo vale -1
        statoPaziente.centriProv(i) = -1
    Next i
    statoPaziente.dataStato = ""
    statoPaziente.donatore = 2        ' non immesso
End Sub

''
' Restituisce una stringa eliminando il problema dell'apice
'
' @param sFieldString stringa da analizzare
' @param
' @return stringa modificata
' @remarks utile per le query SQL
Public Function Apostrophe(sFieldString As String) As String
    If InStr(sFieldString, "'") Then
        Dim iLen As Integer
        Dim ii As Integer
        Dim apostr As Integer
        iLen = Len(sFieldString)
        ii = 1
        Do While ii <= iLen
            If Mid(sFieldString, ii, 1) = "'" Then
                apostr = ii
                sFieldString = Left(sFieldString, apostr) & "'" & _
                    Right(sFieldString, iLen - apostr)
                iLen = Len(sFieldString)
                ii = ii + 1
            End If
            ii = ii + 1
        Loop
    End If
    Apostrophe = sFieldString
End Function

'' Effettua l'autosize delle colonne della griglia
Public Sub AutoResizeGrid(inGrid As MSFlexGrid, inForm As Form, Optional inColonnaDaAllargare As Integer = -1, Optional inColIniziale As Integer = 1)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim decColWidthTotale As Single
    
    Dim oFont As Font
    Dim decSize As Single
    Set oFont = inForm.Font
    decSize = inForm.FontSize
    inForm.FontSize = inGrid.CellFontSize
    inForm.Font = inGrid.Font
    With inGrid
        For intCol = inColIniziale To .Cols - 1
            .ColWidth(intCol) = 0
            For intRow = 0 To .Rows - 1
                If .ColWidth(intCol) < inForm.TextWidth(.TextMatrix(intRow, intCol)) + 100 Then
                   .ColWidth(intCol) = inForm.TextWidth(.TextMatrix(intRow, intCol)) + 100
                End If
            Next
            decColWidthTotale = decColWidthTotale + .ColWidth(intCol)
        Next
        If decColWidthTotale < inGrid.Width Then
            If inColonnaDaAllargare <> -1 Then
                .ColWidth(inColonnaDaAllargare) = .ColWidth(inColonnaDaAllargare) + inGrid.Width - decColWidthTotale - 360
            Else
                Dim decWidthForRow As Single
                decWidthForRow = (inGrid.Width - decColWidthTotale - 380) / (.Cols - inColIniziale)
                For intCol = inColIniziale To .Cols - 1
                    .ColWidth(intCol) = .ColWidth(intCol) + decWidthForRow
                Next intCol
            End If
        End If
    End With
    inForm.Font = oFont
    inForm.FontSize = decSize

    Set oFont = Nothing
End Sub

''
' Ordina un vettori di interi
'
' @param MioArray() vettore da ordinare
' @param
' @return
' @remarks
Public Sub BubbleSort(ByRef MioArray() As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim flag As Boolean
    Dim Temp As Integer
    flag = True
    i = UBound(MioArray, 1)
    Do While (i <> LBound(MioArray, 1) And flag = True)
        flag = False
        For j = LBound(MioArray, 1) To i - 1
            If MioArray(j) > MioArray(j + 1) Then
                Temp = MioArray(j)
                MioArray(j) = MioArray(j + 1)
                MioArray(j + 1) = Temp
                flag = True
            End If
        Next j
        i = i - 1
    Loop
End Sub

'' Gestione apertura form con caricamento automatico dei pazienti
Public Function CaricaPazienteInAperturaForm(inMeCaption As String, inModificato As Boolean, ByRef outPazientiKey As Integer) As enumTipoApertura
    Dim blnProsegui As Boolean
    
    CaricaPazienteInAperturaForm = tpTipoAperturaNone
    If outPazientiKey = 0 Then
        If STESSO_PAZIENTE Then
            If oPazientiKey.intNumeroFormAperti > 0 Then
                outPazientiKey = oPazientiKey.intPazientiKey
                CaricaPazienteInAperturaForm = tpCaricaPaziente
            Else
                CaricaPazienteInAperturaForm = tpTrovaPaziente
            End If
        Else
            CaricaPazienteInAperturaForm = tpTrovaPaziente
        End If
    Else
        If STESSO_PAZIENTE Then
            If oPazientiKey.intNumeroFormAperti > 0 Then
                If outPazientiKey <> oPazientiKey.intPazientiKey Then
                    blnProsegui = False
                    If inModificato Then
                        If MsgBox("Le modifiche apportate alla scheda non sono state salvate." & vbCrLf & "Caricare i dati del paziente " & oPazientiKey.GetPazienteInfo & "?", vbQuestion + vbYesNo, inMeCaption) = vbYes Then
                            blnProsegui = True
                        End If
                    Else
                        blnProsegui = True
                    End If
                    If blnProsegui Then
                        outPazientiKey = oPazientiKey.intPazientiKey
                        CaricaPazienteInAperturaForm = tpCaricaPaziente
                    End If
                End If
            Else
                Call oPazientiKey.ImpostaPazientiKey(outPazientiKey, inMeCaption)
            End If
        End If
    End If
End Function

Public Sub CaricaPso()
    On Error Resume Next
    Dim rsDataset As New Recordset
    Dim i As Integer
    
    If date < CDate(laData) Then
        For i = 1 To 3
            statoPaziente.dataArrivi(i) = ""
            statoPaziente.dataPartenza(i) = ""
            ' nessun valore nella cbo vale -1
            statoPaziente.centriProv(i) = -1
        Next i
    Else
        rsDataset.Open "SELECT * FROM ANAMNESI_ESAMI WHERE CODICE_GRUPPO=1", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            rsDataset.Delete
            rsDataset.Update
        End If
        rsDataset.Close
    End If
    
End Sub

''
' Carica le variabili per la intestazione di stampa
'
' @param
' @param
' @return
' @remarks pone questi valori in una struttura publica
Public Sub CaricaVarPublic()
    Dim rsDataset As New Recordset
    rsDataset.Open "INTESTAZIONE_STAMPA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        structIntestazione.sRagione = rsDataset("RAGIONE_SOCIALE")
        structIntestazione.sTipoAmbulatorio = rsDataset("TIPO")
        structIntestazione.sIndirizzo = rsDataset("INDIRIZZO")
        structIntestazione.sCap = rsDataset("CAP")
        structIntestazione.sCitta = rsDataset("CITTA")
        structIntestazione.sProv = rsDataset("PROV")
        structIntestazione.sTelefono = IIf(rsDataset("TELEFONO") = "", " - - ", rsDataset("TELEFONO"))
        structIntestazione.sFax = IIf(rsDataset("FAX") = "", " - - ", rsDataset("FAX"))
        structIntestazione.sCodiceFiscale = IIf(rsDataset("CODICE_FISCALE") = "", " - - ", rsDataset("CODICE_FISCALE"))
        structIntestazione.sIva = IIf(rsDataset("IVA") = "", " - - ", rsDataset("IVA"))
        structIntestazione.sMail = IIf(rsDataset("MAIL") = "", " - - ", rsDataset("MAIL"))
        structIntestazione.sCodiceSTS = rsDataset("CODICE_STS") & ""
        structIntestazione.sCodiceAsl = rsDataset("CODICE_ASL")
        structIntestazione.sCodiceDistretto = rsDataset("CODICE_DISTRETTO")
        structIntestazione.sLogoISO = CBool(rsDataset("LOGO"))
        structIntestazione.sLogoQualita = CBool(rsDataset("LOGO_QUALITA"))
        structIntestazione.sLogoAziendale = CBool(rsDataset("LOGO_AZIENDALE"))
        structIntestazione.sNomeLogoISO = rsDataset("NOME_LOGOISO")
        structIntestazione.sNomeLogoQualita = rsDataset("NOME_LOGOQUALITA")
        structIntestazione.sNomeLogoAziendale = rsDataset("NOME_LOGOAZIENDALE")
        frmMain.staBar.Panels(2) = structIntestazione.sRagione
        frmMain.staBar.Panels(3) = structIntestazione.sCitta
    End If
    Set rsDataset = Nothing
    
    
    Dim rsDatasetDirettoreSanitario As New Recordset
    rsDatasetDirettoreSanitario.Open "DIRETTORE_SANITARIO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    If Not (rsDatasetDirettoreSanitario.EOF And rsDatasetDirettoreSanitario.BOF) Then
        structIntestazione.sDirettoreSanitarioNome = rsDatasetDirettoreSanitario("NOME")
        structIntestazione.sDirettoreSanitarioCognome = rsDatasetDirettoreSanitario("COGNOME")
    End If
    Set rsDatasetDirettoreSanitario = Nothing

End Sub

''
' Colora l'intera riga di una flexGrid
'
' @param inFlex griglia
' @param numColMax il numero massimo di colonne da colorare
' @param colore il colore della riga
' @return
' @remarks numColMax è quasi sempre inFlex.Cols - 1
Public Sub ColoraFlx(ByRef inFlex As Object, numColMax As Integer, Optional discolora As Boolean = False, Optional colore As ColorConstants = vbCyan, Optional blnAllRows As Boolean = False)
    Dim i As Integer
    Dim k As Integer
    Dim riga As Integer
    Dim Col As Integer
    Dim colAppo As ColorConstants
    
    If inFlex.Row = inFlex.FixedRows - 1 Then Exit Sub
    riga = inFlex.Row
    Col = inFlex.Col
    ' discolora la riga colorata
    For k = inFlex.FixedRows To inFlex.Rows - 1
        inFlex.Row = k                     ' per selezionare la riga successiva
        inFlex.Col = inFlex.FixedCols         ' per essere sicuri di selezionare la cella giusta
        ' utilizzo un var di appoggio perche cosi funziona
        colAppo = inFlex.CellBackColor
        If colAppo = colore And colAppo <> 0 Then
            For i = inFlex.FixedCols To numColMax
                inFlex.Col = i
                inFlex.CellBackColor = vbWhite
            Next i
            If Not blnAllRows Then Exit For
        End If
    Next k
    
    If Not discolora Then
        inFlex.Row = riga
        ' cambia colore della riga
        For i = inFlex.FixedCols To numColMax
            inFlex.Col = i
            inFlex.CellBackColor = colore
        Next i
        inFlex.Col = Col
    End If
End Sub

''
' Colora il vettore di optionButton selezionato
'
' @param obj optionButton selezionato
' @param numSel index del selezionato
' @param numMax numero totali di elementi nel vettore
' @param colSel colore del selezionato
' @param colNotSel colore dei altri
' @return
' @remarks
Public Sub ColoraSel(ByRef obj As Object, numSel As Integer, numMax As Integer, Optional colSel As Long = &HFF&, Optional colNotSel As Long = &H808080)
    ' Colora l option button selezionato
    Dim i As Integer
    For i = 0 To numMax - 1
        obj(i).ForeColor = colNotSel
    Next i
    obj(numSel).ForeColor = colSel
End Sub

''
' Cambia il colore di avvertimento nei form degli esami di lab.
'
' @param flx griglia
' @param vRow riga il cui valore è stato modificato o inserito
' @param vCol colonna il cui valore è stato modificato o inserito
' @param val valore nuovo
' @param valMax valore massimo
' @param valMin valore minimo
' @return
' @remarks imposta il colore bianco se val=-1 cioè il valore è stato eliminato
Public Sub ColoreDiAvviso(flx As Object, vRow As Integer, vCol As Integer, val As Single, valMax As Single, valMin As Single)
    With flx
        .Col = vCol
        .Row = vRow
        If val = -1 Then
            .CellBackColor = vbWhite
        ElseIf CDbl(val) < CDbl(valMin) Then
            .CellBackColor = vbYellow
        ElseIf CDbl(val) > CDbl(valMax) Then
            .CellBackColor = vbRed
        ElseIf CDbl(val) >= CDbl(valMin) And CDbl(val) <= CDbl(valMax) Then
            .CellBackColor = vbGreen
        End If
    End With
End Sub

'' Controlla le modifiche al form prima della chiusura
Public Function ControlloChiusuraForm(inModificato As Boolean, inMeCaption As String) As Boolean
    If inModificato Then
        If MsgBox("ATTENZIONE!!! Non sono state memorizzate le variazioni apportate alla scheda - VUOI ANNULLARLE?", vbQuestion + vbYesNo + vbDefaultButton2, inMeCaption) = vbYes Then
            ControlloChiusuraForm = True
        Else
            ControlloChiusuraForm = False
        End If
    Else
        ControlloChiusuraForm = True
    End If
End Function

''
' Effettua un controllo sul numero (principalmente sui puntini)
'
' @param numero numero da analizzare
' @param
' @return false se è un numero, altrimenti true
' @remarks
Public Function ControlloNumerico(numero As String) As Boolean
    Dim i As Integer
    Dim numOcc As Integer
    If InStr(1, numero, ".", vbTextCompare) = 0 Then
        ' il punto non c'è
        ControlloNumerico = False
    ElseIf InStr(1, numero, ".", vbTextCompare) = 1 Or InStr(1, numero, ".", vbTextCompare) = Len(numero) Then
        ' il punto è alla prima o ultima occorrenza
        ControlloNumerico = True
        MsgBox "Inserire correttamente il valore numerico", vbCritical, "Attenzione"
    Else
        ' verifica il numero di occorrenze
        For i = 1 To Len(numero)
            If Mid(numero, i, 1) = "." Then
                numOcc = numOcc + 1
            End If
        Next i
        If numOcc > 1 Then
            ControlloNumerico = True
            MsgBox "Inserire correttamente il valore numerico", vbCritical, "Attenzione"
        Else
            ControlloNumerico = False
        End If
    End If
End Function

''
' Carica i dati del medico per la stampa
'
' @param codice key del record della tabella MEDICI_REFERTANTI
' @return restituisce "cognome nome" del medico
' @remarks
Private Function DatiMedico(codice As Integer) As String
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COGNOME,NOME FROM MEDICI_REFERTANTI WHERE KEY=" & codice, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        DatiMedico = rsDataset("COGNOME") & " " & rsDataset("NOME")
    Else
        DatiMedico = ""
    End If
    Set rsDataset = Nothing
End Function

''
' Chiade l'eliminazione della voce dalla comboBox
'
' @param cbo comboBox
' @param tabella nome della tabella a cui è associata la cbo
' @return
' @remarks se si chiama la funzione EliminaVoce
Public Sub EliminaFromCbo(cbo As ComboBox, tabella As String)
    If cbo.Text = "" Then
        MsgBox "Selezionare la voce da eliminare", vbCritical, "Attenzione"
    Else
        If MsgBox("Sei sicuro di voler cancellare la voce: """ & cbo.Text & """ dall'archivio?", vbExclamation + vbYesNo, "Cancella voce") = vbYes Then
            ' elimina la voce
            Call EliminaVoce(cbo, tabella)
        End If
    End If
End Sub

'' Elimina eventuali scansioni rimaste in sospeso perche non si è memorizzata la scheda
Public Sub EliminaScansioniSospese(nomeTabella As String)
    Dim nomeFile As String
    Dim rsDataset As Recordset
    
    ' controlla eventuali scansioni memorizzate in sospeso
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE CODICE_SCHEDA=0", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    Do While Not rsDataset.EOF
        nomeFile = rsDataset("NOME_FILE")
        rsDataset.Delete
        rsDataset.MoveNext
        If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
            Kill structApri.pathDB & "\" & nomeFile & ".jpg"
        ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
            Kill structApri.pathDB & "\" & nomeFile & ".pdf"
        End If
    Loop
    rsDataset.Close
    Set rsDataset = Nothing
End Sub

''
' Elimina la voce dalla tabella e ricarica la cbo
'
' @param cbo comboBox
' @param tabella nome della tabella a cui è associata la cbo
' @return
' @remarks il nome del campo è sempre NOME
Public Sub EliminaVoce(cbo As ComboBox, tabella As String)
    Dim rsDataset As New Recordset
    rsDataset.Open "SELECT * FROM " & tabella & " WHERE NOME='" & UCase(Apostrophe(cbo.Text)) & "'", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        rsDataset.Delete
        ' elimina dalla cbo
        cbo.Clear
        rsDataset.Close
        rsDataset.Open tabella, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
        Do While Not rsDataset.EOF
            cbo.AddItem rsDataset("NOME")
            rsDataset.MoveNext
        Loop
    Else
        MsgBox "Elemento non presente in archivio", vbCritical, "Eliminazione"
    End If
    Set rsDataset = Nothing
End Sub

''
' Verifica se esiste un valore nella flx
'
' @param flx griglia
' @param col colonna da analizzare
' @param row riga da non prendere in considerazione
' @param nome valore da cercare
' @return false se non esiste, altrimenti ritorna la riga dove è presente il valore
' @remarks
Public Function Esiste(flx As MSFlexGrid, Col As Integer, Row As Integer, ByVal nome As String) As Integer
    Dim i As Integer
    For i = 1 To flx.Rows - 1
        If UCase(flx.TextMatrix(i, Col)) = UCase(nome) And i <> Row Then
            Esiste = i
            Exit Function
        End If
    Next i
    Esiste = 0
End Function

''
' Esegue un comando da linea di comando aspettando che finisca
'
' @param cmdline comanda da lanciare su linea
' @param
' @return idProces dell'ultima esecuzione
' @remarks viene usata per lanciare il modulo openssl.exe e aspettare che finisca ogni singola chiamata
Public Function ExecCmd(cmdline$) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    start.cb = Len(start)
    start.wShowWindow = 0
    start.dwFlags = 1
    ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)
    ExecCmd = ret
End Function

''
' Gestisce un nuovo valore inserito nella cbo nella relativa tabella
'
' @param inNomeTabella nome della tabella
' @param cbo comboBox
' @return
' @remarks aggiunge il nuovo valore alla cbo e posiziona il listIndex
Public Sub GestisciNuovo(inNomeTabella As String, ByRef inCbo As ComboBox)
    Dim rsDataset As Recordset
    Dim strSelezione As String
    Dim inKeyId As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    
    strSelezione = Replace(inCbo.Text, Chr(39), Chr(96))
    v_Nomi = Array("KEY", "NOME")
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & inNomeTabella & " WHERE NOME='" & UCase(Apostrophe(strSelezione)) & "'", cnPrinc, adOpenDynamic, adLockPessimistic, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        inKeyId = GetNumero(inNomeTabella)
        v_Val = Array(inKeyId, UCase(strSelezione))
        rsDataset.AddNew v_Nomi, v_Val

        inCbo.AddItem UCase(strSelezione)
        inCbo.ItemData(inCbo.NewIndex) = inKeyId
    End If
    Set rsDataset = Nothing
    
    Dim k As Integer
    For k = 0 To inCbo.ListCount - 1
        If UCase(inCbo.List(k)) = UCase(strSelezione) Then
            inCbo.ListIndex = k
        End If
    Next k
End Sub

''
' Gestisce un nuovo valore inserito nella cbo nella tabella Apparato
'
' @param inNomeTabella nome della tabella
' @param cbo comboBox
' @return
' @remarks aggiunge il nuovo valore alla cbo e posiziona il listIndex
Public Sub GestisciNuovoApparato(inNomeTabella As String, ByRef inCbo As ComboBox)
    Dim rsDataset As Recordset
    Dim strSelezione As String
    Dim inKeyId As Integer
    Dim v_Nomi() As Variant
    Dim v_Val() As Variant
    
    strSelezione = Replace(inCbo.Text, Chr(39), Chr(96))
   ' strSelezione = inCbo.Text
    v_Nomi = Array("KEY", "NOME")
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & inNomeTabella & " WHERE NOME='" & (Apostrophe(strSelezione)) & "'", cnPrinc, adOpenDynamic, adLockPessimistic, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        inKeyId = GetNumero(inNomeTabella)
        v_Val = Array(inKeyId, (strSelezione))
        rsDataset.AddNew v_Nomi, v_Val

        inCbo.AddItem (strSelezione)
        inCbo.ItemData(inCbo.NewIndex) = inKeyId
    End If
    Set rsDataset = Nothing
    
    Dim k As Integer
    For k = 0 To inCbo.ListCount - 1
        If UCase(inCbo.List(k)) = (strSelezione) Then
            inCbo.ListIndex = k
        End If
    Next k
End Sub

''
' Scrive positivo o  negativo nella flx
'
' @param flx griglia
' @param col colonna interessata
' @param nomeEsteso determina se deve scrivere POS o POSITIVO
' @return
' @remarks
Public Sub GestisciPN(flx As MSFlexGrid, Col As Integer, nomeEsteso As Boolean)
    Dim testo As String
    testo = flx.TextMatrix(flx.Row, Col)
    flx.CellForeColor = vbRed
    If testo = "" Then
         flx.TextMatrix(flx.Row, Col) = IIf(nomeEsteso, "NEGATIVO", "NEG")
    ElseIf testo = IIf(nomeEsteso, "NEGATIVO", "POS") Then
             flx.TextMatrix(flx.Row, Col) = IIf(nomeEsteso, "POSITIVO", "POS")
    Else
         flx.TextMatrix(flx.Row, Col) = ""
    End If

'    If testo = "" Then
'         flx.TextMatrix(flx.Row, Col) = IIf(nomeEsteso, "POSITIVO", "POS")
'    ElseIf testo = IIf(nomeEsteso, "POSITIVO", "POS") Then
'         flx.TextMatrix(flx.Row, Col) = IIf(nomeEsteso, "NEGATIVO", "NEG")
'    Else
'         flx.TextMatrix(flx.Row, Col) = ""
 '   End If
End Sub

''
' Restituisce il listIndex della cbo al nome corrispondente dato il suo key nella tabella analizzando il campo itemdata
'
' @param inKey key del nome nella tabella
' @param inCbo comboBox
' @return posizione del nome nella cbo
' @remarks restituire -1 se non lo trova
Public Function GetCboListIndex(inKey As Integer, ByRef inCbo As ComboBox) As Integer
    Dim i As Integer
    
    For i = 0 To inCbo.ListCount - 1
        If inCbo.ItemData(i) = inKey Then
            GetCboListIndex = i
            Exit Function
        End If
    Next i
    
    GetCboListIndex = -1
End Function

'' Ritorna le posizioni centrali del form
Public Sub GetCenterForm(inMeHeight As Single, inMeWidth As Single, ByRef outTop As Single, ByRef outLeft As Single)
    If (frmMain.Height - 3000 - inMeHeight) / 2 < 0 Then
        outTop = 0
    Else
        outTop = (frmMain.Height - 3000 - inMeHeight) / 2
    End If
    outLeft = (frmMain.Width - 300 - inMeWidth) / 2
End Sub

''
' Restituisce l'indice all'interno di una cbo
'
' @param cbo comboBox
' @param nome nome da cercare nella cbo
' @return posizione del nome nella cbo
' @remarks restituire -1 se non lo trova
Public Function GetIndex(cbo As ComboBox, nome As String) As Integer
    Dim i As Integer
    GetIndex = -1
    For i = 0 To cbo.ListCount - 1
        If UCase(cbo.List(i)) = UCase(nome) Then
            GetIndex = i
            Exit For
        End If
    Next i
End Function

''
' Restituisce il NOME di un record di un tabella conoscendo il key
'
' @param key key da cercare
' @param nomeTabella nome della tabella
' @return valore del campo NOME
' @remarks questa funzione si puo eliminare facendo i join dove è usata
Public Function GetNome(key As Integer, nomeTabella As String) As String
    Dim rsDataset As Recordset
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT * FROM " & nomeTabella & " WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.BOF And rsDataset.EOF Then
        GetNome = ""
    Else
        GetNome = rsDataset("NOME")
    End If
    Set rsDataset = Nothing
End Function

''
' Ritorna il primo key libero nella tabella
'
' @param inNomeTabella nome della tabella
' @return il primo key libero nella tabella
' @remarks utilizza metodi per velocizzare la ricerca, puo restituire numeri di record cancellati
Public Function GetNumero(inNomeTabella As String) As Long
    Dim rsDataset As Recordset
    Dim blnTrovato As Boolean
    Dim intNumero As Long
    
    Set rsDataset = New Recordset
    rsDataset.Open "SELECT COUNT(KEY) AS TOTALE FROM " & inNomeTabella, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset("TOTALE") > 500 And rsDataset("TOTALE") < 1000 Then
        intNumero = 500
    ElseIf rsDataset("TOTALE") > 1000 Then
        intNumero = rsDataset("TOTALE")
    Else
        intNumero = 0
    End If
    rsDataset.Close
    
    rsDataset.Open "SELECT KEY FROM " & inNomeTabella & " WHERE KEY>" & intNumero, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do
        intNumero = intNumero + 1
        rsDataset.Filter = "KEY=" & intNumero
        If Not (rsDataset.BOF And rsDataset.EOF) Then
            blnTrovato = True
        ElseIf rsDataset.BOF And rsDataset.EOF Then
            blnTrovato = False
        End If
    Loop Until blnTrovato = False
    
    GetNumero = intNumero

    Set rsDataset = Nothing
End Function

''
' Restituisce il key della tabella associato al campo nome da cercare
'
' @param inNomeTabella nome della tabella
' @param inNomeCampo nome del campo da cercare
' @param inNomeDaCercare valore da cercare
' @return
' @remarks
Public Function GetNumeroDaNome(inNomeTabella As String, inNomeCampo As String, inNomeDaCercare As String) As Integer
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT * FROM " & inNomeTabella & " WHERE " & inNomeCampo & "='" & Apostrophe(inNomeDaCercare) & "'", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsDataset.EOF And rsDataset.BOF Then
        GetNumeroDaNome = -1
    Else
        GetNumeroDaNome = rsDataset("KEY")
    End If
    rsDataset.Close
        
    Set rsDataset = Nothing
End Function

''
' Ritorna il key libero dopo il massimo key
'
' @param nomeTabella nome della tabella
' @param
' @return key libero dopo il massimo
' @remarks evita che siano restituiti key gia utilizzati in passato
Public Function GetNumeroNuovo(nomeTabella As String) As Integer
    Dim rsDataset As Recordset

    Set rsDataset = New Recordset
    rsDataset.Open "SELECT MAX(KEY) AS MASSIMO FROM " & nomeTabella, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not IsNull(rsDataset("MASSIMO")) Then
        GetNumeroNuovo = rsDataset("MASSIMO") + 1
    Else
        GetNumeroNuovo = 1
    End If
    rsDataset.Close
    Set rsDataset = Nothing
End Function

'' Restituisce l'ultimo gg di un mese di un anno
'
' @param mese mese da analizzare
' @param anno anno da analizzare
' @return gg/mm/aaaa
Public Function GetUltimoGiorno(mese As Integer, anno As Integer, Optional americano = False) As Date
    Dim IsAnnoBisestile As Boolean
    Dim ggFebbraio As Integer
    
    IsAnnoBisestile = IsDate(Format("2/29/" & anno, "mm/dd/yyyy"))
    If IsAnnoBisestile Then
        ggFebbraio = 29
    Else
        ggFebbraio = 28
    End If
    
    If Not americano Then
        Select Case mese
            Case 2
                GetUltimoGiorno = DateValue(ggFebbraio & "/" & mese & "/" & anno)
            Case 1, 3, 5, 7, 8, 10, 12
                GetUltimoGiorno = DateValue("31/" & mese & "/" & anno)
            Case Else
                GetUltimoGiorno = DateValue("30/" & mese & "/" & anno)
        End Select
    Else
        Select Case mese
            Case 2
                GetUltimoGiorno = DateValue(mese & "/" & ggFebbraio & "/" & anno)
            Case 1, 3, 5, 7, 8, 10, 12
                GetUltimoGiorno = DateValue(mese & "/31/" & anno)
            Case Else
                GetUltimoGiorno = DateValue(mese & "/30/" & anno)
        End Select
    End If
End Function

'' Restituisce i dati dell'utente
'
' @param key key dell'utente
' @return Cognome Nome dell'utente
Public Function GetUtente(key As Integer) As String
    Dim rsDataset As New Recordset
    
    rsDataset.Open "SELECT COGNOME, NOME FROM LOGIN WHERE KEY=" & key, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        GetUtente = rsDataset("COGNOME") & " " & rsDataset("NOME")
    Else
        GetUtente = ""
    End If
    rsDataset.Close
End Function

''
' Riproduce il tasto tab se si preme invio
'
' @param keyA tasto premuto
' @param
' @return
' @remarks utilizzato nelle txt di sola lettura
Public Sub InvioTab(ByRef keyA As Integer)
    If keyA = vbKeyReturn Then
        keyA = 0 'elimina il beep
        keybd_event VK_TAB, 0, 0, 0 'riproduce il TAB
    End If
End Sub

'' Controlla se ci sono record in questa tabella, altrimenti da il via libera per l'eliminazione del record
Public Function IsPossibleDelete(inNomeTabella As String, inNomeCampo As String, inKey As Integer) As Boolean
    On Error GoTo gestione
    Dim rsDataset As New Recordset
    rsDataset.Open "Select " & inNomeCampo & " From " & inNomeTabella & " Where " & inNomeCampo & "=" & inKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        IsPossibleDelete = False
    Else
        IsPossibleDelete = True
    End If
    rsDataset.Close
    Set rsDataset = Nothing
    
    Exit Function
gestione:
    IsPossibleDelete = False
End Function

''
' Verifica se ci sono clien connessi
'
' @param numClient numero di client connessi
' @param
' @return true se ci sono client, altrimenti false
' @remarks
Public Function nessunClient(ByRef numClient As Integer) As Boolean
    Dim rsDataset As New Recordset
    numClient = 0
    rsDataset.Open "CLIENT", cnPrinc, adOpenForwardOnly, adLockPessimistic, adCmdTable
    numClient = rsDataset("NUMERO")
    If numClient Then
        nessunClient = False
    Else
        nessunClient = True
    End If
    Set rsDataset = Nothing
End Function

''
' Verifica se nel db ci sono tutte le tabelle memorizzate nel file tabelle.xml
'
' @param
' @param
' @return
' @remarks
Public Function nonCorrotto() As Boolean
    On Error GoTo gestione
    Dim tabelle() As String
    Dim obj As DOMDocument
    Dim nome As IXMLDOMNodeList
    Dim elemento As IXMLDOMElement
    Dim nodo As IXMLDOMNode

    Set obj = New DOMDocument
    obj.async = False
    obj.Load structApri.pathExe & "\tabelle.xml"

    Set elemento = obj.documentElement
    Set nome = elemento.selectNodes("tabella/nome")

    ReDim tabelle(0)
    For Each nodo In nome
        ReDim Preserve tabelle(UBound(tabelle) + 1)
        tabelle(UBound(tabelle)) = nodo.Text
    Next

    Set nodo = Nothing
    Set nome = Nothing
    Set elemento = Nothing
    Set obj = Nothing

    On Error GoTo gestione
    Dim rsDataset As New Recordset
    Dim i As Integer
    
    For i = 1 To UBound(tabelle) - 1
        rsDataset.Open tabelle(i), cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
        rsDataset.Close
    Next i

    isCorrotto = False
    nonCorrotto = True
    Exit Function
gestione:
    ' identificativo errore 6-
    If Err.Number = 91 Then
        MsgBox "Errore n° 6-" & Err.Number & ":  " & vbCrLf & "File per il controllo di coerenza non trovato", vbCritical, "Attenzione"
    Else
        MsgBox "Errore n° 6-" & Err.Number & ":  " & vbCrLf & Err.Description, vbCritical, "Attenzione"
    End If
    End
End Function

''
' Permette di inserire solo numeri nel obj
'
' @param obj oggetto in cui inserire il valore (txt)
' @param lettera valore ultimo inserito nell'obj
' @return
' @remarks
Public Sub OnlyNumber(obj As Object, lettera As String)
    On Error GoTo gestione
    Static secondaVolta As Boolean
    If Not IsNumeric(lettera) Then
        If Not secondaVolta Or lettera = "," Then
            secondaVolta = Not secondaVolta
            If Len(obj.Text) <= 0 Then
                obj.Text = ""
            Else
                Dim Posizione As Byte
                Posizione = InStr(obj.Text, lettera)
                obj.Text = IIf(Posizione - 1 > 0, Left(obj.Text, Posizione - 1), vbNullString) & IIf(Len(obj.Text) - Posizione > 0, Right(obj.Text, Len(obj.Text) - Posizione), vbNullString)
            End If
        End If
    End If
    secondaVolta = False
    ' preme end per portare il cursore alla fine
    keybd_event vbKeyEnd, 0, KEYEVENTF_EXTENDEDKEY, 0       'simula la pressione del tasto end
    keybd_event vbKeyEnd, 0, KEYEVENTF_KEYUP, 0             'simula il rilascio del tasto end
    Exit Sub
gestione:
    If Err.Number = 5 Then
        Exit Sub
    Else
        MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Attenzione"
    End If
End Sub

''
' Determina la posizione del cursore
'
' @param PuntoX ascissa
' @param PuntoY ordinata
' @param frm form rispetto al quale calcolare la posizione
' @return
' @remarks
Public Sub PosizioneCursore(ByRef PuntoX As Integer, ByRef PuntoY As Integer, Optional frm As Long = 0)
    Dim Posizione As POINTAPI
    GetCursorPos Posizione
    If frm <> 0 Then
        ScreenToClient frm, Posizione 'converte la posizione x,y relativamente alla finestra specificata (frm)
    End If
    'per ricavare la posizione x e y
    PuntoX = Posizione.X * Screen.TwipsPerPixelX 'coordinata del punto x
    PuntoY = Posizione.Y * Screen.TwipsPerPixelY 'coordinata del punto y
End Sub

''
' Pulisce tutte le txt, lst e cbo
'
' @param frm form da pulire
' @param
' @return
' @remarks evita le flx
Public Sub PulisciForm(frm As Form, Optional inPulisciIntestazionePaziente As Boolean = True)
    Dim obj As Object
    On Error Resume Next
    For Each obj In frm.Controls
        If Mid(obj.Name, 1, 3) <> "flx" Then
            If inPulisciIntestazionePaziente And (obj.Name = "lblCognome" Or obj.Name = "lblNome" Or obj.Name = "lblEta") Then
                obj.Caption = ""
            Else
                obj.Text = ""
                obj.ListIndex = -1
            End If
        End If
    Next
End Sub

''
' Azzera il campo NUMERO nella tab clienti
'
' @param
' @param
' @return
' @remarks
Public Sub PulisciTabCLIENTI()
    Dim rsDataset As New Recordset
    rsDataset.Open "CLIENT", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        rsDataset.Update "NUMERO", 0
    Else
        rsDataset.AddNew
        rsDataset("NUMERO") = 0
        rsDataset.Update
    End If
    Set rsDataset = Nothing
End Sub

''
' Ricarica la cbo con la tab selezionata
'
' @param inQuery query da utilizzare
' @param inNomeCampo nome del campo da mostrare
' @remarks
Public Sub RicaricaComboBox(inQuery As String, inNomeCampo As String, ByRef inCbo As ComboBox)
    Dim rsDataset As New Recordset
    Dim strSelezione As String
    
   strSelezione = inCbo.Text
    inCbo.Clear
    rsDataset.Open inQuery, cnPrinc, adOpenForwardOnly, adLockReadOnly
    Do While Not rsDataset.EOF
        inCbo.AddItem rsDataset(inNomeCampo)
        inCbo.ItemData(inCbo.NewIndex) = rsDataset("KEY")
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    inCbo.ListIndex = GetIndex(inCbo, strSelezione)
    
    Set rsDataset = Nothing
End Sub

'' Ridispone i form nell form padre
Public Function RidisponiForms(ByRef inForm As Form) As Boolean
    Dim intTop As Single
    Dim intLeft As Single
    Dim i As Integer
    
    RidisponiForms = True
    If Not RIDISPONI_FORMS Then
        Exit Function
    End If
    
    If inForm.Tag <> "" Then
        inForm.Tag = ""
        RidisponiForms = False
        Exit Function
    End If
        
    For i = 0 To Forms.count - 1
        If Not TypeOf Forms(i) Is MDIForm Then
            If Forms(i).Caption <> inForm.Caption Then
                Forms(i).Top = intTop
                Forms(i).Left = intLeft
                Forms(i).ZOrder
                Forms(i).Tag = "SistemazioneForms"
                intTop = intTop + 260
                intLeft = intLeft + 20
            End If
        End If
    Next

    Call GetCenterForm(inForm.Height, inForm.Width, intTop, intLeft)
    inForm.Top = intTop
    inForm.Left = intLeft
    inForm.Tag = "SistemazioneForms"
    inForm.ZOrder
End Function

''
' Allarga la cbo
'
' @param cbo comboBox
' @param lWidth larghezza richiesta
' @return
' @remarks
Public Sub SetComboWidth(cbo As ComboBox, lWidth As Long)
    SendMessage cbo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0
End Sub

'' Stampa diario clinico
Public Sub StampaDecimaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim rsDiario As Recordset
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    strShape = "SHAPE APPEND " & _
                "   NEW adVarChar(50) AS TITOLO, " & _
                "   NEW adDate AS DATA, " & _
                "   NEW adLongVarChar AS DATI "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
        
    strSql = "SELECT    DATA, TITOLI_DIARIO.NOME as TITOLI_DIARIONOME, DATI, UTENTE_MODIFICATORE " & _
             "FROM      (DIARI_CLINICI " & _
             "          INNER JOIN TITOLI_DIARIO ON TITOLI_DIARIO.KEY=DIARI_CLINICI.CODICE_TITOLO) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          STAMPA=TRUE " & _
             "ORDER BY  CODICE_TITOLO, DATA"
    
    ' carica il recordset padre
    Set rsDiario = New Recordset
    rsDiario.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDiario.EOF And rsDiario.BOF) Then
        With rsMain
            Do While Not rsDiario.EOF
                .AddNew
                .Fields("TITOLO") = rsDiario("TITOLI_DIARIONOME")
                .Fields("DATA") = rsDiario("DATA")
                .Fields("DATI") = rsDiario("DATI") & vbCrLf & vbCrLf & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsDiario("UTENTE_MODIFICATORE"))
                rsDiario.MoveNext
            Loop
        End With
    End If
    rsDiario.Close
    
    Set rptCartellaClinica_10 = Nothing
    Set rptCartellaClinica_10.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_10.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_10.PrintReport Not formPazienti, rptRangeAllPages
End Sub

'' Stampa accessi vascolari
Public Sub StampaNonaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim SQLString As String
    Dim rsAccessi As Recordset
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    SQLString = "SHAPE APPEND " & _
                "   NEW adDate AS DATA, " & _
                "   NEW adLongVarChar AS INTERVENTO, " & _
                "   NEW adDate AS DATA_CHIUSURA_ACCESSO , " & _
                "   NEW adLongVarChar AS CAUSA_CHIUSURA_ACCESSO, " & _
                "   NEW adVarChar (100) AS OPERATORE1, " & _
                "   NEW adVarChar (100) AS OPERATORE2, " & _
                "   NEW adInteger AS ANESTESIA, " & _
                "   NEW adLongVarChar AS DATI"
                
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
        
    ' carica il recordset padre
    Set rsAccessi = New Recordset
    rsAccessi.Open "SELECT * FROM ACCESSI_VASCOLARI_TAB WHERE (CODICE_PAZIENTE=" & codicePaziente & ") ORDER BY DATA", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsAccessi.EOF And rsAccessi.BOF) Then
        With rsMain
            Do While Not rsAccessi.EOF
                .AddNew
                .Fields("DATA") = rsAccessi("DATA")
                .Fields("INTERVENTO") = rsAccessi("INTERVENTO")
                .Fields("DATA_CHIUSURA_ACCESSO") = rsAccessi("DATA_CHIUSURA_ACCESSO")
                .Fields("CAUSA_CHIUSURA_ACCESSO") = rsAccessi("CAUSA_CHIUSURA_ACCESSO")
                .Fields("OPERATORE1") = DatiMedico(rsAccessi("CODICE_MEDICO1"))
                .Fields("OPERATORE2") = DatiMedico(rsAccessi("CODICE_MEDICO2"))
                .Fields("ANESTESIA") = rsAccessi("ANESTESIA")
                .Fields("DATI") = rsAccessi("DATI")
                rsAccessi.MoveNext
            Loop
        End With
    End If
    
    Set rptCartellaClinica_9 = Nothing
    Set rptCartellaClinica_9.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_9.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_9.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_9.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_9.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_9.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_9.PrintReport Not formPazienti, rptRangeAllPages
End Sub

'' Stampa terapia domiciliare corrente
Public Sub StampaTerapiaDomiciliareCorrente(codicePaziente As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 20
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    strSql = "SELECT    TERAPIE_DOMICILIARI.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DOMICILIARI " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DOMICILIARI.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=FALSE " & _
             "ORDER BY  DATA DESC"
    
    Set rsTerapia = New Recordset
    
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Domiciliare"
            .Fields("INTESTAZIONE_SOSPESA") = ""
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = ""
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Domiciliare"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = ""
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
        End With
    End If
    rsTerapia.Close

    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_8 = Nothing
        Set rptCartellaClinica_8.DataSource = rsMain
        rptCartellaClinica_8.LeftMargin = 0
        rptCartellaClinica_8.RightMargin = 0
        rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        rptCartellaClinica_8.PrintReport True, rptRangeAllPages
    End If
End Sub

'' Stampa terapia domiciliare sospesa
Public Sub StampaTerapiaDomiciliareSospesa(codicePaziente As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 20
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    strSql = "SELECT    TERAPIE_DOMICILIARI.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DOMICILIARI " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DOMICILIARI.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=TRUE " & _
             "ORDER BY  DATA DESC"
             
    Set rsTerapia = New Recordset
             
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Domiciliare - Sospesa"
            .Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = rsTerapia("DATA_SOSPESA")
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Domiciliare - Sospesa"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
            rsTerapia.Close
        End With
    End If
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_8 = Nothing
        Set rptCartellaClinica_8.DataSource = rsMain
        rptCartellaClinica_8.LeftMargin = 0
        rptCartellaClinica_8.RightMargin = 0
        rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        rptCartellaClinica_8.PrintReport True, rptRangeAllPages
    End If
End Sub

'' Stampa terapia domiciliare
Public Sub StampaOttavaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 20
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    ' terapie non sospese
    strSql = "SELECT    TERAPIE_DOMICILIARI.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DOMICILIARI " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DOMICILIARI.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=FALSE " & _
             "ORDER BY  DATA DESC"
    
    Set rsTerapia = New Recordset
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Domiciliare"
            .Fields("INTESTAZIONE_SOSPESA") = ""
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = ""
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Domiciliare"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = ""
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
        End With
    End If
    rsTerapia.Close

    k = k + 1
    ' terapie sospese
    strSql = "SELECT    TERAPIE_DOMICILIARI.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DOMICILIARI " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DOMICILIARI.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=TRUE " & _
             "ORDER BY  DATA DESC"
             
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Domiciliare - Sospesa"
            .Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = rsTerapia("DATA_SOSPESA")
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Domiciliare - Sospesa"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
            rsTerapia.Close
        End With
    End If
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_8 = Nothing
        Set rptCartellaClinica_8.DataSource = rsMain
        rptCartellaClinica_8.LeftMargin = 0
        rptCartellaClinica_8.RightMargin = 0
        If Not formPazienti Then
            rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
            rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        Else
            rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
            rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
            rptCartellaClinica_8.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
        End If
        rptCartellaClinica_8.PrintReport Not formPazienti, rptRangeAllPages
    End If
End Sub

'' Stampa anamnesi dialitica
Public Sub StampaQuartaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strSqlStampa As String
    Dim strSql As String
    Dim i As Integer
    Dim valore As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsTabelle As Recordset           ' tabelle per la stampa (nefrologia, patologica, ..
    Dim rsAppo As Recordset
        
    strSqlStampa = "        NEW adBoolean as DIURESI, " & _
                "       NEW adSingle as AUMENTO, " & _
                "       NEW adSingle as PESO_SECCO, " & _
                "       NEW adSingle as QUANTITA, " & _
                "       NEW adDate AS DATA_PESO_SECCO, " & _
                "       NEW adVarChar (50) as FILTRO, " & _
                "       NEW adDate AS DATA_FILTRO, " & _
                "       NEW adVarChar (50) as LINEE, " & _
                "       NEW adDate AS DATA_LINEE, " & _
                "       NEW adVarChar (50) as ACCESSO_VASCOLARE, " & _
                "       NEW adVarChar (50) as SEDE_ACCESSO, " & _
                "       NEW adVarChar (50) as TIPO_DIALISI, " & _
                "       NEW adSingle as SODIO, " & _
                "       NEW adSingle as POTASSIO, " & _
                "       NEW adSingle as BICARBONATO, " & _
                "       NEW adSingle as CALCIO, " & _
                "       NEW adInteger as MINUTI, " & _
                "       NEW adInteger as ORE, " & _
                "       NEW adVarChar (50) as ANTICOAGULANTE1, " & _
                "       NEW adSingle as DOSE1, " & _
                "       NEW adSingle as DOSE2, " & _
                "       NEW adSingle as DOSE3, " & _
                "       NEW adVarChar (50) as ANTICOAGULANTE2, " & _
                "       NEW adSingle as DOSE4, " & _
                "       NEW adSingle as FLUSSO,       NEW adSingle as FLUSSO_SANGUE, "
    strSqlStampa = strSqlStampa & _
                "       NEW adVarChar (50) as SOL_DIALITICA, " & _
                "       NEW adVarChar (50) as SOL_INFUSIONALE, " & _
                "       NEW adSingle as VALORE_CC, " & _
                "       NEW adVarChar (50) as CARTUCCIA, " & _
                "       NEW adVarChar (3) as SEDUTE, " & _
                "       NEW adVarChar (10) as EPO, " & _
                "       NEW adVarChar (10) as UI, " & _
                "       NEW adLongVarChar as NOTE, " & _
                "       NEW adVarChar (70) as ESAME1, " & _
                "       NEW adVarChar (70) as ESAME2, " & _
                "       NEW adVarChar (70) as ESAME3, " & _
                "       NEW adVarChar (50) as AGO1, " & _
                "       NEW adVarChar (50) as AGO2, " & _
                "       NEW adVarChar (5) as DOSI_UNITA_MISURA, " & _
                "       NEW adVarChar (5) as UNITA_VAL_SOL_INF, " & _
                "       NEW adSingle as CODICE_PRESTAZIONE, " & _
                "       NEW adSingle as GLUCOSIO "

    ' stringa di shape
    strSqlStampa = "SHAPE APPEND " & strSqlStampa
     
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    ' carica il recordset padre
    Set rsTabelle = New Recordset
    strSql = "SELECT    ANAMNESI_DIALITICHE.*, FILTRI.NOME AS FILTRINOME, LINEE.NOME AS LINEENOME, ACCESSI_VASCOLARI.NOME AS ACCESSI_VASCOLARINOME, " & _
            "           AGO1.NOME AS AGO1NOME, AGO2.NOME AS AGO2NOME, TIPI_DIALISI.NOME AS TIPI_DIALISINOME, ANTICOAGULANTI1.NOME AS ANTICOAGULANTI1NOME, " & _
            "           ANTICOAGULANTI2.NOME AS ANTICOAGULANTI2NOME, SOL_DIALITICHE.NOME AS SOL_DIALITICHENOME, SOL_INFUSIONALI.NOME AS SOL_INFUSIONALINOME, " & _
            "           CARTUCCE.NOME AS CARTUCCENOME " & _
            " FROM      (((((((((((ANAMNESI_DIALITICHE " & _
            "           LEFT OUTER JOIN FILTRI ON FILTRI.KEY=ANAMNESI_DIALITICHE.TIPO_FILTRO) " & _
            "           LEFT OUTER JOIN LINEE ON LINEE.KEY=ANAMNESI_DIALITICHE.TIPO_LINEE) " & _
            "           LEFT OUTER JOIN ACCESSI_VASCOLARI ON ACCESSI_VASCOLARI.KEY=ANAMNESI_DIALITICHE.ACCESSO_VASCOLARE) " & _
            "           LEFT OUTER JOIN AGO AGO1 ON AGO1.KEY=ANAMNESI_DIALITICHE.AGO1) " & _
            "           LEFT OUTER JOIN AGO AGO2 ON AGO2.KEY=ANAMNESI_DIALITICHE.AGO2) " & _
            "           LEFT OUTER JOIN TIPI_DIALISI ON TIPI_DIALISI.KEY=ANAMNESI_DIALITICHE.TIPO_DIALISI) " & _
            "           LEFT OUTER JOIN ANTICOAGULANTI ANTICOAGULANTI1 ON ANTICOAGULANTI1.KEY=ANAMNESI_DIALITICHE.ANTICOAGULANTE1) " & _
            "           LEFT OUTER JOIN ANTICOAGULANTI ANTICOAGULANTI2 ON ANTICOAGULANTI2.KEY=ANAMNESI_DIALITICHE.ANTICOAGULANTE2) " & _
            "           LEFT OUTER JOIN SOL_DIALITICHE ON SOL_DIALITICHE.KEY=ANAMNESI_DIALITICHE.SOL_DIALITICA) " & _
            "           LEFT OUTER JOIN SOL_INFUSIONALI ON SOL_INFUSIONALI.KEY=ANAMNESI_DIALITICHE.SOL_INFUSIONALE) " & _
            "           LEFT OUTER JOIN CARTUCCE ON CARTUCCE.KEY=ANAMNESI_DIALITICHE.CARTUCCIA) " & _
            " Where     CODICE_PAZIENTE = " & codicePaziente
    
    rsTabelle.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTabelle.EOF And rsTabelle.BOF) Then
        With rsMain
            .AddNew
            If rsTabelle("RITMO_DIALITICO") = 0 Then
                .Fields("SEDUTE") = "- -"
            Else
                .Fields("SEDUTE") = rsTabelle("RITMO_DIALITICO")
            End If
            .Fields("DIURESI") = CBool(rsTabelle("DIURESI"))
            .Fields("QUANTITA") = rsTabelle("QUANTITA")
            .Fields("AUMENTO") = rsTabelle("AUMENTO_POND")
            .Fields("PESO_SECCO") = rsTabelle("PESO_SECCO")
            .Fields("DATA_PESO_SECCO") = rsTabelle("DATA_PESO")
            .Fields("FILTRO") = rsTabelle("FILTRINOME")
            .Fields("DATA_FILTRO") = rsTabelle("DATA_FILTRO")
            .Fields("LINEE") = rsTabelle("LINEENOME")
            .Fields("DATA_LINEE") = rsTabelle("DATA_LINEE")
            .Fields("ACCESSO_VASCOLARE") = rsTabelle("ACCESSI_VASCOLARINOME")
            .Fields("AGO1") = rsTabelle("AGO1NOME")
            .Fields("AGO2") = rsTabelle("AGO2NOME")
            .Fields("SEDE_ACCESSO") = rsTabelle("SEDE_ACCESSO")
            .Fields("TIPO_DIALISI") = rsTabelle("TIPI_DIALISINOME")
            .Fields("CODICE_PRESTAZIONE") = rsTabelle("CODICE_PRESTAZIONE")
            .Fields("SODIO") = rsTabelle("SODIO")
            .Fields("POTASSIO") = rsTabelle("POTASSIO")
            .Fields("BICARBONATO") = rsTabelle("BICARBONATO")
            .Fields("CALCIO") = rsTabelle("CALCIO")
            .Fields("GLUCOSIO") = rsTabelle("GLUCOSIO")
            .Fields("MINUTI") = rsTabelle("MINUTI")
            .Fields("ORE") = rsTabelle("ORE")
            .Fields("ANTICOAGULANTE1") = rsTabelle("ANTICOAGULANTI1NOME")
            .Fields("ANTICOAGULANTE2") = rsTabelle("ANTICOAGULANTI2NOME")
            .Fields("DOSI_UNITA_MISURA") = rsTabelle("DOSI_UNITA_MISURA")
            .Fields("DOSE1") = rsTabelle("DOSE1")
            .Fields("DOSE2") = rsTabelle("DOSE2")
            .Fields("DOSE3") = rsTabelle("DOSE3")
            .Fields("DOSE4") = rsTabelle("DOSE4")
            .Fields("FLUSSO") = rsTabelle("FLUSSO")
            .Fields("FLUSSO_SANGUE") = rsTabelle("FLUSSO_SANGUE")
            .Fields("SOL_DIALITICA") = rsTabelle("SOL_DIALITICHENOME")
            .Fields("SOL_INFUSIONALE") = rsTabelle("SOL_INFUSIONALINOME")
            .Fields("VALORE_CC") = rsTabelle("SOL_INF_CC")
            .Fields("UNITA_VAL_SOL_INF") = rsTabelle("UNITA_VAL_SOL_INF")
            .Fields("CARTUCCIA") = rsTabelle("CARTUCCENOME")
            .Fields("EPO") = rsTabelle("EPO")
            If .Fields("EPO") = 2 Or .Fields("EPO") = 3 Then
                .Fields("UI") = rsTabelle("UI") & "  " & "mcg"
            Else
                .Fields("UI") = rsTabelle("UI") & "  " & "UI"
            End If
            .Fields("NOTE") = rsTabelle("NOTE")
            rsTabelle.Close
            
            i = 0
            Set rsAppo = New Recordset
            rsTabelle.Open "SELECT KEY, NOME FROM VOCI_ESAMI WHERE STAMPA=TRUE", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do While Not rsTabelle.EOF
                i = i + 1
                strSql = "SELECT        TOP 1 VALORE, DATA " & _
                         "FROM          (ANAMNESI_ESAMI " & _
                         "              INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                         "WHERE         CODICE_PAZIENTE=" & codicePaziente & " AND " & _
                         "              CODICE_ESAME=" & rsTabelle("KEY") & " " & _
                         "ORDER BY      DATA DESC"
                         
                rsAppo.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsAppo.EOF And rsAppo.BOF) Then
                    Select Case rsAppo("VALORE")
                        Case -2
                            valore = "NEGATIVO"
                        Case -1
                            valore = "POSITIVO"
                        Case Else
                            valore = rsAppo("VALORE")
                    End Select
                    .Fields("ESAME" & i) = rsTabelle("NOME") & vbCrLf & valore & vbCrLf & rsAppo("DATA")
                Else
                    .Fields("ESAME" & i) = rsTabelle("NOME") & vbCrLf & "Non definito"
                End If
                rsAppo.Close
                rsTabelle.MoveNext
            Loop
            rsTabelle.Close
            Set rsAppo = Nothing
            
            If i < 3 Then
                For i = i + 1 To 3
                    .Fields("ESAME" & i) = ""
                Next i
            End If
            
            '.Fields("SEDUTE") = 0
            
            'rsTabelle.Open "SELECT * FROM TURNI WHERE CODICE_PAZIENTE=" & codicePaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
            'If rsTabelle.EOF And rsTabelle.BOF Then
            '    .Fields("SEDUTE") = 0
            'Else
            '    For i = 1 To 7
            '        If rsTabelle("AM_INIZIO" & i) <> "" Then
            '            .Fields("SEDUTE") = .Fields("SEDUTE") + 1
            '        ElseIf rsTabelle("PM_INIZIO" & i) <> "" Then
            '            .Fields("SEDUTE") = .Fields("SEDUTE") + 1
            '        ElseIf rsTabelle("SR_INIZIO" & i) <> "" Then
            '            .Fields("SEDUTE") = .Fields("SEDUTE") + 1
            '        End If
            '    Next i
            'End If
            'rsTabelle.Close
            
        End With
    End If
    Set rsTabelle = Nothing

    Set rptCartellaClinica_4 = Nothing
    Set rptCartellaClinica_4.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_4.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_4.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_4.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_4.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_4.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_4.PrintReport Not formPazienti, rptRangeAllPages
End Sub

'' Stampa esami strumentali
Public Sub StampaQuintaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim rsEsami As Recordset
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    
    strShape = "SHAPE APPEND " & _
                "   NEW adVarChar(50) AS NOME_ORGANO, " & _
                "   NEW adVarChar(50) AS NOME_ESAME, " & _
                "   NEW adDate AS DATA, " & _
                "   NEW adLongVarChar AS REFERTO "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
        
    ' carica il recordset padre
    strSql = "SELECT    ORGANI.NOME AS ORGANINOME, ESAMI.NOME AS ESAMINOME, DATA, REFERTO, UTENTE_MODIFICATORE " & _
             "FROM      ((ESAMI_STRUMENTALI INNER JOIN ORGANI ON ORGANI.KEY=ESAMI_STRUMENTALI.CODICE_ORGANO) " & _
             "          INNER JOIN ESAMI ON ESAMI.KEY=ESAMI_STRUMENTALI.CODICE_ESAME) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          STAMPA=TRUE " & _
             "ORDER BY  ESAMI_STRUMENTALI.CODICE_ORGANO, CODICE_ESAME, DATA"
             
    Set rsEsami = New Recordset
    rsEsami.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsEsami.EOF And rsEsami.BOF) Then
        With rsMain
            Do While Not rsEsami.EOF
                .AddNew
                .Fields("NOME_ORGANO") = rsEsami("ORGANINOME")
                .Fields("NOME_ESAME") = rsEsami("ESAMINOME")
                .Fields("DATA") = rsEsami("DATA")
                .Fields("REFERTO") = rsEsami("REFERTO") & vbCrLf & vbCrLf & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsEsami("UTENTE_MODIFICATORE"))
                rsEsami.MoveNext
            Loop
        End With
    End If
    rsEsami.Close

    Set rptCartellaClinica_5 = Nothing
    Set rptCartellaClinica_5.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_5.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_5.PrintReport Not formPazienti, rptRangeAllPages
End Sub

'' Stampa anamnesi patologica
Public Sub StampaSecondaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim SQLString As String
    Dim v_voci() As Variant
    Dim i As Integer
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shapE
    Dim rsFiglio As Recordset
    Dim rsTabelle As Recordset           ' tabelle per la stampa  patologica
           
    ' stringa di shape
    SQLString = "SHAPE APPEND " & _
            "       NEW adVarChar(30) AS NOME_ANAMNESI, " & _
            "       NEW adInteger AS LINK1, " & _
            "       (( SHAPE APPEND " & _
            "           NEW adInteger AS LINK1, " & _
            "           NEW adLongVarChar AS TESTO " & _
            "       ) RELATE LINK1 TO LINK1 " & _
            "       ) AS Res1 "
     
    v_voci = Array("Malattie cardiovascolari", "Malattie polmonari", "Malattie tubo digerente", "Malattie endocrino - dismetaboliche", _
                 "Malattie nefro - uro - genitali", "Malattie infettive", "Malattie osteo articolari", "Interventi chirurgici", "Ricoveri", "Varie", "Generale")
    
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic

    Set rsTabelle = New Recordset
    rsTabelle.Open "SELECT * FROM ANAMNESI_PATOLOGICHE WHERE CODICE_PAZIENTE=" & codicePaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsMain
        If Not (rsTabelle.EOF And rsTabelle.BOF) Then
            ' familiare
            .AddNew
            .Fields("NOME_ANAMNESI") = "Anamnesi familiare"
            .Fields("LINK1") = 1
            Set rsFiglio = .Fields("Res1").Value
            rsFiglio.AddNew
            rsFiglio.Fields("LINK1") = 1
            If rsTabelle("ANAMNESI_FAMILIARE") = "" Then
                rsFiglio.Fields("TESTO") = "ANAMNESI FAMILIARE NON COMPILATA"
            Else
                rsFiglio.Fields("TESTO") = rsTabelle("ANAMNESI_FAMILIARE") & vbCrLf & vbCrLf & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsTabelle("UTENTE_MODIFICATORE_FAMILIARE"))
            End If
            rsFiglio.Update
            .Update
        
            ' patologica
            .AddNew
            .Fields("NOME_ANAMNESI") = "Anamnesi patologica remota"
            .Fields("LINK1") = 2
            Set rsFiglio = .Fields("Res1").Value
            If CBool(rsTabelle("STAMPA11")) = True And rsTabelle("SCHEDA_CLINICA11") <> "" Then
                rsFiglio.AddNew
                rsFiglio.Fields("LINK1") = 2
                rsFiglio.Fields("TESTO") = v_voci(10) & ":" & "       " & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsTabelle("UTENTE_MODIFICATORE11")) & vbCrLf & vbCrLf & rsTabelle("SCHEDA_CLINICA11")
                rsFiglio.Update
            End If
            For i = 1 To 10
                If CBool(rsTabelle("STAMPA" & i)) = True And rsTabelle("SCHEDA_CLINICA" & i) <> "" Then
                    rsFiglio.AddNew
                    rsFiglio.Fields("LINK1") = 2
                    rsFiglio.Fields("TESTO") = v_voci(i - 1) & ":" & "    " & "Ultimo aggiornamento del dr./dr.ssa: " & GetUtente(rsTabelle("UTENTE_MODIFICATORE" & i)) & vbCrLf & vbCrLf & rsTabelle("SCHEDA_CLINICA" & i)
                    rsFiglio.Update
                End If
            Next i
            If rsFiglio.RecordCount = 0 Then
                rsFiglio.AddNew
                rsFiglio.Fields("LINK1") = 2
                rsFiglio.Fields("TESTO") = "ANAMNESI PATOLOGICA NON COMPILATA"
                rsFiglio.Update
            End If
            .Update
        Else
            ' familiare
            .AddNew
            .Fields("NOME_ANAMNESI") = "Anamnesi familiare"
            .Fields("LINK1") = 1
            Set rsFiglio = .Fields("Res1").Value
            rsFiglio.AddNew
            rsFiglio.Fields("LINK1") = 1
            rsFiglio.Fields("TESTO") = "ANAMNESI FAMILIARE NON COMPILATA"
            rsFiglio.Update
            .Update
            ' patologica
            .AddNew
            .Fields("NOME_ANAMNESI") = "Anamnesi patologica remota"
            .Fields("LINK1") = 2
            Set rsFiglio = .Fields("Res1").Value
            rsFiglio.AddNew
            rsFiglio.Fields("LINK1") = 2
            rsFiglio.Fields("TESTO") = "ANAMNESI PATOLOGICA NON COMPILATA"
            rsFiglio.Update
            .Update
        End If
    End With
    rsTabelle.Close

    Set rptCartellaClinica_2 = Nothing
    Set rptCartellaClinica_2.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_2.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_2.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_2.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_2.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
        rptCartellaClinica_2.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
    End If
    rptCartellaClinica_2.PrintReport Not formPazienti, rptRangeAllPages
End Sub

'' Stampa esami di laboratorio
Public Sub StampaSestaParte(formPazienti As Boolean, codicePaziente As Integer, condizione As String, quantimesi As Integer, Optional codiceId As Integer)
    Dim valoreAppo As Variant
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim p As Integer
    Dim numPag As Integer
    Dim strShape As String
    Dim strSql As String
    Dim peso As Single
    Dim legenda As String
    Dim vett() As tabEsami
    
    Const pesoLivello1 As Single = 0.6
    Const pesoLivello2 As Single = 1.8
    Const pesoLivello3 As Single = 0.5
    Const pesoTotale As Single = 11
    Const numRigaGruppo As Integer = 1
    Const numRigaUtente As Integer = 0
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio1 As Recordset      ' recorset figli
    Dim rsFiglio2 As Recordset
    Dim rsEsami As Recordset
    
    strShape = " SHAPE APPEND " & _
                "   NEW adVarChar(7) AS ANNO, " & _
                "   (( SHAPE APPEND  NEW adVarChar(7) AS ANNO, " & _
                "       NEW adVarChar(50) AS NOME_GRUPPO, " & _
                "       NEW adVarChar(10) AS CODICE_GRUPPO, " & _
                "       NEW adVarChar(5) AS DATA1, " & _
                "       NEW adVarChar(5) AS DATA2, " & _
                "       NEW adVarChar(5) AS DATA3, " & _
                "       NEW adVarChar(5) AS DATA4, " & _
                "       NEW adVarChar(5) AS DATA5, " & _
                "       NEW adVarChar(5) AS DATA6, " & _
                "       NEW adVarChar(5) AS DATA7, " & _
                "       NEW adVarChar(5) AS DATA8, " & _
                "       NEW adVarChar(5) AS DATA9, " & _
                "       NEW adVarChar(5) AS DATA10, " & _
                "       NEW adVarChar(5) AS DATA11, " & _
                "       NEW adVarChar(5) AS DATA12, "
    strShape = strShape & _
                "       NEW adVarChar(5) AS UTENTE1, " & _
                "       NEW adVarChar(5) AS UTENTE2, " & _
                "       NEW adVarChar(5) AS UTENTE3, " & _
                "       NEW adVarChar(5) AS UTENTE4, " & _
                "       NEW adVarChar(5) AS UTENTE5, " & _
                "       NEW adVarChar(5) AS UTENTE6, " & _
                "       NEW adVarChar(5) AS UTENTE7, " & _
                "       NEW adVarChar(5) AS UTENTE8, " & _
                "       NEW adVarChar(5) AS UTENTE9, " & _
                "       NEW adVarChar(5) AS UTENTE10, " & _
                "       NEW adVarChar(5) AS UTENTE11, " & _
                "       NEW adVarChar(5) AS UTENTE12, "
    strShape = strShape & _
                "       (( SHAPE APPEND " & _
                "           NEW adVarChar(10) AS CODICE_GRUPPO, " & _
                "           NEW adVarChar(50)  AS NOME_ESAME, " & _
                "           NEW adInteger AS CODICE_ESAME, " & _
                "           NEW adVarChar (8) AS VALORE1, " & _
                "           NEW adVarChar (8) AS VALORE2, " & _
                "           NEW adVarChar (8) AS VALORE3, " & _
                "           NEW adVarChar (8) AS VALORE4, " & _
                "           NEW adVarChar (8) AS VALORE5, " & _
                "           NEW adVarChar (8) AS VALORE6, " & _
                "           NEW adVarChar (8) AS VALORE7, " & _
                "           NEW adVarChar (8) AS VALORE8, " & _
                "           NEW adVarChar (8) AS VALORE9, " & _
                "           NEW adVarChar (8) AS VALORE10, " & _
                "           NEW adVarChar (8) AS VALORE11, " & _
                "           NEW adVarChar (8) AS VALORE12, " & _
                "           NEW adLongVarChar AS LEGENDA, " & _
                "           NEW adVarChar (12) AS UNITA, " & _
                "           NEW adVarChar (16) AS V_MINMAX " & _
                "       ) RELATE CODICE_GRUPPO TO CODICE_GRUPPO ) AS RES2 " & _
                "   ) RELATE ANNO TO ANNO ) AS RES1"

    
    legenda = "Legenda:"
    Set rsEsami = New Recordset
    strSql = "SELECT  DISTINCT LOGIN.KEY, COGNOME, NOME " & _
            "FROM   (ANAMNESI_ESAMI " & _
            "       INNER JOIN LOGIN ON LOGIN.KEY=ANAMNESI_ESAMI.UTENTE_MODIFICATORE) " & _
            "WHERE  CODICE_PAZIENTE=" & codicePaziente & condizione
    rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not rsEsami.EOF
        legenda = legenda & vbCrLf & Mid(rsEsami("COGNOME"), 1, 2) & " " & Mid(rsEsami("NOME"), 1, 2) & " = " & rsEsami("COGNOME") & " " & rsEsami("NOME")
        rsEsami.MoveNext
    Loop
        If rsEsami.RecordCount = 0 Then
            MsgBox "Non sono presenti esami da stampare", vbInformation, "Attenzione"
            Exit Sub
        End If
    rsEsami.Close
    
    ReDim vett(0)
    ' carica tutti gli anni
    strSql = "SELECT    DISTINCT YEAR([DATA]) AS ANNO " & _
            "FROM       ANAMNESI_ESAMI " & _
            "WHERE      CODICE_PAZIENTE=" & codicePaziente & condizione
    rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
    rsEsami.MoveLast
    Do While Not rsEsami.BOF
        ReDim Preserve vett(UBound(vett) + 1)
        vett(UBound(vett)).anno = rsEsami("ANNO")
        ReDim vett(UBound(vett)).esami(0)
        rsEsami.MovePrevious
    Loop
    rsEsami.Close
    
    ' carica tutti i gruppi
    For i = 1 To UBound(vett)
        strSql = "SELECT    DISTINCT CODICE_GRUPPO, NOME " & _
                 "FROM      (ANAMNESI_ESAMI " & _
                 "          INNER JOIN GRUPPI_ESAMI ON GRUPPI_ESAMI.KEY=ANAMNESI_ESAMI.CODICE_GRUPPO) " & _
                 "WHERE     (CODICE_PAZIENTE=" & codicePaziente & _
                 condizione & _
                 "          AND YEAR([DATA])=" & vett(i).anno & ") "
        rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
        Do While Not rsEsami.EOF
            ReDim Preserve vett(i).esami(UBound(vett(i).esami) + 1)
            ReDim vett(i).esami(UBound(vett(i).esami)).righe(numRigaGruppo)
            vett(i).esami(UBound(vett(i).esami)).righe(numRigaGruppo).codice = rsEsami("CODICE_GRUPPO")
            vett(i).esami(UBound(vett(i).esami)).righe(numRigaGruppo).nome = rsEsami("NOME")
            rsEsami.MoveNext
        Loop
        rsEsami.Close
    Next
    
    ' carica le date
    For i = 1 To UBound(vett)
        For k = 1 To UBound(vett(i).esami)
            strSql = "SELECT      DATA " & _
                    "FROM       ANAMNESI_ESAMI " & _
                    "WHERE      (CODICE_PAZIENTE=" & codicePaziente & _
                    condizione & _
                    "           AND YEAR([DATA])=" & vett(i).anno & " AND " & _
                    "           CODICE_GRUPPO=" & vett(i).esami(k).righe(numRigaGruppo).codice & " ) " & _
                    "ORDER BY   DATA DESC"
            rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
       ' controlla e limita la stampa ad una sola data
            If quantimesi = 12 Then
                For j = 1 To IIf(rsEsami.RecordCount > 12, 12, rsEsami.RecordCount)
                    vett(i).esami(k).righe(numRigaGruppo).valori(j) = Format(rsEsami("DATA"), "dd/mm")
                    rsEsami.MoveNext
                Next j
            Else
                vett(i).esami(k).righe(numRigaGruppo).valori(1) = Format(rsEsami("DATA"), "dd/mm")
            End If
            rsEsami.Close
        Next k
    Next i
    
    ' carica gli esami
    For i = 1 To UBound(vett)
        For k = 1 To UBound(vett(i).esami)
            strSql = "SELECT     CODICE_ESAME, VOCI_ESAMI.NOME as VOCI_ESAMINOME, UNITA, MAX, MIN, PN " & _
                     "FROM       (ASSOCIAZIONE_ESAMI_LAB " & _
                     "           INNER JOIN VOCI_ESAMI ON VOCI_ESAMI.KEY=ASSOCIAZIONE_ESAMI_LAB.CODICE_ESAME) " & _
                     "WHERE      (CODICE_GRUPPO=" & vett(i).esami(k).righe(numRigaGruppo).codice & " ) " & _
                     "ORDER BY   ORDINE_VISUALIZZAZIONE"
            rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
            Do While Not rsEsami.EOF
                ReDim Preserve vett(i).esami(k).righe(UBound(vett(i).esami(k).righe) + 1)
                vett(i).esami(k).righe(UBound(vett(i).esami(k).righe)).codice = rsEsami("CODICE_ESAME")
                vett(i).esami(k).righe(UBound(vett(i).esami(k).righe)).nome = rsEsami("VOCI_ESAMINOME")
                If Not CBool(rsEsami("PN")) Then
                    vett(i).esami(k).righe(UBound(vett(i).esami(k).righe)).unita = rsEsami("UNITA")
                    vett(i).esami(k).righe(UBound(vett(i).esami(k).righe)).minmax = IIf(rsEsami("MAX") = "" And rsEsami("MIN") = "", "", rsEsami("MIN") & "÷" & rsEsami("MAX"))
                End If
                rsEsami.MoveNext
            Loop
            rsEsami.Close
        Next k
    Next i

    ' carica i valori
    For i = 1 To UBound(vett)
        For k = 1 To UBound(vett(i).esami)
            ' utente modificatore
            strSql = "SELECT    ANAMNESI_ESAMI.DATA, LOGIN.NOME, LOGIN.COGNOME  " & _
                    "FROM       (ANAMNESI_ESAMI " & _
                    "           INNER JOIN LOGIN ON LOGIN.KEY=ANAMNESI_ESAMI.UTENTE_MODIFICATORE) " & _
                    "WHERE      (CODICE_PAZIENTE=" & codicePaziente & _
                    condizione & " AND " & _
                    "           YEAR([ANAMNESI_ESAMI.DATA])=" & vett(i).anno & " AND " & _
                    "           CODICE_GRUPPO=" & vett(i).esami(k).righe(numRigaGruppo).codice & ") " & _
                    "ORDER BY   ANAMNESI_ESAMI.DATA DESC"
            rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
            Do While Not rsEsami.EOF
                For p = 1 To 12
                    If vett(i).esami(k).righe(numRigaGruppo).valori(p) = "" Then Exit For
                    If rsEsami("DATA") = CDate(vett(i).esami(k).righe(numRigaGruppo).valori(p) & "/" & vett(i).anno) Then
                        vett(i).esami(k).righe(numRigaUtente).valori(p) = Mid(rsEsami("COGNOME"), 1, 2) & " " & Mid(rsEsami("NOME"), 1, 2)
                        Exit For
                    End If
                Next p
                rsEsami.MoveNext
            Loop
            rsEsami.Close
            
            For j = 2 To UBound(vett(i).esami(k).righe)
                strSql = "SELECT    VALORE, DATA " & _
                        "FROM       (ANAMNESI_ESAMI " & _
                        "           INNER JOIN ESAMI_LAB ON ANAMNESI_ESAMI.KEY=ESAMI_LAB.CODICE_ANAMNESI_ESAMI) " & _
                        "WHERE      (CODICE_PAZIENTE=" & codicePaziente & _
                        condizione & " AND " & _
                        "           YEAR([DATA])=" & vett(i).anno & " AND " & _
                        "           CODICE_GRUPPO=" & vett(i).esami(k).righe(numRigaGruppo).codice & " AND " & _
                        "           CODICE_ESAME=" & vett(i).esami(k).righe(j).codice & " ) " & _
                        "ORDER BY   DATA DESC"
                rsEsami.Open strSql, cnPrinc, adOpenDynamic, adLockOptimistic, adCmdText
                Do While Not rsEsami.EOF
                    For p = 1 To 12
                        If vett(i).esami(k).righe(numRigaGruppo).valori(p) = "" Then Exit For
                        If rsEsami("DATA") = CDate(vett(i).esami(k).righe(numRigaGruppo).valori(p) & "/" & vett(i).anno) Then
                            Select Case rsEsami("VALORE")
                                Case -3
                                    valoreAppo = ""
                                Case -2
                                    valoreAppo = "NEG"
                                Case -1
                                    valoreAppo = "POS"
                                Case Else
                                    valoreAppo = VirgolaOrPunto(rsEsami("VALORE"), ",") & ""
                            End Select
                            vett(i).esami(k).righe(j).valori(p) = valoreAppo
                            Exit For
                        End If
                    Next p
                    rsEsami.MoveNext
                Loop
                rsEsami.Close
            Next j
        Next k
    Next i
    
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
        
    numPag = 1
    For i = 1 To UBound(vett)
        rsMain.AddNew
        numPag = numPag + 1
        rsMain("ANNO") = Format(numPag, "000") & vett(i).anno
        peso = pesoLivello1
        For k = 1 To UBound(vett(i).esami)
            peso = peso + pesoLivello2
            If peso > pesoTotale Then
                peso = pesoLivello1 + pesoLivello2
                rsMain.AddNew
                numPag = numPag + 1
                rsMain("ANNO") = Format(numPag, "000") & vett(i).anno
            End If
            Set rsFiglio1 = rsMain.Fields("RES1").Value
            rsFiglio1.AddNew
            rsFiglio1.Fields("ANNO") = rsMain.Fields("ANNO")
            rsFiglio1.Fields("CODICE_GRUPPO") = rsFiglio1.Fields("ANNO") & Format(vett(i).esami(k).righe(numRigaGruppo).codice, "000")
            rsFiglio1.Fields("NOME_GRUPPO") = vett(i).esami(k).righe(numRigaGruppo).nome
            For p = 1 To 12
                rsFiglio1.Fields("DATA" & p) = vett(i).esami(k).righe(numRigaGruppo).valori(p)
                rsFiglio1.Fields("UTENTE" & p) = vett(i).esami(k).righe(numRigaUtente).valori(p)
            Next p
            For j = 2 To UBound(vett(i).esami(k).righe)
                Set rsFiglio2 = rsFiglio1.Fields("RES2").Value
                rsFiglio2.AddNew
                rsFiglio2.Fields("CODICE_GRUPPO") = rsFiglio1.Fields("CODICE_GRUPPO")
                rsFiglio2.Fields("NOME_ESAME") = vett(i).esami(k).righe(j).nome
                rsFiglio2.Fields("CODICE_ESAME") = vett(i).esami(k).righe(j).codice
                rsFiglio2.Fields("V_MINMAX") = vett(i).esami(k).righe(j).minmax
                rsFiglio2.Fields("UNITA") = vett(i).esami(k).righe(j).unita
                
                For p = 1 To quantimesi
                    rsFiglio2.Fields("VALORE" & p) = vett(i).esami(k).righe(j).valori(p)
                 ' elimina gli esami non valorizzati
                    If quantimesi = 1 And rsFiglio2.Fields("VALORE" & p) = "" Then
                      rsFiglio2.Delete
                      peso = peso - pesoLivello3
                    Else
                      rsFiglio2.Fields("LEGENDA") = ""
                      rsFiglio2.Update
                    End If
                Next p
                peso = peso + pesoLivello3
         
                If peso > pesoTotale And Not j = UBound(vett(i).esami(k).righe) Then
                    rsMain.AddNew
                    numPag = numPag + 1
                    rsMain("ANNO") = Format(numPag, "000") & vett(i).anno
                    Set rsFiglio1 = rsMain.Fields("RES1").Value
                    rsFiglio1.AddNew
                    rsFiglio1.Fields("ANNO") = rsMain.Fields("ANNO")
                    rsFiglio1.Fields("CODICE_GRUPPO") = rsFiglio1.Fields("ANNO") & Format(vett(i).esami(k).righe(numRigaGruppo).codice, "000")
                    rsFiglio1.Fields("NOME_GRUPPO") = vett(i).esami(k).righe(numRigaGruppo).nome
                    For p = 1 To 12
                        rsFiglio1.Fields("DATA" & p) = vett(i).esami(k).righe(numRigaGruppo).valori(p)
                        rsFiglio1.Fields("UTENTE" & p) = vett(i).esami(k).righe(numRigaUtente).valori(p)
                    Next p
                    peso = pesoLivello1 + pesoLivello2
                End If
            Next j
            If i = UBound(vett) And k = UBound(vett(i).esami) Then
                rsFiglio2.AddNew
                rsFiglio2.Fields("CODICE_GRUPPO") = rsFiglio1.Fields("CODICE_GRUPPO")
                rsFiglio2.Fields("NOME_ESAME") = ""
                rsFiglio2.Fields("CODICE_ESAME") = 0
                rsFiglio2.Fields("V_MINMAX") = ""
                rsFiglio2.Fields("UNITA") = ""
                For p = 1 To 12
                    rsFiglio2.Fields("VALORE" & p) = ""
                Next p
                rsFiglio2.Fields("LEGENDA") = legenda
                rsFiglio2.Update
            End If
            rsFiglio1.Update
        Next k
        rsMain.Update
    Next i

    Set rptCartellaClinica_6 = Nothing
    Set rptCartellaClinica_6.DataSource = rsMain
    rptCartellaClinica_6.LeftMargin = 0
    rptCartellaClinica_6.RightMargin = 0
    If Not formPazienti Then
        rptCartellaClinica_6.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_6.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_6.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_6.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_6.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_6.PrintReport Not formPazienti, rptRangeAllPages
    Set rsEsami = Nothing
End Sub

'' Stampa terapia dialitica corrente
Public Sub StampaTerapiaDialiticaCorrente(codicePaziente As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 16
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(10) AS DATA_1, " & _
                "           NEW adVarChar(10) AS DATA_2, " & _
                "           NEW adVarChar(10) AS DATA_3, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adInteger AS SOMMINISTRAZIONE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    ' terapie non sospese
    strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DIALITICHE " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=FALSE " & _
             "ORDER BY  DATA DESC"
             
    Set rsTerapia = New Recordset
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Dialitica"
            .Fields("INTESTAZIONE_SOSPESA") = ""
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("NOTE")
                    .Fields("SOMMINISTRAZIONE") = rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = ""
                    .Fields("DATA_1") = rsTerapia("DATA_1") & ""
                    .Fields("DATA_2") = rsTerapia("DATA_2") & ""
                    .Fields("DATA_3") = rsTerapia("DATA_3") & ""
                    
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Dialitica"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = ""
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
        End With
    End If
    rsTerapia.Close
    
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_7 = Nothing
        Set rptCartellaClinica_7.DataSource = rsMain
        rptCartellaClinica_7.LeftMargin = 0
        rptCartellaClinica_7.RightMargin = 0
        rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        rptCartellaClinica_7.PrintReport True, rptRangeAllPages
    End If
End Sub

'' Stampa terapia dialitica sospesa
Public Sub StampaTerapiaDialiticaSospesa(codicePaziente As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 16
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(10) AS DATA_1, " & _
                "           NEW adVarChar(10) AS DATA_2, " & _
                "           NEW adVarChar(10) AS DATA_3, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adInteger AS SOMMINISTRAZIONE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic

    ' terapie sospese
    strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DIALITICHE " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=TRUE " & _
             "ORDER BY  DATA DESC"
             
    Set rsTerapia = New Recordset
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Dialitica - Sospesa"
            .Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("NOTE")
                    .Fields("SOMMINISTRAZIONE") = rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = rsTerapia("DATA_SOSPESA")
                    .Fields("DATA_1") = rsTerapia("DATA_1") & ""
                    .Fields("DATA_2") = rsTerapia("DATA_2") & ""
                    .Fields("DATA_3") = rsTerapia("DATA_3") & ""
                    
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Dialitica - Sospesa"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
            rsTerapia.Close
        End With
    End If
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_7 = Nothing
        Set rptCartellaClinica_7.DataSource = rsMain
        rptCartellaClinica_7.LeftMargin = 0
        rptCartellaClinica_7.RightMargin = 0
        rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        rptCartellaClinica_7.PrintReport True, rptRangeAllPages
    End If
End Sub

'' Stampa terapia dialitica
Public Sub StampaSettimaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strShape As String
    Dim strSql As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Const numMaxRecord As Integer = 16
    
    strShape = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS TITOLO, " & _
                "       NEW adVarChar(35) AS INTESTAZIONE_SOSPESA, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(10) AS DATA_1, " & _
                "           NEW adVarChar(10) AS DATA_2, " & _
                "           NEW adVarChar(10) AS DATA_3, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adInteger AS SOMMINISTRAZIONE, " & _
                "           NEW adLongVarChar as GIORNI, " & _
                "           NEW adVarChar(10) AS DATA_SOSPESA " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strShape, cnConn, adOpenStatic, adLockOptimistic
    
    ' terapie non sospese
    strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DIALITICHE " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=FALSE " & _
             "ORDER BY  DATA DESC"
             
    Set rsTerapia = New Recordset
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Dialitica In Corso"
            .Fields("INTESTAZIONE_SOSPESA") = ""
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("NOTE")
                    .Fields("SOMMINISTRAZIONE") = rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = ""
                    .Fields("DATA_1") = rsTerapia("DATA_1") & ""
                    .Fields("DATA_2") = rsTerapia("DATA_2") & ""
                    .Fields("DATA_3") = rsTerapia("DATA_3") & ""
                    
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Dialitica"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = ""
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
        End With
    End If
    rsTerapia.Close

    k = k + 1
    ' terapie sospese
    strSql = "SELECT    TERAPIE_DIALITICHE.*, MEDICINALI.NOME AS MEDICINALINOME " & _
             "FROM      (TERAPIE_DIALITICHE " & _
             "          INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
             "WHERE     CODICE_PAZIENTE=" & codicePaziente & " AND " & _
             "          SOSPESA=TRUE " & _
             "ORDER BY  DATA DESC"
             
    rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTerapia.EOF And rsTerapia.BOF) Then
        With rsMain
            .AddNew
            .Fields("TITOLO") = "Terapia Dialitica - Sospesa"
            .Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
            .Fields("LINK1") = k
            j = 0
            Do While Not rsTerapia.EOF
                Set rsFiglio = .Fields("Res1").Value
                With rsFiglio
                    .AddNew
                    .Fields("LINK1") = k
                    .Fields("DATA") = rsTerapia("DATA")
                    .Fields("MEDICINALE") = rsTerapia("MEDICINALINOME")
                    .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & " " & rsTerapia("NOTE")
                    .Fields("SOMMINISTRAZIONE") = rsTerapia("SOMMINISTRAZIONE")
                    .Fields("DATA_SOSPESA") = rsTerapia("DATA_SOSPESA")
                    .Fields("DATA_1") = rsTerapia("DATA_1") & ""
                    .Fields("DATA_2") = rsTerapia("DATA_2") & ""
                    .Fields("DATA_3") = rsTerapia("DATA_3") & ""
                    If CBool(rsTerapia("TUTTI_GIORNI")) Then
                        .Fields("GIORNI") = "Tutti"
                    Else
                        For i = 1 To 7
                            If CBool(rsTerapia("GIORNO" & i)) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & " " & UCase(Mid(WeekdayName(i, False, vbMonday), 1, 1)) & Mid(WeekdayName(i, False, vbMonday), 2, 2)
                            End If
                        Next i
                    End If
                    .Update
                End With
                j = j + 1
                If j = numMaxRecord Then
                    rsMain.Update
                    rsMain.AddNew
                    rsMain.Fields("TITOLO") = "Terapia Dialitica - Sospesa"
                    rsMain.Fields("INTESTAZIONE_SOSPESA") = "Dt.Sosp."
                    j = 0
                    k = k + 1
                    rsMain.Fields("LINK1") = k
                End If
                rsTerapia.MoveNext
            Loop
            rsTerapia.Close
        End With
    End If
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptCartellaClinica_7 = Nothing
        Set rptCartellaClinica_7.DataSource = rsMain
        rptCartellaClinica_7.LeftMargin = 0
        rptCartellaClinica_7.RightMargin = 0
        If Not formPazienti Then
            rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
            rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
        Else
            rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
            rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
            rptCartellaClinica_7.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
        End If
        rptCartellaClinica_7.PrintReport Not formPazienti, rptRangeAllPages
    End If
End Sub

'' Stampa anamnesi nefrologica
Public Sub StampaTerzaParte(formPazienti As Boolean, codicePaziente As Integer, Optional codiceId As Integer)
    Dim strSqlStampa As String
    Dim strSql As String
    
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsTabelle As Recordset           ' tabelle per la stampa nefrologia
        
    strSqlStampa = "       NEW adVarChar (150) as MALATTIA_RENALE_EDTA, " & _
                "       NEW adBoolean as ISTOLOGICA, " & _
                "       NEW adBoolean as TRATTAMENTO, " & _
                "       NEW adVarChar (10) as DATA_DIAGNOSI, " & _
                "       NEW adVarChar (10) as DATA_INIZIO, " & _
                "       NEW adVarChar (10) as DATA_INIZIO_SEDE, " & _
                "       NEW adVarChar (10) as DATA_FINE_SEDE, " & _
                "       NEW adVarChar (50) as SEDE, " & _
                "       NEW adBoolean as ATTESA_TRAPIANTO, " & _
                "       NEW adVarChar (50) as PRIMA_SEDE, " & _
                "       NEW adVarChar (50) as SECONDA_SEDE, " & _
                "       NEW adVarChar (25) as NOTE1, " & _
                "       NEW adVarChar (25) as NOTE2, " & _
                "       NEW adVarChar (20) as SOSPENSIONE1, " & _
                "       NEW adVarChar (20) as SOSPENSIONE2, " & _
                "       NEW adBoolean as PRECEDENTE_TRAPIANTO, " & _
                "       NEW adVarChar (10) as DATA_TRAPIANTO, " & _
                "       NEW adVarChar (50) as SEDE_TRAPIANTO, " & _
                "       NEW adBoolean as PRECEDENTE_ESPIANTO, " & _
                "       NEW adVarChar (10) as DATA_ESPIANTO, " & _
                "       NEW adVarChar (50) as SEDE_ESPIANTO "
    
    ' stringa di shape
    strSqlStampa = "SHAPE APPEND " & strSqlStampa
     
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    ' carica il recordset padre
    Set rsTabelle = New Recordset
    strSql = "SELECT    ANAMNESI_NEFROLOGICHE.*, EDTA.NOME AS EDTANOME, " & _
            "           CENTRI_PROVENIENZA1.NOME AS CENTRI_PROVENIENZA1NOME, " & _
            "           CENTRI_PROVENIENZA2.NOME AS CENTRI_PROVENIENZA2NOME, " & _
            "           CENTRI_PROVENIENZA3.NOME AS CENTRI_PROVENIENZA3NOME, " & _
            "           CENTRI_PROVENIENZA4.NOME AS CENTRI_PROVENIENZA4NOME, " & _
            "           CENTRI_PROVENIENZA5.NOME AS CENTRI_PROVENIENZA5NOME " & _
            " FROM      ((((((ANAMNESI_NEFROLOGICHE " & _
            "           LEFT OUTER JOIN EDTA ON EDTA.KEY=ANAMNESI_NEFROLOGICHE.CODICE_EDTA) " & _
            "           LEFT OUTER JOIN CENTRI_PROVENIENZA CENTRI_PROVENIENZA1 ON CENTRI_PROVENIENZA1.KEY=ANAMNESI_NEFROLOGICHE.SEDE0) " & _
            "           LEFT OUTER JOIN CENTRI_PROVENIENZA CENTRI_PROVENIENZA2 ON CENTRI_PROVENIENZA2.KEY=ANAMNESI_NEFROLOGICHE.SEDE1) " & _
            "           LEFT OUTER JOIN CENTRI_PROVENIENZA CENTRI_PROVENIENZA3 ON CENTRI_PROVENIENZA3.KEY=ANAMNESI_NEFROLOGICHE.SEDE2) " & _
            "           LEFT OUTER JOIN CENTRI_PROVENIENZA CENTRI_PROVENIENZA4 ON CENTRI_PROVENIENZA4.KEY=ANAMNESI_NEFROLOGICHE.SEDE3) " & _
            "           LEFT OUTER JOIN CENTRI_PROVENIENZA CENTRI_PROVENIENZA5 ON CENTRI_PROVENIENZA5.KEY=ANAMNESI_NEFROLOGICHE.SEDE4) " & _
            " Where     CODICE_PAZIENTE = " & codicePaziente

    rsTabelle.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsTabelle.EOF And rsTabelle.BOF) Then
        With rsMain
            .AddNew
            .Fields("MALATTIA_RENALE_EDTA") = rsTabelle("EDTANOME")
            .Fields("ISTOLOGICA") = CBool(rsTabelle("ISTOLOGICA"))
            .Fields("TRATTAMENTO") = CBool(rsTabelle("TRATTAMENTO_CONS"))
            .Fields("DATA_DIAGNOSI") = CStr(IIf(IsNull(rsTabelle("DATA0")), " --- ", rsTabelle("DATA0")))
            .Fields("DATA_INIZIO") = CStr(IIf(IsNull(rsTabelle("DATA1")), " --- ", rsTabelle("DATA1")))
            .Fields("DATA_INIZIO_SEDE") = CStr(IIf(IsNull(rsTabelle("DATA_INIZIO")), " --- ", rsTabelle("DATA_INIZIO")))
            .Fields("DATA_FINE_SEDE") = CStr(IIf(IsNull(rsTabelle("DATA_FINE")), " --- ", rsTabelle("DATA_FINE")))
            .Fields("SEDE") = rsTabelle("CENTRI_PROVENIENZA1NOME")
            .Fields("ATTESA_TRAPIANTO") = CBool(rsTabelle("ATTESA_TRAPIANTO"))
            .Fields("PRIMA_SEDE") = rsTabelle("CENTRI_PROVENIENZA2NOME")
            .Fields("SECONDA_SEDE") = rsTabelle("CENTRI_PROVENIENZA3NOME")
            .Fields("NOTE1") = rsTabelle("NOTE1")
            .Fields("NOTE2") = rsTabelle("NOTE2")
            .Fields("SOSPENSIONE1") = IIf(CBool(rsTabelle("SOSPENSIONE1")), "SOSP. TEMP.", "")
            .Fields("SOSPENSIONE2") = IIf(CBool(rsTabelle("SOSPENSIONE2")), "SOSP. TEMP.", "")
            .Fields("PRECEDENTE_TRAPIANTO") = CBool(rsTabelle("PREC_TRAPIANTO"))
            .Fields("DATA_TRAPIANTO") = CStr(IIf(IsNull(rsTabelle("DATA2")), " --- ", rsTabelle("DATA2")))
            .Fields("SEDE_TRAPIANTO") = rsTabelle("CENTRI_PROVENIENZA4NOME")
            .Fields("PRECEDENTE_ESPIANTO") = CBool(rsTabelle("PREC_ESPIANTO"))
            .Fields("DATA_ESPIANTO") = CStr(IIf(IsNull(rsTabelle("DATA3")), " --- ", rsTabelle("DATA3")))
            .Fields("SEDE_ESPIANTO") = rsTabelle("CENTRI_PROVENIENZA5NOME")
        End With
    End If
    rsTabelle.Close

    Set rptCartellaClinica_3 = Nothing
    Set rptCartellaClinica_3.DataSource = rsMain
    If Not formPazienti Then
        rptCartellaClinica_3.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = ""
        rptCartellaClinica_3.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = ""
    Else
        rptCartellaClinica_3.Sections("Intestazione").Controls.Item("lblId").Caption = codiceId
        rptCartellaClinica_3.Sections("Intestazione").Controls.Item("lblIDLabel").Caption = "ID: "
        rptCartellaClinica_3.Sections("Intestazione").Controls.Item("lblCartellaClinica").Caption = "CARTELLA CLINICA"
    End If
    rptCartellaClinica_3.PrintReport Not formPazienti, rptRangeAllPages
End Sub

''
' Inizia la chiamata al form per la barra
Public Sub StartProgressBar(inValoreMax As Integer, inValoreStart As Integer, ByRef frm As Form)
    frmBarra.Show
    frmBarra.prgBar.max = inValoreMax
    frmBarra.prgBar.Value = inValoreStart
    frmBarra.Refresh
    frm.Enabled = False
End Sub

''
' Ferma il form della barra
Public Sub StopProgressBar(ByRef frm As Form)
    frm.Enabled = True
    Unload frmBarra
End Sub

''
' Effettua un ucase di tutte le txt e aggiunge uno spazio
'
' @param frm form da analizzare
' @param
' @return
' @remarks evita le cbo e txtEmail
Public Sub SuperUcase(frm As Form)
    Dim obj As Object
    On Error Resume Next
    For Each obj In frm.Controls
        ' evita le cbo e le e-mail
        If Mid(obj.Name, 1, 3) <> "cbo" And obj.Name <> "txtEmail" Then
            obj.Text = UCase(obj.Text) & ""
        End If
    Next
End Sub

''
' Elimina la x del form
'
' @param Handle handle del form
' @param
' @return
' @remarks
Public Sub TakeCloseOff(handle As Long)
    Dim SysMenHandle As Long, RetVal As Long
    'Prende l'handle del menu di sistema di Form1
    SysMenHandle = GetSystemMenu(handle, 0)
    'Elimina la voce Close
    RetVal = RemoveMenu(SysMenHandle, 6, MF_BYPOSITION)
    'Elimina il separatore che ora si trova in basso
    RetVal = RemoveMenu(SysMenHandle, 5, MF_BYPOSITION)
End Sub

''
' Disegna una scritta nella picture
Public Sub Text3D(pic As PictureBox, Strng As String, Fnt As String, Font_size As Integer, XVal As Integer, YVal As Integer, Depth As Integer, Redcol As Integer, Greencol As Integer, Bluecol As Integer)
    Dim i As Integer
    Dim ShadowY As Integer
    Dim ShadowX As Integer
    pic.AutoRedraw = True
    pic.FontSize = Font_size
    pic.Font = Fnt
    pic.ForeColor = RGB(Redcol, Greencol, Bluecol)
    ShadowY = YVal
    ShadowX = XVal
    For i = 0 To Depth
        pic.CurrentX = ShadowX - i
        pic.CurrentY = ShadowY + i
        If i = Depth Then _
            pic.ForeColor = RGB(Redcol + 80, Greencol + 80, Bluecol + 80)
            pic.Print Strng
    Next i
    pic.AutoRedraw = False
End Sub

''
' Verifica se il click del mouse è avvenuto su una cella o nel vuoto della inFlex
'
' @param inFlex griglia
' @param
' @return true se il click è nella cella, altrimenti false
' @remarks restituisce false se il click è sul titolo
Public Function VerificaClickFlx(inFlex As MSFlexGrid) As Boolean
    Dim colPos As Long
    Dim rowPos As Long
    Dim PuntoX As Integer
    Dim PuntoY As Integer
    ' verifica se il click è stato fatto sul titolo
    If inFlex.Row = 0 Then
        VerificaClickFlx = False
        Exit Function
    End If
    With inFlex
        colPos = .colPos(.Cols - 1) + .ColWidth(.Cols - 1)
        rowPos = .rowPos(.Rows - 1) + .RowHeight(.Rows - 1)
    End With
    Call PosizioneCursore(PuntoX, PuntoY, inFlex.hWnd)
    If PuntoX > colPos Or PuntoY > rowPos Then
        VerificaClickFlx = False
    Else
        VerificaClickFlx = True
    End If
End Function

''
' Sostituisce il chrCosaSostituire in num e lo sostituisce con chrSostituto
'
' @param num numero da analizzare
' @param chrCosaSostituire carattere da trovare e sostituire
' @return nuovo numero
' @remarks utile per le incongruenze dei decimali tra db e vb
Public Function VirgolaOrPunto(ByRef num As String, chrCosaSostituire As String) As String
    On Error GoTo gestione
    Dim chrSostituto As String * 1
    Dim pos As Integer
    chrSostituto = IIf(chrCosaSostituire = ",", ".", ",")
    pos = InStr(1, num, chrCosaSostituire)
    If pos <> 0 Then Mid(num, pos, 1) = chrSostituto
    VirgolaOrPunto = num
    Exit Function
gestione:
    VirgolaOrPunto = num
End Function

Public Function FileCopyEx(Source As String, Destination As String)
Dim sFile As String, sSPath As String
Dim sDPath As String

On Error GoTo ErrHandle
sSPath = Mid$(Source, 1, InStrRev(Source, "\"))
sDPath = Mid$(Destination, 1, InStrRev(Destination, "\"))
sFile = Dir$(Source)
Do While Len(sFile) > 0
  FileCopy sSPath & sFile, sDPath & sFile
  sFile = Dir$
  DoEvents
Loop
ErrHandle:
  If Err.Number > 0 Then
    MsgBox Err.Description, vbExclamation, Err.Number
    Exit Function
  End If
End Function

'Public Sub Select_Data()
'Dim periodo As Integer
'        Unload frmTrova
 '       tTrova.isOpenFromEsamiPrescriz = True
 '       frmPannelloPeriodo.LetSenzaData = False
 '       frmPannelloPeriodo.Show 1
 '       periodo = frmPannelloPeriodo.GetPeriodo
 '       laData = frmPannelloPeriodo.getData
 '       Unload frmPannelloPeriodo
 '       If periodo = -1 Then
 '           scelta = True
 '           Unload 'Me
 '       Else
'            cmdTrova_Click
'           If scelta = True Then
'            Exit Do
'           End If
 '           If tTrova.keyReturn = 0 Then
 '               scelta = True
 '               Unload Me
 '           End If
 '       End If

'End Sub
