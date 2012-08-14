VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVisualizzaScansione 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualizza scansione"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStampa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   9000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      LargeChange     =   100
      Left            =   10080
      SmallChange     =   10
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      SmallChange     =   10
      TabIndex        =   6
      Top             =   7320
      Width           =   10215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   10215
      Begin VB.CommandButton cmdCambiaPag 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCambiaPag 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
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
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdChiudi 
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
         Height          =   375
         Left            =   8760
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdZoom 
         Height          =   375
         Index           =   1
         Left            =   840
         Picture         =   "frmVisualizzaScansione.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdZoom 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmVisualizzaScansione.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "&Elimina"
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
         Height          =   375
         Left            =   7200
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.Slider sliderZoom 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
   End
   Begin VB.PictureBox picContenitore 
      AutoSize        =   -1  'True
      Height          =   7455
      Left            =   120
      ScaleHeight     =   493
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   120
      Width           =   10185
      Begin VB.PictureBox picScansione 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7320
         Left            =   0
         ScaleHeight     =   488
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   664
         TabIndex        =   9
         Top             =   0
         Width           =   9960
      End
      Begin VB.Image imgVisualizza 
         Height          =   7320
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   9960
      End
   End
   Begin MSComDlg.CommonDialog cdlStampa 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVisualizzaScansione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numPag As Integer
Private nomeFile As String
Private codiceRecord As Integer
Private codiceCentro As Integer

Dim v_nomeTabella(4) As String
Dim paginaCorrente As Integer
Dim nomeFileSingolo As String

Public Property Let letcodiceRecord(ByVal vcodiceRecord As Integer)
    codiceRecord = vcodiceRecord
End Property

Public Property Let letNumPag(ByVal vnumPag As Integer)
    numPag = vnumPag
End Property

Public Property Let LetNomeFile(ByVal vnomeFile As String)
    nomeFile = vnomeFile
End Property

Public Property Let LetcodiceCentro(ByVal vcodiceCentro As String)
    codiceCentro = vcodiceCentro
End Property

Private Sub ImpostaImage(picBox As PictureBox, immagine As Picture, sizeWidth As Single, sizeHeight As Single)
    On Error GoTo gestione
    picBox.Picture = LoadPicture("")
    picBox.Width = sizeWidth
    picBox.Height = sizeHeight
    If picBox.Width < picContenitore.Width Then
        picBox.Left = (picContenitore.Width - picBox.Width) / 2
    Else
        picBox.Left = 0
    End If
    If picBox.Height < picContenitore.Height Then
        picBox.Top = (picContenitore.Height - picBox.Height) / 2
    Else
        picBox.Top = 0
    End If
    picBox.AutoRedraw = True
    picBox.PaintPicture immagine, 0, 0, sizeWidth, sizeHeight
    picBox.Picture = picBox.Image
    picBox.AutoRedraw = False
    Exit Sub
gestione:
    MsgBox "Grandezza immagine errata", vbCritical, "Impostazione zoom"
    Unload Me
End Sub

Private Sub ImpostaScrollBar()
    If picContenitore.Width <= picScansione.Width Then
        With HScroll1
            .max = (picScansione.Width - picContenitore.Width)
            .Enabled = True
        End With
    Else
        With HScroll1
            .Enabled = False
        End With
    End If
    If picContenitore.Height <= picScansione.Height Then
        With VScroll1
            .max = (picScansione.Height - picContenitore.Height)
            .Enabled = True
        End With
    Else
        With VScroll1
            .Enabled = False
        End With
    End If
End Sub

Private Function CaricaPagina(nomeFile As String) As Boolean

    Dim wImage As Single
    Dim hImage As Single
    Dim coeffProporzione As Single
    
    If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
        ' imposta l immagine base
        picStampa.Picture = LoadPicture(structApri.pathDB & "\" & nomeFile & ".jpg")
        wImage = picStampa.Picture.Width
        hImage = picStampa.Picture.Height
        ' tiene ferma l'altezza e riduce la larghezza in proporzione
        coeffProporzione = hImage / imgVisualizza.Height
        imgVisualizza.Width = wImage / coeffProporzione
        imgVisualizza.Stretch = True
        imgVisualizza.Picture = LoadPicture(structApri.pathDB & "\" & nomeFile & ".jpg")
        ' imposta l immagine nella picturebox
        Call ImpostaImage(picScansione, imgVisualizza.Picture, imgVisualizza.Width, imgVisualizza.Height)
        Call ImpostaScrollBar
        CaricaPagina = True
        cmdStampa.Enabled = True
    ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
        ShellExecute Me.hWnd, "open", structApri.pathDB & "\" & nomeFile & ".pdf", "", "", 5
        picScansione.Picture = Nothing
        picScansione.Left = 0
        picScansione.Top = 0
        picScansione.Width = 650
        picScansione.Height = 490
        Call Text3D(picScansione, "FILE PDF", "Times New Roman", 50, 200, 200, 5, 128, 128, 128)   'modificare tutti e tre i numeri
        CaricaPagina = False
        cmdStampa.Enabled = False
    End If
End Function

Private Sub cmdCambiaPag_Click(Index As Integer)
    If Index = 0 Then
        paginaCorrente = paginaCorrente - 1
        cmdCambiaPag(1).Enabled = True
        If paginaCorrente = 1 Then
            cmdCambiaPag(0).Enabled = False
        End If
    Else
        paginaCorrente = paginaCorrente + 1
        cmdCambiaPag(0).Enabled = True
        If paginaCorrente = numPag Then
            cmdCambiaPag(1).Enabled = False
        End If
    End If
    Call CaricaPagina(nomeFileSingolo & Format(paginaCorrente, "00"))
End Sub

Private Sub Form_Activate()
    v_nomeTabella(0) = "SCAN_ESAMI_STRUMENTALI"
    v_nomeTabella(1) = "SCAN_PSICO_SOCIALE"
    v_nomeTabella(2) = "SCAN_TRAPIANTI"
    v_nomeTabella(3) = "SCAN_TRATT_ACQUE"
    v_nomeTabella(4) = "SCAN_DOCUMENTI_PAZIENTI"
    
    nomeFileSingolo = Mid(nomeFile, 1, Len(nomeFile) - 2)
    paginaCorrente = Int(Mid(nomeFile, Len(nomeFile) - 1, 2))
    If numPag = 1 Then
        cmdCambiaPag(0).Enabled = False
        cmdCambiaPag(1).Enabled = False
    Else
        If paginaCorrente = 1 Then
            cmdCambiaPag(0).Enabled = False
        ElseIf paginaCorrente = numPag Then
            cmdCambiaPag(1).Enabled = False
        End If
    End If
    
    Call CaricaPagina(nomeFile)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdElimina_Click()
    Dim i As Integer
    Dim rsDataset As Recordset
    Dim condTrapianto As String
    
    If tDocumentiEsterni = tpSCANTRAPIANTI Then
        condTrapianto = " AND CODICE_CENTRO=" & codiceCentro
    End If
    
    If MsgBox("Sicuro di voler eliminare la pagina selezionata?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminazione") = vbYes Then
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT * FROM " & v_nomeTabella(tDocumentiEsterni) & " WHERE CODICE_SCHEDA=" & codiceRecord & condTrapianto, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        Do While Not rsDataset.EOF
            If Right(rsDataset("NOME_FILE"), 2) = Format(paginaCorrente, "00") Then
                nomeFile = rsDataset("NOME_FILE")
                rsDataset.Delete
                ' elimina anche il file
                If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                    Kill structApri.pathDB & "\" & nomeFile & ".jpg"
                ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                    Kill structApri.pathDB & "\" & nomeFile & ".pdf"
                End If
                Exit Do
            End If
            rsDataset.MoveNext
        Loop
        rsDataset.Close
        
        rsDataset.Open "SELECT * FROM " & v_nomeTabella(tDocumentiEsterni) & " WHERE CODICE_SCHEDA=" & codiceRecord & condTrapianto, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
        If Not (rsDataset.EOF And rsDataset.BOF) Then
            i = 1
            Do While Not rsDataset.EOF
                nomeFile = rsDataset("NOME_FILE")
                rsDataset("NOME_FILE") = nomeFileSingolo & Format(i, "00")
                rsDataset.Update
                If Dir(structApri.pathDB & "\" & nomeFile & ".jpg") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".jpg" As structApri.pathDB & "\" & nomeFileSingolo & Format(i, "00") & ".jpg"
                ElseIf Dir(structApri.pathDB & "\" & nomeFile & ".pdf") <> "" Then
                    Name structApri.pathDB & "\" & nomeFile & ".pdf" As structApri.pathDB & "\" & nomeFileSingolo & Format(i, "00") & ".pdf"
                End If
                i = i + 1
                rsDataset.MoveNext
            Loop
            paginaCorrente = 1
            numPag = numPag - 1
            If numPag = 1 Then
                cmdCambiaPag(0).Enabled = False
                cmdCambiaPag(1).Enabled = False
            Else
                If paginaCorrente = 1 Then
                    cmdCambiaPag(0).Enabled = False
                ElseIf paginaCorrente = numPag Then
                    cmdCambiaPag(1).Enabled = False
                End If
            End If
    
            Call CaricaPagina(nomeFileSingolo & "01")
        Else
            Select Case tDocumentiEsterni
                Case tpSCANESAMISTRUMENTALI
                    frmEsamiStrumentali.SalvaEliminazioneReferto nomeFile
                Case tpSCANDOCPAZIENTI
                    frmScanDocumenti.LetAggiorna = True
            End Select
            Unload Me
        End If
        rsDataset.Close
        Set rsDataset = Nothing
    End If
End Sub

Private Sub cmdStampa_Click()
    On Error GoTo gestione
    Dim i As Integer
    
    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    For i = 1 To numPag
        If CaricaPagina(nomeFileSingolo & Format(i, "00")) Then
            Printer.ScaleMode = vbPixels
            Printer.PaintPicture picStampa.Picture, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
            Printer.EndDoc
        End If
    Next i
    
    Unload Me
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub cmdZoom_Click(Index As Integer)
    Dim Incremento As Integer
    Dim altezza As Single
    Dim larghezza As Single
    
    If Index = 0 Then
        Incremento = 1
    Else
        Incremento = -1
    End If
    
    altezza = picScansione.Height + Incremento * (picScansione.Height * sliderZoom.Value / 10)
    larghezza = picScansione.Width + Incremento * (picScansione.Width * sliderZoom.Value / 10)
    imgVisualizza.Width = larghezza
    imgVisualizza.Height = altezza
    ' imposta l immagine nella picturebox ridimensionandola
    Call ImpostaImage(picScansione, imgVisualizza.Picture, larghezza, altezza)
    Call ImpostaScrollBar
End Sub

Private Sub picContenitore_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub picScansione_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub VScroll1_Change()
   picScansione.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   picScansione.Top = -VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   picScansione.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
   picScansione.Left = -HScroll1.Value
End Sub

