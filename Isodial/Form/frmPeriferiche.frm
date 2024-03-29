VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPeriferiche 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ripristina archivi"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   Icon            =   "frmPeriferiche.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxGriglia 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      MousePointer    =   15
      FormatString    =   "| Nome file                 | Data             "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPeriferiche.frx":000C
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3495
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdRipristina 
         Cancel          =   -1  'True
         Caption         =   "&Ripristina"
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
         Left            =   280
         TabIndex        =   2
         Top             =   240
         Width           =   1260
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
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmPeriferiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' <b>Progetto</b>:      Isodial.vbp
'
' <b>Tipo e nome modulo</b>:        Form - frmPeriferiche.frm
'
' <b>Descrizione</b>: Pannello Ripristina Archivio per ricaricare il volume cryptato dalla penna usb
'
' @remarks
'
' @author
'
' @date 19/03/2011 12.41
Option Explicit

' struttura per caricare il file dati.dat che tiene traccia dei backup sulla penna usb
Private Type structFile
    data As Date
    num As Integer
End Type
Dim records() As structFile
Dim numFile As Integer
Dim lettera As String

Private Sub Form_Load()
    Dim i As Integer
    Me.Left = 10
    Me.Top = 0
    With flxGriglia
        .ColWidth(0) = 0
        .Row = 0
        For i = 1 To 2
            .Col = i
            .CellFontBold = True
            .ColAlignment(i) = vbLeftJustify
        Next i
        .MousePointer = flexCustom
    End With
    Call TakeCloseOff(Me.hWnd)
    If VerificaDiscoRimovibile(lettera) Then
        Call LeggiDati
    End If
    flxGriglia.Row = 0
End Sub

'' Legge il file Dati.dat e carica i backup nella flx

Private Sub LeggiDati()
    Dim numRecord As Integer
    Dim i As Integer

    ReDim records(0)
    flxGriglia.Rows = 1
    
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
        numRecord = UBound(records) + 1
        Call BubbleSort(records)
        
        ' carica la griglia
        For i = 0 To numRecord - 1
            With flxGriglia
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = records(i).num
                .TextMatrix(.Rows - 1, 1) = nomeVolume & records(i).num
                .TextMatrix(.Rows - 1, 2) = records(i).data
            End With
        Next i
        
    End If
End Sub

' Ordina un array di structFile
Private Sub BubbleSort(ByRef MioArray() As structFile)
    Dim i As Integer
    Dim j As Integer
    Dim flag As Boolean
    Dim Temp As structFile
    flag = True
    i = UBound(MioArray, 1)
    Do While (i <> LBound(MioArray, 1) And flag = True)
        flag = False
        For j = LBound(MioArray, 1) To i - 1
            If MioArray(j).data < MioArray(j + 1).data Then
                Temp = MioArray(j)
                MioArray(j) = MioArray(j + 1)
                MioArray(j + 1) = Temp
                flag = True
            End If
        Next j
        i = i - 1
    Loop
End Sub

' Mostra un msg di errore
Private Sub GestioneErrore()
    Dim strMsg As String
    Select Case Err.Number
        Case 68
            strMsg = "L'unit� non � disponibile"
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
            strMsg = "Archivio inesistente"
        Case 62
            strMsg = "L'archivio risulta danneggiato"
        Case Else
            strMsg = Err.Description
    End Select
    Screen.MousePointer = 0
    Me.Enabled = True
    MsgBox strMsg, vbCritical, "ATTENZIONE"
    Call ApriVolume
End Sub

' Ripristina il volume selezionato
Private Sub Ripristina()
    On Error GoTo gestioneerror
    Dim ret As Double
    Dim numClient As Integer
       
    ' prima chiude la connessione
    Set cnPrinc = Nothing
    Set cnTrac = Nothing
    ' chiude la condivisione
    Call Shell("NET SHARE RISORSA /DELETE", vbHide)
    ' smonta il volume
    ret = Shell(structApri.pathTrueCrypt & "\TrueCrypt.exe /d X /q /s /f", vbHide)
  
    Screen.MousePointer = cc2Hourglass
    numFile = frmPeriferiche!flxGriglia.TextMatrix(frmPeriferiche!flxGriglia.Row, 0)
' If VerificaDiscoRimovibile(lettera) And Dir(lettera & ":\" & nomeVolume & numFile) <> "" Then
    Dim fileorigine As String
    Dim filedestinazione As String
    fileorigine = lettera & ":\" & nomeVolume & numFile
    filedestinazione = structApri.pathVolume & "\" & nomeVolume & numFile
  ' ripristina il file

    Call CopiaFile(fileorigine, filedestinazione, ProgressBar1)
   
  ' FileCopy lettera & ":\" & nomeVolume & numFile, structApri.pathVolume & "\" & nomeVolume & numFile
  ' elimina il vecchio database
    Kill structApri.pathVolume & "\" & nomeVolume
    Name structApri.pathVolume & "\" & nomeVolume & numFile As structApri.pathVolume & "\" & nomeVolume

    Screen.MousePointer = 0
    Me.SetFocus
    
    Call ApriVolume

    Exit Sub

gestioneerror:
    Call GestioneErrore
End Sub

Private Sub ApriVolume()
    Call VerificaErrori
    Call MontaVolume
    Call CaricaDati
    ' verifica che il db non sia corrotto
 '   If Not nonCorrotto Then
 '       MsgBox "Impossibile procedere" & vbCrLf & "Ripristinare un precedente backup o contattare l'autore" & vbCrLf & "Accesso consentito al solo amministratore di sistema", vbCritical, "ATTENZIONE!!! DATABASE CORROTTO"
 '       isCorrotto = True
 '   Else
        isCorrotto = False
 '   End If
End Sub

Private Sub cmdRipristina_Click()
    If VerificaDiscoRimovibile(lettera) And Dir(lettera & ":\" & nomeVolume & numFile) = "" Then
       MsgBox "Impossibile procedere al ripristino - Archivio inesistente", vbCritical, "ATTENZIONE!!!"
       Exit Sub
    ElseIf flxGriglia.Row <> 0 Then
        If MsgBox("ATTENZIONE!!! Il ripristino sovrascrive tutti i dati attuali." & vbCrLf & "Sicuro di voler ripristinare i dati precedenti?", vbQuestion + vbYesNo, "RIPRISTINO ARCHIVI") = vbNo Then
            Exit Sub
        End If
        Call Ripristina
        MsgBox "RIPRISTINO ARCHIVIO EFFETTUATO CORRETTAMENTE", vbInformation, "RIPRISTINO ARCHIVI"
        Unload Me
    Else
        MsgBox "Selezionare l'archivio da ripristinare", vbCritical, "ATTENZIONE!!!"
    End If
End Sub

Private Sub cmdChiudi_Click()
    If isCorrotto Then
        End
    Else
        Unload Me
    End If
End Sub

Private Sub flxGriglia_Click()
    flxGriglia.SetFocus
    If VerificaClickFlx(flxGriglia) = False Then
        ' discolora
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1, True)
        ' annulla le row e col
        flxGriglia.Row = 0
        flxGriglia.Col = 0
    Else
        Call ColoraFlx(flxGriglia, flxGriglia.Cols - 1)
    End If
End Sub

Private Sub flxGriglia_dblClick()
    If VerificaClickFlx(flxGriglia) = False Then Exit Sub
    Call Ripristina
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
'----------------------------------------
