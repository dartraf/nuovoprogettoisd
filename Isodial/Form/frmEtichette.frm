VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEtichette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etichette per provette"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton optPomeriggio 
         Caption         =   "Pomeriggio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optSera 
         Caption         =   "Sera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMattina 
         Caption         =   "Mattina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optPerTuttiPazienti 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblPazienti 
         AutoSize        =   -1  'True
         Caption         =   "Tutti i pazienti"
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Per turno:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      Begin VB.CommandButton cmdStampa44 
         Caption         =   "&Stampa 44 etichette"
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
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
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
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStampa40 
         Caption         =   "S&tampa 40 etichette"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1935
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
Attribute VB_Name = "frmEtichette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim v_note(1 To 4) As String
Dim numMaxCell As Integer

Private Sub Form_Load()

    ' carica il turno di default
    
    If Int(Hour(Now)) < 13 Then                                 'Mattina
        optMattina.Value = True
    ElseIf Int(Hour(Now)) > 12 And Int(Hour(Now)) < 18 Then     'Pomeriggio
        optPomeriggio.Value = True
    Else
        optSera.Value = True                                    'Sera
    End If
    
    v_note(1) = "Siero-Pre"
    v_note(2) = "Siero-Post"
    v_note(3) = "Emocromo"
    v_note(4) = ""
End Sub

Private Function AdattaStr(nome As String, Optional lung As Integer = 34) As String
    AdattaStr = Left(nome, 13) & Space(lung - Len(nome)) 'limita nome campo a 12 chr
    If Len(nome) = 13 Then  ' controlla lunghezza nome per adattarlo alle colonne
      AdattaStr = Left(nome, 13) & Space(lung - Len(nome) - 2)
    ElseIf Len(nome) = 4 Then
      AdattaStr = Left(nome, 4) & Space(lung - Len(nome) + 2)
      ElseIf Len(nome) = 5 Then
      AdattaStr = Left(nome, 5) & Space(lung - Len(nome) + 2)
    End If
End Function

'' stampa etichette per turno
Private Sub StampaEtichettePerTurno()
    Dim rsPazienti As New Recordset
    Dim tipostrTurno As String
    Dim giorno As Integer
    Dim i As Integer
    Dim strSql As String
    
    If optMattina.Value Then
        tipostrTurno = "AM_INIZIO"
    ElseIf optPomeriggio.Value Then
        tipostrTurno = "PM_INIZIO"
    Else
        tipostrTurno = "SR_INIZIO"
    End If
    giorno = Weekday(date, vbMonday)
    
    i = 0
    Printer.Print
    Printer.Print
    If numMaxCell = 10 Then     'aggiunge quattro righe per i fogli da 40 etichette
       Printer.Print
       Printer.Print
       Printer.Print
       Printer.Print
    End If

    strSql = "SELECT    COGNOME, NOME, DATA_NASCITA " & _
            "FROM       (PAZIENTI " & _
            "           INNER JOIN TURNI ON TURNI.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
            "WHERE      (STATO=0 AND " & _
            "           TURNI." & tipostrTurno & giorno & "<>"""" ) " & _
            "ORDER BY   COGNOME"
    Set rsPazienti = New Recordset
    rsPazienti.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsPazienti.EOF

        Printer.Print vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME"))
        Printer.Print vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME"))
        Printer.Print vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA"))
        Printer.Print vbTab & AdattaStr(v_note(1)) & vbTab & vbTab & AdattaStr(v_note(2)) & vbTab & vbTab & AdattaStr(v_note(3)) & vbTab & vbTab & AdattaStr(v_note(4))

        Printer.Print vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione)
        Printer.Print
        Printer.Print
        i = i + 1
        If i < numMaxCell Then
            Printer.Print
        End If
        If numMaxCell = 10 And i = numMaxCell Then   'se foglio da 40 etichette
             Printer.NewPage                         'salta pagina e aggiunge righe
             Printer.Print                           'azzera contatore
             Printer.Print
             Printer.Print
             Printer.Print
             Printer.Print
             Printer.Print
             i = 0
         ElseIf numMaxCell = 11 And i = numMaxCell Then   'se foglio da 44 etichette
             i = 0                                        ' azzera solo contatore
         End If
           
        rsPazienti.MoveNext
    Loop
    rsPazienti.Close
    Printer.EndDoc
    Set rsPazienti = Nothing
End Sub

'' stampa etichette per tutti i pazienti
Private Sub StampaEtichettePerPaziente()
    Dim rsPazienti As New Recordset
    Dim i As Integer

    i = 0
    Printer.Print
    Printer.Print
    If numMaxCell = 10 Then     'aggiunge quattro righe per i fogli da 40 etichette
       Printer.Print
       Printer.Print
       Printer.Print
       Printer.Print
    End If

    Set rsPazienti = New Recordset
    rsPazienti.Open "SELECT * FROM PAZIENTI WHERE STATO=0 ORDER BY COGNOME", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsPazienti.EOF

        Printer.Print vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME")) & vbTab & vbTab & AdattaStr(rsPazienti("COGNOME"))
        Printer.Print vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME")) & vbTab & vbTab & AdattaStr(rsPazienti("NOME"))
        Printer.Print vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA")) & vbTab & vbTab & AdattaStr(rsPazienti("DATA_NASCITA"))
        Printer.Print vbTab & AdattaStr(v_note(1)) & vbTab & vbTab & AdattaStr(v_note(2)) & vbTab & vbTab & AdattaStr(v_note(3)) & vbTab & vbTab & AdattaStr(v_note(4))

        Printer.Print vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione) & vbTab & vbTab & AdattaStr(structIntestazione.sRagione)
        Printer.Print
        Printer.Print
        i = i + 1
        If i < numMaxCell Then
            Printer.Print
        End If
        If numMaxCell = 10 And i = numMaxCell Then   'se foglio da 40 etichette
             Printer.NewPage                         'salta pagina e aggiunge righe
             Printer.Print                           'azzera contatore
             Printer.Print
             Printer.Print
             Printer.Print
             Printer.Print
             Printer.Print
             i = 0
         ElseIf numMaxCell = 11 And i = numMaxCell Then   'se foglio da 44 etichette
             i = 0                                        ' azzera solo contatore
         End If
           
        rsPazienti.MoveNext
    Loop
    rsPazienti.Close
    Printer.EndDoc
    Set rsPazienti = Nothing
End Sub

Private Sub SceltaStampa()
    On Error GoTo gestione

    cdlStampa.Flags = &H40  ' Finestra dialogo Imposta stampante.
    cdlStampa.CancelError = True
    cdlStampa.ShowPrinter
    
    If optMattina.Value = True Or optPomeriggio.Value = True Or optSera.Value = True Then
       Call StampaEtichettePerTurno
    Else
       Call StampaEtichettePerPaziente
    End If
     
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdStampa40_Click()
    numMaxCell = 10
    Call SceltaStampa
End Sub

Private Sub cmdStampa44_Click()
    numMaxCell = 11
    Call SceltaStampa
End Sub

