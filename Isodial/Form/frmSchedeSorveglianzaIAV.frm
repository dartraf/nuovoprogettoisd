VERSION 5.00
Begin VB.Form frmSchedeSorveglianzaFAV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schede Sorveglianza FAV"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmSegni 
      Caption         =   "Segni e Sintomi locali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   840
      Width           =   10575
      Begin VB.OptionButton optNoEritema 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optNoDolore 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optNoGonfiore 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optNoInfiltrazione 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optNoPresenzaFremiti 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton optSiEritema 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optSiDolore 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optSiGonfiore 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton optSiInfiltrazione 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton optSiPresenzaFremiti 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton optEritemaLieve 
         Caption         =   "Lieve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Lieve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Lieve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Lieve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Lieve"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton optEritemaMeno 
         Caption         =   "Meno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Meno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Meno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Meno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option20 
         Caption         =   "Meno"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optEritemaGrave 
         Caption         =   "Grave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option22 
         Caption         =   "Grave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option23 
         Caption         =   "Grave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option24 
         Caption         =   "Grave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option25 
         Caption         =   "Grave"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eritema"
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
         Index           =   11
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dolore"
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
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gonfiore"
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
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Infiltrazione"
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
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Presenza fremiti"
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
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   360
         Picture         =   "frmSchedeSorveglianzaIAV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anni"
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
         Index           =   3
         Left            =   11280
         TabIndex        =   7
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   6480
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cognome"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblEta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11880
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblCognome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   960
      TabIndex        =   8
      Top             =   7080
      Width           =   7815
      Begin VB.CommandButton cmdMemorizza 
         Caption         =   "&Memorizza"
         CausesValidation=   0   'False
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
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdChiudi 
         Caption         =   "&Chiudi"
         CausesValidation=   0   'False
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
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSchedeSorveglianzaFAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDataset As Recordset
Dim PazienteKey As Integer
'Dim ColRosso As Long
'Dim ColNero As Long
'ColRosso = &HFF&
'ColNero = &H80000012
Dim keyId As Integer
Dim modifica As Boolean

Private Sub cmdChiudi_Click()
    Unload frmSchedeSorveglianzaFAV
End Sub

Private Sub cmdMemorizza_Click()
    Dim v_Val() As Variant
    Dim v_Nomi() As Variant

    
    If Not modifica Then
        keyId = GetNumero("SCHEDA_SORV_FAV")
    End If
             
    v_Nomi = Array("KEY", "KEY_PAZIENTE", "ERITEMA")
    
    v_Val = Array(keyId, PazienteKey, GestisciOptEritema)
        
    Set rsDataset = New Recordset
        If modifica = False Then
            rsDataset.Open "SCHEDA_SORV_FAV", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.AddNew v_Nomi, v_Val
            rsDataset.Update
            modifica = True
        Else
            rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
            rsDataset.Update v_Nomi, v_Val
        End If
    Set rsDataset = Nothing

    MsgBox "Salvataggio effettuato", vbInformation, "Salvataggio"
    
End Sub

Private Function GestisciOptEritema() As String
    If optNoEritema.Value = True Then
        GestisciOptEritema = "NO"
    ElseIf optSiEritema.Value = True Then
        GestisciOptEritema = "SI"
    ElseIf optEritemaLieve.Value = True Then
        GestisciOptEritema = "LIEVE"
    ElseIf optEritemaMeno.Value = True Then
        GestisciOptEritema = "MENO"
    ElseIf optEritemaGrave.Value = True Then
        GestisciOptEritema = "GRAVE"
    End If
End Function

Private Sub Form_Activate()
    If Not RidisponiForms(Me) Then Exit Sub
    
    If PazienteKey = 0 Then
        cmdTrova_Click
        If tTrova.keyReturn = 0 Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdTrova_Click()
    ' pulisce per evitare problemi
    tTrova.Tipo = tpPAZIENTE
    tTrova.condizione = ""
    tTrova.condStato = ""
    frmTrova.Show 1
    If tTrova.keyReturn <> -1 Then
        If PazienteKey = tTrova.keyReturn Then
            PazienteKey = 0
            Call CaricaPaziente
            PazienteKey = tTrova.keyReturn
            Call CaricaPaziente
        Else
            PazienteKey = tTrova.keyReturn
            Call CaricaPaziente
        End If
    End If
End Sub

Private Sub CaricaPaziente()
    
    If PazienteKey = 0 Then
    
    Else
        ' carica i dati del paziente
        Set rsDataset = New Recordset
        rsDataset.Open "SELECT COGNOME,NOME,DATA_NASCITA FROM PAZIENTI WHERE KEY=" & PazienteKey, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        lblCognome = rsDataset("COGNOME")
        lblNome = rsDataset("NOME")
        Dim somma As Integer
        If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
            somma = -1
        ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
            somma = -1
        Else
            somma = 0
        End If
        lblEta = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
        Set rsDataset = Nothing
       
        ' cerca i riferimenti al paziente
        Call CaricaValori
    
    End If
End Sub

Private Sub CaricaValori()
    
    Set rsDataset = New Recordset
    
    rsDataset.Open "SELECT * FROM SCHEDA_SORV_FAV WHERE KEY_PAZIENTE=" & PazienteKey, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        Call CaricaOptEritema
        'da inserire qui
        modifica = True
    Else
        'Se non trova il paziente associato vuol dire che è in fase di inserimento
        modifica = False
    End If
    
    Set rsDataset = Nothing
    
End Sub

Private Sub CaricaOptEritema()
    If rsDataset("ERITEMA") = "NO" Then
        optNoEritema.Value = True
    ElseIf rsDataset("ERITEMA") = "SI" Then
        optSiEritema.Value = True
    ElseIf rsDataset("ERITEMA") = "LIEVE" Then
        optEritemaLieve.Value = True
    ElseIf rsDataset("ERITEMA") = "MENO" Then
        optEritemaMeno.Value = True
    ElseIf rsDataset("ERITEMA") = "GRAVE" Then
        optEritemaGrave.Value = True
    End If
End Sub

Private Sub optNoEritema_GotFocus()
    optEritemaLieve.Enabled = False
    optEritemaMeno.Enabled = False
    optEritemaGrave.Enabled = False
End Sub

Private Sub optSiEritema_GotFocus()
    optEritemaLieve.Enabled = True
    optEritemaMeno.Enabled = True
    optEritemaGrave.Enabled = True
End Sub


