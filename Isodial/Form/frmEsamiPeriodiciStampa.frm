VERSION 5.00
Begin VB.Form frmEsamiPeriodiciStampa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Esami Periodici"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStampa 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton optStampaPrescrizioniTuttiPazienti 
         Caption         =   "Stampa prescrizione per tutti i pazienti"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   4455
      End
      Begin VB.OptionButton optStampaPrescrizioni 
         Caption         =   "Stampa prescrizione"
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
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton optStampaStandard 
         Caption         =   "Stampa standard"
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
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipologia di stampa:"
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
         TabIndex        =   14
         Top             =   240
         Width           =   2160
      End
   End
   Begin VB.Frame fraFrequenza 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   4815
      Begin VB.OptionButton optStampaDicituraImpostata 
         Caption         =   "Stampa dicitura impostata"
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
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   3495
      End
      Begin VB.OptionButton optStampaDicitura 
         Caption         =   "Stampa dicitura"
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox chkFrequenzaSeProblemiClinici 
         Caption         =   "Se problemi clinici"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkFrequenzaAnnuale 
         Caption         =   "Annuale"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkFrequenzaSemestrale 
         Caption         =   "Semestrale"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkFrequenzaTrimestrale 
         Caption         =   "Trimestrale"
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkFrequenzaMensile 
         Caption         =   "Mensile"
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
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkFrequenzaBimestrale 
         Caption         =   "Bimestrale"
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frequenza:"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   4815
      Begin VB.CommandButton cmdImpostaDicitura 
         Caption         =   "&Imposta dicitura"
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1860
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Stampa"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdAnnulla 
         Cancel          =   -1  'True
         Caption         =   "&Annulla"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmEsamiPeriodiciStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnStampa As Boolean
Public intPeriodo As tipoPeriodo
Public intTipoStampa As Integer
Public blnDicitura As Boolean
Public blnStampaDicituraImpostata As Boolean


Private Sub AbilitaFrequenza(inStato)
    chkFrequenzaAnnuale.Enabled = inStato
    chkFrequenzaBimestrale.Enabled = inStato
    chkFrequenzaMensile.Enabled = inStato
    chkFrequenzaSemestrale.Enabled = inStato
    chkFrequenzaSeProblemiClinici.Enabled = inStato
    chkFrequenzaTrimestrale.Enabled = inStato
    optStampaDicitura.Enabled = inStato
    optStampaDicituraImpostata.Enabled = inStato
End Sub

Private Sub cmdAnnulla_Click()
    blnStampa = False
    Unload Me
End Sub

Private Sub cmdImpostaDicitura_Click()
    frmImpostaDicitura.Show 1
End Sub

Private Sub cmdStampa_Click()
    If optStampaStandard.Value Or _
        chkFrequenzaAnnuale.Value = Checked Or _
        chkFrequenzaBimestrale.Value = Checked Or _
        chkFrequenzaMensile.Value = Checked Or _
        chkFrequenzaSemestrale.Value = Checked Or _
        chkFrequenzaSeProblemiClinici.Value = Checked Or _
        chkFrequenzaTrimestrale.Value = Checked Then
        
        If chkFrequenzaAnnuale.Value = Checked Then intPeriodo = tpANNUALE
        If chkFrequenzaBimestrale.Value = Checked Then intPeriodo = tpBIMESTRALE
        If chkFrequenzaMensile.Value = Checked Then intPeriodo = tpMENSILE
        If chkFrequenzaSemestrale.Value = Checked Then intPeriodo = tpSEMESTRALE
        If chkFrequenzaSeProblemiClinici.Value = Checked Then intPeriodo = tpPROBLEMI
        If chkFrequenzaTrimestrale.Value = Checked Then intPeriodo = tpTRIMESTRALE
        If optStampaPrescrizioni.Value Then intTipoStampa = 2
        If optStampaPrescrizioniTuttiPazienti.Value Then intTipoStampa = 3
        If optStampaStandard.Value Then intTipoStampa = 1
        If optStampaDicitura.Value = True Then
            blnDicitura = True
        Else
            blnDicitura = False
        End If
        If optStampaDicituraImpostata.Value = True Then
            blnStampaDicituraImpostata = True
        Else
            blnStampaDicituraImpostata = False
        End If
        blnStampa = True
        Unload Me
    Else
        MsgBox "Selezionare almeno un periodo", vbExclamation, Me.Caption
    End If
End Sub

Private Sub optStampaPrescrizioniTuttiPazienti_Click()
    Call AbilitaFrequenza(optStampaPrescrizioniTuttiPazienti.Value)
End Sub

Private Sub optStampaPrescrizioni_Click()
    Call AbilitaFrequenza(optStampaPrescrizioni.Value)
End Sub

Private Sub optStampaStandard_Click()
    Call AbilitaFrequenza(Not optStampaStandard.Value)
End Sub

