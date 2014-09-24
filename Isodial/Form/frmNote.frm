VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note di Rilascio"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "frmNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabNote 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483626
      ForeColor       =   8388608
      TabCaption(0)   =   "Versione 3.6.5"
      TabPicture(0)   =   "frmNote.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Versione 3.6.4"
      TabPicture(1)   =   "frmNote.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10(3)"
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(2)=   "Label10(1)"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(5)=   "Label16"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label4(3)"
      Tab(1).Control(8)=   "Label14"
      Tab(1).Control(9)=   "Label13"
      Tab(1).Control(10)=   "Label8(0)"
      Tab(1).Control(11)=   "Label12"
      Tab(1).Control(12)=   "Label10(0)"
      Tab(1).Control(13)=   "Label11"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Versione 3.6.3"
      TabPicture(2)   =   "frmNote.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10(8)"
      Tab(2).Control(1)=   "Label10(7)"
      Tab(2).Control(2)=   "Label31"
      Tab(2).Control(3)=   "Label30"
      Tab(2).Control(4)=   "Label29"
      Tab(2).Control(5)=   "Label10(6)"
      Tab(2).Control(6)=   "Label28"
      Tab(2).Control(7)=   "Label10(5)"
      Tab(2).Control(8)=   "Label27"
      Tab(2).Control(9)=   "Label26"
      Tab(2).Control(10)=   "Label25"
      Tab(2).Control(11)=   "Label24"
      Tab(2).Control(12)=   "Label4(2)"
      Tab(2).Control(13)=   "Label23"
      Tab(2).Control(14)=   "Label22"
      Tab(2).Control(15)=   "Label8(1)"
      Tab(2).Control(16)=   "Label21"
      Tab(2).Control(17)=   "Label10(4)"
      Tab(2).Control(18)=   "Label20"
      Tab(2).ControlCount=   19
      Begin VB.Label Label10 
         Caption         =   "Ampliate le voci del campo ACCESSO VASCOLARE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   44
         Top             =   3080
         Width           =   7095
      End
      Begin VB.Label Label10 
         Caption         =   "Predisposto un pulsante ""MARKERS VIRALI TRIMESTRALI"" per la pre-assegnazione del nome del documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   -74880
         TabIndex        =   43
         Top             =   5445
         Width           =   7455
      End
      Begin VB.Label Label31 
         Caption         =   "-SCANSIONE DOCUMENTI PAZIENTI-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   42
         Top             =   5220
         Width           =   4335
      End
      Begin VB.Label Label30 
         Caption         =   "Aggiunta l'unità di misura ""lt"" del valore del campo SOLUZIONE INFUSIONALE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   3300
         Width           =   7095
      End
      Begin VB.Label Label29 
         Caption         =   "Aggiunta la ricerca per le iniziali del nome dell'esame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   2280
         Width           =   7095
      End
      Begin VB.Label Label10 
         Caption         =   "Aggiunto il campo ALTEZZA (cm.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   39
         Top             =   4905
         Width           =   7095
      End
      Begin VB.Label Label28 
         Caption         =   "-ANAGRAFICA PAZIENTI-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   4680
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "Aggiunta la media totale annuale calcolata su tutti i pazienti in base allo stato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   37
         Top             =   4365
         Width           =   7095
      End
      Begin VB.Label Label27 
         Caption         =   "-STAMPE ANNUALI-KTV-TSAT%-CA/P-PTH-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   4125
         Width           =   4335
      End
      Begin VB.Label Label26 
         Caption         =   "-AVVIO ISODIAL-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   520
         Width           =   2775
      End
      Begin VB.Label Label25 
         Caption         =   "Attivato il controllo sulla data di sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   795
         Width           =   7455
      End
      Begin VB.Label Label24 
         Caption         =   "-SCHEDA SORVEGLIANZA FAV-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Implementata la gestione e la stampa della scheda FAV-Il modulo e' opzionale ed è attivabile a richiesta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   -74880
         TabIndex        =   32
         Top             =   1320
         Width           =   7455
      End
      Begin VB.Label Label23 
         Caption         =   "Aggiunta la colonna che visualizza il gruppo al quale l'esame è associato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   2040
         Width           =   7095
      End
      Begin VB.Label Label22 
         Caption         =   "-TABELLA ESAMI DI LABORATORIO-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "Aggiunto il campo RITMO DIALISI SETTIMANALE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   29
         Top             =   2840
         Width           =   7455
      End
      Begin VB.Label Label21 
         Caption         =   "-ANAMNESI DIALITICA-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   2600
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Abilitato il tasto INVIO per velocizzare l'inserimento dei valori"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   27
         Top             =   3825
         Width           =   7095
      End
      Begin VB.Label Label20 
         Caption         =   "-REGISTRAZIONE ESAMI DI LABORATORIO-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   3615
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "Abilitata la tabella dei COMUNI all'utente MEDICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   25
         Top             =   4600
         Width           =   7095
      End
      Begin VB.Label Label19 
         Caption         =   "-TABELLE-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   4360
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "Eliminato definitivamente il mancato allineamento delle voci"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   23
         Top             =   4000
         Width           =   7095
      End
      Begin VB.Label Label18 
         Caption         =   "-STAMPE ETICHETTE-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   3760
         Width           =   4335
      End
      Begin VB.Label Label17 
         Caption         =   "-TERAPIA DIALITICA-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   520
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Aggiunti 3 campi DATA per la prescrizione di farmaci a somministrazione non settimanale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   20
         Top             =   800
         Width           =   7455
      End
      Begin VB.Label Label15 
         Caption         =   "-ANAMNESI DIALITICA-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1365
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Modificato il campo RITMO DIALISI SETTIMANALE (prescrizione)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   18
         Top             =   1605
         Width           =   7095
      End
      Begin VB.Label Label14 
         Caption         =   "Implementata la stampa della vista CORRENTE dell'esame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   2205
         Width           =   7095
      End
      Begin VB.Label Label13 
         Caption         =   "-ESAMI STRUMENTALI-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1965
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Abilitata l'eliminazione degli apparati con dati relazionati alle altre gestioni del software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   15
         Top             =   2800
         Width           =   7455
      End
      Begin VB.Label Label12 
         Caption         =   "-GESTIONE APPARATI-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   2560
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Implementata la stampa dei grafici 2D/3D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   13
         Top             =   3400
         Width           =   7095
      End
      Begin VB.Label Label11 
         Caption         =   "-GESTIONE INDICATORI - Kt/V- TSAT% - Ca/P -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   3160
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "-PREDISPOSIZIONE FATTURA ELETTRONICA-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "Implementate le tabelle per la gestione e generazione delle fatture elettroniche"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   7095
      End
      Begin VB.Label Label7 
         Caption         =   "-E.D.T.A. CAUSA MORTE-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Implementata la tabella E.D.T.A. con i codici e le specifiche della causa della morte del paziente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   7095
      End
      Begin VB.Label Label4 
         Caption         =   "Eliminate le specifiche ""lieve-medio-grave"" dal campo PRESENZA FREMITI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1720
         Width           =   7095
      End
      Begin VB.Label Label5 
         Caption         =   "-INTESTAZIONE CENTRO-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2200
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Aggiunto il campo SITO WEB per riportare in tutte le stampe, qualora presente, l'indirizzo del sito internet del centro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2440
         Width           =   7095
      End
      Begin VB.Label Label4 
         Caption         =   "Aggiunto il campo ""INTERVENTI SULL'ACCESSO VASCOLARE"" con relativa data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1480
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "-SCHEDA RILEVAZIONE FAV-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Implementata la stampa delle previsioni"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   760
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "-CONSUMI e PREVISIONI-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   520
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
