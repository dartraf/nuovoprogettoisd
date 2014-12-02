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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483626
      ForeColor       =   8388608
      TabCaption(0)   =   "Versione 3.6.6"
      TabPicture(0)   =   "frmNote.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label19"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Versione 3.6.5"
      TabPicture(1)   =   "frmNote.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10(0)"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label10(1)"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "Label8(0)"
      Tab(1).Control(6)=   "Label4(3)"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(9)=   "Label4(5)"
      Tab(1).Control(10)=   "Label16"
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(12)=   "Label18"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Versione 3.6.4"
      TabPicture(2)   =   "frmNote.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(1)=   "Label10(6)"
      Tab(2).Control(2)=   "Label27"
      Tab(2).Control(3)=   "Label8(1)"
      Tab(2).Control(4)=   "Label26"
      Tab(2).Control(5)=   "Label25"
      Tab(2).Control(6)=   "Label4(2)"
      Tab(2).Control(7)=   "Label24"
      Tab(2).Control(8)=   "Label23"
      Tab(2).Control(9)=   "Label22"
      Tab(2).Control(10)=   "Label21"
      Tab(2).Control(11)=   "Label10(5)"
      Tab(2).Control(12)=   "Label20"
      Tab(2).Control(13)=   "Label10(4)"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Versione 3.6.3"
      TabPicture(3)   =   "frmNote.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label44"
      Tab(3).Control(1)=   "Label10(14)"
      Tab(3).Control(2)=   "Label43"
      Tab(3).Control(3)=   "Label8(3)"
      Tab(3).Control(4)=   "Label42"
      Tab(3).Control(5)=   "Label41"
      Tab(3).Control(6)=   "Label4(4)"
      Tab(3).Control(7)=   "Label40"
      Tab(3).Control(8)=   "Label39"
      Tab(3).Control(9)=   "Label38"
      Tab(3).Control(10)=   "Label37"
      Tab(3).Control(11)=   "Label10(13)"
      Tab(3).Control(12)=   "Label36"
      Tab(3).Control(13)=   "Label10(12)"
      Tab(3).Control(14)=   "Label35"
      Tab(3).Control(15)=   "Label34"
      Tab(3).Control(16)=   "Label33"
      Tab(3).Control(17)=   "Label10(11)"
      Tab(3).Control(18)=   "Label10(10)"
      Tab(3).ControlCount=   19
      Begin VB.Label Label19 
         Caption         =   "Implementato il salto dei campi con il tasto INVIO"
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
         TabIndex        =   58
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label Label18 
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label17 
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
         Left            =   -74880
         TabIndex        =   56
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label16 
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
         Left            =   -74880
         TabIndex        =   55
         Top             =   1200
         Width           =   2775
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   54
         Top             =   1440
         Width           =   7095
      End
      Begin VB.Label Label15 
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
         Left            =   -74880
         TabIndex        =   53
         Top             =   2400
         Width           =   7095
      End
      Begin VB.Label Label14 
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
         Left            =   -74880
         TabIndex        =   52
         Top             =   2160
         Width           =   2775
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
         Index           =   3
         Left            =   -74880
         TabIndex        =   51
         Top             =   1680
         Width           =   7095
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
         Index           =   0
         Left            =   -74880
         TabIndex        =   50
         Top             =   3315
         Width           =   7095
      End
      Begin VB.Label Label13 
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
         Left            =   -74880
         TabIndex        =   49
         Top             =   3075
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Inserito nel pannello con status ""Deceduto"" la specifica della causa di morte del paziente secondo la tabella E.D.T.A."
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   48
         Top             =   4155
         Width           =   7095
      End
      Begin VB.Label Label12 
         Caption         =   "-ANAGRAFIC A PAZIENTI-"
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
         TabIndex        =   47
         Top             =   3915
         Width           =   4335
      End
      Begin VB.Label Label11 
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
         Left            =   -74880
         TabIndex        =   46
         Top             =   4755
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
         Index           =   0
         Left            =   -74880
         TabIndex        =   45
         Top             =   4995
         Width           =   7095
      End
      Begin VB.Label Label28 
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
         TabIndex        =   44
         Top             =   3120
         Width           =   4335
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
         Index           =   6
         Left            =   -74880
         TabIndex        =   43
         Top             =   3360
         Width           =   7095
      End
      Begin VB.Label Label27 
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
         TabIndex        =   42
         Top             =   2520
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   41
         Top             =   2760
         Width           =   7455
      End
      Begin VB.Label Label26 
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
         TabIndex        =   40
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label25 
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
         TabIndex        =   39
         Top             =   2160
         Width           =   7095
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
         Index           =   2
         Left            =   -74880
         TabIndex        =   38
         Top             =   1560
         Width           =   7095
      End
      Begin VB.Label Label24 
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
         TabIndex        =   37
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label23 
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
         TabIndex        =   36
         Top             =   765
         Width           =   7455
      End
      Begin VB.Label Label22 
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
         TabIndex        =   35
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label21 
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
         TabIndex        =   34
         Top             =   3720
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   33
         Top             =   3960
         Width           =   7095
      End
      Begin VB.Label Label20 
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
         TabIndex        =   32
         Top             =   4320
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
         Index           =   4
         Left            =   -74880
         TabIndex        =   31
         Top             =   4560
         Width           =   7095
      End
      Begin VB.Label Label44 
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
         TabIndex        =   30
         Top             =   3570
         Width           =   4335
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
         Index           =   14
         Left            =   -74880
         TabIndex        =   29
         Top             =   3780
         Width           =   7095
      End
      Begin VB.Label Label43 
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
         Top             =   2565
         Width           =   2775
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
         Index           =   3
         Left            =   -74880
         TabIndex        =   27
         Top             =   2805
         Width           =   7455
      End
      Begin VB.Label Label42 
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
         TabIndex        =   26
         Top             =   1755
         Width           =   3495
      End
      Begin VB.Label Label41 
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
         TabIndex        =   25
         Top             =   1995
         Width           =   7095
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
         Index           =   4
         Left            =   -74880
         TabIndex        =   24
         Top             =   1275
         Width           =   7455
      End
      Begin VB.Label Label40 
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
         TabIndex        =   23
         Top             =   1035
         Width           =   3135
      End
      Begin VB.Label Label39 
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
         TabIndex        =   22
         Top             =   750
         Width           =   7455
      End
      Begin VB.Label Label38 
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
         TabIndex        =   21
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label37 
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
         TabIndex        =   20
         Top             =   4080
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
         Index           =   13
         Left            =   -74880
         TabIndex        =   19
         Top             =   4320
         Width           =   7095
      End
      Begin VB.Label Label36 
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
         TabIndex        =   18
         Top             =   4635
         Width           =   4335
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   17
         Top             =   4860
         Width           =   7095
      End
      Begin VB.Label Label35 
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
         TabIndex        =   16
         Top             =   2235
         Width           =   7095
      End
      Begin VB.Label Label34 
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
         TabIndex        =   15
         Top             =   3255
         Width           =   7095
      End
      Begin VB.Label Label33 
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
         TabIndex        =   14
         Top             =   5175
         Width           =   4335
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
         Index           =   11
         Left            =   -74880
         TabIndex        =   13
         Top             =   5400
         Width           =   7455
      End
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   12
         Top             =   3045
         Width           =   7095
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
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Width           =   7095
      End
      Begin VB.Label Label32 
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
         TabIndex        =   10
         Top             =   4800
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "-ANAGRAFIC A PAZIENTI-"
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
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Label10 
         Caption         =   "Inserito nel pannello con status ""Deceduto"" la specifica della causa di morte del paziente secondo la tabella E.D.T.A."
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
         Top             =   4200
         Width           =   7095
      End
      Begin VB.Label Label7 
         Caption         =   "-FATTURAZIONE - FILE C-"
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
         TabIndex        =   7
         Top             =   2540
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Aggiunti i file C in formato TXT"
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
         TabIndex        =   6
         Top             =   2780
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "-ESAMI DI LABORATORIO-REGISTRAZIONE"
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
         TabIndex        =   5
         Top             =   1725
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "Predisposto un pulsante per sostituire in caso di errori la data di registrazione degli esami"
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
         TabIndex        =   4
         Top             =   1980
         Width           =   7095
      End
      Begin VB.Label Label4 
         Caption         =   "Implementate le descrizioni dei comandi - Per attivarle posizionare il mouse sul pulsante"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   7455
      End
      Begin VB.Label Label3 
         Caption         =   "-TUTTE LE FINESTRE-"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "-SCHEDE DIALITICHE GIORNALIERE e SEDUTE STRAORDINARIE-REGISTRAZIONE-"
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
         Left            =   110
         TabIndex        =   1
         Top             =   480
         Width           =   7540
      End
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
