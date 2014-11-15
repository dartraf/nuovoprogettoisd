VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5B6D0C10-C25A-4015-8142-215041993551}#4.0#0"; "ACPRibbon.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Centro Dialisi"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -5310
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":030A
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRibTab 
      Align           =   1  'Align Top
      Height          =   1920
      Left            =   0
      ScaleHeight     =   1860
      ScaleWidth      =   15180
      TabIndex        =   20
      Top             =   855
      Visible         =   0   'False
      Width           =   15240
      Begin Progetto1.ACPRibbon ribTab 
         Height          =   1740
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3069
         BackColor       =   -2147483636
         ForeColor       =   -2147483630
      End
   End
   Begin MSComctlLib.ImageList imgListRibbonTab 
      Left            =   4440
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A34F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AC27
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B500
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picContenitore 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   795
         Index           =   17
         Left            =   14250
         Picture         =   "frmMain.frx":BDC7
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Chiudi Tutte le Finestre Aperte"
         Top             =   0
         Width           =   770
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   16
         Left            =   13355
         Picture         =   "frmMain.frx":C3BD
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Genera File XML"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   15
         Left            =   12555
         Picture         =   "frmMain.frx":C851
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Genera File C"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   14
         Left            =   11760
         Picture         =   "frmMain.frx":CC9C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gestione Ricette"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   13
         Left            =   10860
         Picture         =   "frmMain.frx":D5CC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Piano di Lavoro"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   12
         Left            =   10060
         Picture         =   "frmMain.frx":DE18
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Scheda Dialitica - Compilazione"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   11
         Left            =   9260
         Picture         =   "frmMain.frx":E5BA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Turni e Reni"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   10
         Left            =   8345
         Picture         =   "frmMain.frx":EF68
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Diario Clinico"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   9
         Left            =   7550
         Picture         =   "frmMain.frx":F8D8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Accessi Vascolari"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   8
         Left            =   5865
         Picture         =   "frmMain.frx":FE9B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Terapia Dialitica"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   7
         Left            =   6660
         Picture         =   "frmMain.frx":1073C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Terapia Domiciliare"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   6
         Left            =   4990
         Picture         =   "frmMain.frx":10F28
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Esami di Laboratorio - Consultazione"
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   5
         Left            =   4180
         Picture         =   "frmMain.frx":11815
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Esami di Laboratorio - Registrazione"
         Top             =   0
         Width           =   810
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   4
         Left            =   3350
         Picture         =   "frmMain.frx":120C1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Esami Strumentali"
         Top             =   0
         Width           =   840
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   3
         Left            =   2495
         Picture         =   "frmMain.frx":12AA6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Anamnesi Dialitica"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   2
         Left            =   1690
         Picture         =   "frmMain.frx":1336E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anamnesi Nefrologica"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   1
         Left            =   890
         Picture         =   "frmMain.frx":13C37
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anamnesi Patologica"
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00C0C0C0&
         Height          =   800
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":144EE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Anagrafica Generale"
         Top             =   0
         Width           =   800
      End
   End
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7815
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4419
            Text            =   "ISODIAL 2.9"
            TextSave        =   "ISODIAL 2.9"
            Object.ToolTipText     =   "Versione in Uso"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5821
            MinWidth        =   5821
            Text            =   "Boscoreale"
            TextSave        =   "Boscoreale"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4940
            MinWidth        =   4940
            Object.ToolTipText     =   "Utente Connesso"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2999
            MinWidth        =   2999
            TextSave        =   "15/11/2014"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPaziente 
      Caption         =   "&Gestione Pazienti"
      Tag             =   "&Gestione pazienti|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Anagrafica Generale"
         Index           =   1
         Tag             =   "&Informazioni generali|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Anamnesi"
         Index           =   2
         Tag             =   "&Anamnesi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoPazAne 
            Caption         =   "&Patologica Remota e Familiare"
            Index           =   1
            Tag             =   "&Patologica remota e familiare|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoPazAne 
            Caption         =   "&Nefrologica"
            Index           =   2
            Tag             =   "&Nefrologica|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoPazAne 
            Caption         =   "&Dialitica"
            Index           =   3
            Tag             =   "&Scheda dialitica|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Esami"
         Index           =   3
         Tag             =   "&Esami|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoPazEsami 
            Caption         =   "&Strumentali"
            Index           =   1
            Tag             =   "&Strumentali|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoPazEsami 
            Caption         =   "&Laboratorio"
            Index           =   2
            Tag             =   "&Laboratorio|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
            Begin VB.Menu mnuSottoPazEsamiLab 
               Caption         =   "&Registrazione"
               Index           =   1
               Tag             =   "&Registrazione|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
            End
            Begin VB.Menu mnuSottoPazEsamiLab 
               Caption         =   "&Consultazione"
               Index           =   2
               Tag             =   "&Consultazione|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
            End
            Begin VB.Menu mnuSottoPazEsamiLab 
               Caption         =   "&Prescrizione"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Terapia"
         Index           =   4
         Tag             =   "&Terapia|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoPazTerapia 
            Caption         =   "&In dialisi"
            Index           =   1
            Tag             =   "&Dialitica|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoPazTerapia 
            Caption         =   "&Domiciliare"
            Index           =   2
            Tag             =   "&Domiciliare|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoPazTerapia 
            Caption         =   "Stampa Riepiloghi"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Accessi Vascolari"
         Index           =   5
         Tag             =   "&Accessi vascolari|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Diario Clinico"
         Index           =   6
         Tag             =   "&Diario clinico|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoPaz 
         Caption         =   "&Scansione Documenti Pazienti"
         Index           =   7
      End
   End
   Begin VB.Menu mnuDialisi 
      Caption         =   "Gestione &Emodialisi"
      Tag             =   "Gestione &Emodialisi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuSottoDialisi 
         Caption         =   "&Turni Pazienti"
         Index           =   1
         Tag             =   "&Turni e Reni|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoDialisiTurni 
            Caption         =   "&Associa Turni/Reni"
            Index           =   1
         End
         Begin VB.Menu mnuSottoDialisiTurni 
            Caption         =   "&Stampa Turni"
            Index           =   2
         End
      End
      Begin VB.Menu mnuSottoDialisi 
         Caption         =   "&Seduta Dialitica Giornaliera"
         Index           =   2
         Tag             =   "&Scheda dialitica giornaliera|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoDialisiScheda 
            Caption         =   "&Compilazione"
            Index           =   1
            Tag             =   "&Compilazione|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoDialisiScheda 
            Caption         =   "C&onsultazione"
            Index           =   2
            Tag             =   "C&onsultazione|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnuSottoDialisi 
         Caption         =   "S&eduta Supplementare"
         Index           =   3
         Tag             =   "S&eduta straordinaria|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoDialisi 
         Caption         =   "&Piano di Lavoro"
         Index           =   4
         Tag             =   "Stam&pa Piano di Lavoro|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoDialisi 
         Caption         =   "C&onsumi e Previsioni"
         Index           =   5
      End
   End
   Begin VB.Menu mnuGestioneIndicatori 
      Caption         =   "Gestione &Indicatori"
      Tag             =   "Gestione &Indicatori|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Calcolo Kt/V"
         Index           =   1
         Tag             =   "&Indicatore Adeguatezza Dialitica (Kt/V)|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "C&alcolo TSAT %"
         Index           =   2
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "Calcolo Pr&odotto Ca / P"
         Index           =   3
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Eventi"
         Index           =   4
         Tag             =   "&Eventi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Colture"
         Index           =   5
         Tag             =   "&Colture|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "E&ritropoietina per Paziente"
         Index           =   6
         Tag             =   "E&ritropoietina per Paziente|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Monitoraggi"
         Index           =   7
         Tag             =   "&Monitoraggi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Trattamento Acque"
         Index           =   8
         Tag             =   "&Trattamento acque|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Pazienti Candidati ai Trapianti"
         Index           =   9
         Tag             =   "&Pazienti candidati al trapianto|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "Esami &Periodici"
         Index           =   10
         Tag             =   "Esami &Periodici|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuGestioneIndicatoriSotto 
         Caption         =   "&Scheda Rilevazione FAV"
         Index           =   11
      End
   End
   Begin VB.Menu mnuArchivi 
      Caption         =   "&Setup Tabelle"
      Tag             =   "&Setup Tabelle|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Organigramma"
         Index           =   1
         Tag             =   "&Organigramma|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnuSottoTabOrgan 
            Caption         =   "&Direttore Sanitario"
            Index           =   1
         End
         Begin VB.Menu mnuSottoTabOrgan 
            Caption         =   "&Medici in Dialisi"
            Index           =   2
            Tag             =   "&Medici dialisi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mnuSottoTabOrgan 
            Caption         =   "&Infermieri"
            Index           =   3
            Tag             =   "&Infermieri|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Medici &Refertanti"
         Index           =   2
         Tag             =   "Medici &refertanti|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Medici di &Base"
         Index           =   3
         Tag             =   "Medici di &base|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Psicologi"
         Index           =   4
         Tag             =   "&Psicologi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Or&gano/Apparato"
         Index           =   5
         Tag             =   "Or&gano/apparato|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Esami &Strumentali per Organo/Apparato"
         Index           =   6
         Tag             =   "Tipo di &esame|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Filtri"
         Index           =   7
         Tag             =   "Tipo di &filtro|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Linee"
         Index           =   8
         Tag             =   "Tipo di &linee|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Aghi"
         Index           =   9
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Far&maci in Uso"
         Index           =   10
         Tag             =   "&Medicinali in dialisi|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "A&nticoagulanti"
         Index           =   11
         Tag             =   "&Anticoagulanti|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Titoli Diario Clinico"
         Index           =   12
         Tag             =   "Titoli per &diario clinico|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "&Esami di Laboratorio"
         Index           =   13
         Tag             =   "&Voci per esami di lab.|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Raggr&uppamento Esami di Laboratorio"
         Index           =   14
         Tag             =   "Raggruppamento esami di &lab.|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Co&dici Era - EDTA"
         Index           =   15
         Tag             =   "E.&D.T.A.|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuSottoTab 
         Caption         =   "Codici E.D.T.A. - Causa &Morte"
         Index           =   16
      End
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "&Strumenti"
      Tag             =   "&Strumenti|(Checked=0)(Enabled=-1)(Visible=0)(WindowList=0)"
      Visible         =   0   'False
      Begin VB.Menu mnuGesPass 
         Caption         =   "&Gestione Utenti"
         Tag             =   "&Gestione password|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuImpostaStampa 
         Caption         =   "&Intestazione Centro"
      End
      Begin VB.Menu mnuImpostaBackup 
         Caption         =   "&N° Backup"
      End
      Begin VB.Menu mnuRipristina 
         Caption         =   "&Ripristino Archivi"
      End
      Begin VB.Menu mnuEsportaDb 
         Caption         =   "&Esporta Database"
      End
   End
   Begin VB.Menu mnuStampe 
      Caption         =   "Stam&pe"
      Tag             =   "Stam&pe|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnuStampaPaz 
         Caption         =   "Lista &Pazienti"
         Tag             =   "Lista pazienti in dialisi |(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuStampaMediciBase 
         Caption         =   "Lista &Medici di Base"
      End
      Begin VB.Menu mnuMostraFattElaborazione 
         Caption         =   "&Visualizza Giorni Dialisi"
      End
      Begin VB.Menu mnuImpegnativeDialisi 
         Caption         =   "Richieste &Impegnative Dialisi"
      End
      Begin VB.Menu mnuEtichettePerProvetta 
         Caption         =   "&Etichette per Provette"
      End
      Begin VB.Menu mnuModuloFirmePaziente 
         Caption         =   "&Modulo Firme Paziente"
      End
      Begin VB.Menu mnuKtvAnnuale 
         Caption         =   "&KT/V Annuale"
      End
      Begin VB.Menu mnuTsatAnnuale 
         Caption         =   "&TSAT% Annuale"
      End
      Begin VB.Menu mnuCaPAnnuale 
         Caption         =   "&Ca/P Annuale"
      End
      Begin VB.Menu mnuPthAnnuale 
         Caption         =   "&PTH Annuale"
      End
      Begin VB.Menu mnuSchedaDialiticaSettimanale 
         Caption         =   "&Scheda Dialitica Settimanale"
      End
   End
   Begin VB.Menu mnuFatturazione 
      Caption         =   "&Fatturazione"
      Tag             =   "&Fatturazione|(Checked=0)(Enabled=-1)(Visible=0)(WindowList=0)"
      Visible         =   0   'False
      Begin VB.Menu mnuIntestazioneFattura 
         Caption         =   "&Parametri Fattura"
      End
      Begin VB.Menu mnuTabelleFatturazione 
         Caption         =   "&Tabelle"
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&Regioni"
            Index           =   0
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&Comuni"
            Index           =   1
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&ASL"
            Index           =   2
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&Distretti"
            Index           =   3
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&Esenzioni"
            Index           =   4
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "&Nomenclatore Tariffario"
            Index           =   5
         End
         Begin VB.Menu mnuTabFatt 
            Caption         =   "Acc&ompagnatori"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCaricaPrescrizione 
         Caption         =   "&Gestione Ricette"
      End
      Begin VB.Menu mnuGestioneFileC 
         Caption         =   "&Genera File C"
      End
      Begin VB.Menu mnuGestioneFileXml 
         Caption         =   "&Genera File XML"
      End
      Begin VB.Menu mnuFattStampaFogli 
         Caption         =   "S&tampa Fogli di Viaggio"
         Tag             =   "S&tampa Fogli di Viaggio|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuRimborsi 
         Caption         =   "St&ampa Rimborsi"
      End
      Begin VB.Menu mnuStampaRiepiloghi 
         Caption         =   "&Stampa Riepiloghi"
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "&Fattura"
            Index           =   0
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Paziente"
            Index           =   1
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Totali-Prestazioni"
            Index           =   2
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x Totali - A&sl"
            Index           =   3
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Totali - Mazzette x Distretti"
            Index           =   4
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Mazzette - Mensili"
            Index           =   5
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Mazzetta - Singola"
            Index           =   6
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Asl - Distretti"
            Index           =   7
         End
         Begin VB.Menu mnuStampaRiepilogo 
            Caption         =   "x &Impegnative"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnuVassoio 
      Caption         =   "PopupVassoio"
      Tag             =   "PopupVassoio|(Checked=0)(Enabled=-1)(Visible=0)(WindowList=0)"
      Visible         =   0   'False
      Begin VB.Menu mnuApriVassoio 
         Caption         =   "Ripristina"
         Tag             =   "Ripristina|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu mnuPopUpEsamiPeriodici 
      Caption         =   "PopUpEsamiPeriodici"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpInserisciEsame 
         Caption         =   "&Inserisci Esame"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpStampaStandard 
         Caption         =   "Stampa Standard"
      End
      Begin VB.Menu mnuPopUpStampaPrescrizione 
         Caption         =   "&Stampa Prescrizione"
      End
   End
   Begin VB.Menu mnuApparati 
      Caption         =   "&Apparati"
      Begin VB.Menu mnuSottoApparati 
         Caption         =   "Gestione A&pparati"
         Index           =   0
      End
      Begin VB.Menu mnuSottoApparati 
         Caption         =   "Stampa Re&gistro"
         Index           =   1
      End
      Begin VB.Menu mnuSottoApparati 
         Caption         =   "Parco Ren&i Artificiali"
         Index           =   2
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "&?"
      Tag             =   "&?|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnunoterilascio 
         Caption         =   "&Note di Rilascio "
      End
      Begin VB.Menu mnulicenza 
         Caption         =   "&Licenza d'uso"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&Informazioni su"
         Tag             =   "&Informazioni su|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ret As Integer

Private Sub MDIForm_Load()
    '/ If gbSubClassMenu is False, the menu is not subclassed
    gbSubClassMenu = True

    If gbSubClassMenu Then SubClassMenuXP
    
    frmMain.staBar.Panels(2) = structIntestazione.sRagione
    frmMain.staBar.Panels(3) = structIntestazione.sCitta
    Me.Top = -60
    Me.Left = -60
    Me.Height = 11190
    Me.Width = 15480
    Me.staBar.Panels(1) = "ISODIAL " & App.Major & "." & App.Minor & "." & App.Revision
    
    If NUOVA_TOOLBAR Then
        Call CaricaRibbonTab
        picRibTab.Visible = True
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Me.WindowState = vbNormal
        gHW = Me.hWnd
        myNID.cbSize = Len(myNID)
        myNID.hWnd = gHW
        myNID.uID = uID
        myNID.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        myNID.uCallbackMessage = cbNotify
        myNID.hIcon = Me.Icon
        myNID.szTip = Trim(Me.Caption) & Chr(0)
        Shell_NotifyIcon NIM_ADD, myNID
        HOOK
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Cancel = True
    frmDisconnetti.Show 1
    If tDisconnetti = tpDCHIUDICONBACKUP Then
        If Not structApri.server Then
            ' esce dalla lista dei client collegati
            Dim rsDataset As New Recordset
            rsDataset.Open "CLIENT", cnPrinc, adOpenKeyset, adLockPessimistic, adCmdTable
            rsDataset.Update "NUMERO", rsDataset("NUMERO") - 1
            Set rsDataset = Nothing
            Set cnPrinc = Nothing
            Set cnTrac = Nothing
            tRete = tpDISCONNETTI
            frmAttendi.Show 1
        End If
        End
    ElseIf tDisconnetti = tpDANNULLA Then
        Call SubClassMenuXP
    End If
End Sub

Private Sub CaricaRibbonTab()
    On Error GoTo gestione
    
 '   Const TAB_PAZIENTI As Integer = 1
    
    ribTab.Width = picRibTab.Width
    ribTab.Theme = 1

    ' immagini 32x32
 '   Set ribTab.zImg = imgListRibbonTab
 '   ribTab.ButtonCenter = False
'ribTab.ImageList = imgListRibbonTab
    '# Add Tabs ---   ID - Caption
    ribTab.AddTab "1", "Gestione Pazienti"
    ribTab.AddTab "2", "Gestione Emodialisi"
    ribTab.AddTab "3", "Gestione Indicatori"
    ribTab.AddTab "4", "Setup Tabelle"
    ribTab.AddTab "5", "Strumenti"
    ribTab.AddTab "6", "Stampe"
    ribTab.AddTab "7", "Fatturazione"
    ribTab.AddTab "8", "Apparati"
    ribTab.AddTab "9", "?"
    
    '# Add Cats ---   ID - Tab - Caption - ShowDialogButton
    ribTab.AddCat "1", "1", "Anagrafica Generale", False
    ribTab.AddCat "2", "1", "Anamnesi", True
    ribTab.AddCat "3", "1", "Esami", True
    ribTab.AddCat "4", "1", "Terapia", True
    ribTab.AddCat "5", "1", "Accessi Vascolari", False
    ribTab.AddCat "6", "1", "Diario clinico", False
    ribTab.AddCat "7", "1", "Scansione Documenti Pazienti", False
    
    ribTab.AddCat "1", "2", "Turni Pazienti", True
    ribTab.AddCat "2", "2", "Seduta Dialitica Giornaliera", True
    ribTab.AddCat "3", "2", "Seduta Supplementare", False
    ribTab.AddCat "4", "2", "Piano di Lavoro", False
    ribTab.AddCat "5", "2", "Consumi e Previsioni", False
   
 '   ribTab.AddCat "11", "3", "Indicatori", False
 '   ribTab.AddCat "12", "3", "Scansioni", False
    
 '   ribTab.AddCat "14", "4", "Tabelle", False
    
    
    
    '# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
  '  ribTab.AddButton "0", "1", "Info", 1
  '  ribTab.AddButton "1", "2", "Dialitica", 2, False, "Anamnesi Dialitica"
  '  ribTab.AddButton "2", "2", "Nefrologica", 3, False, "Anamnesi Nefrologica"
  '   ribTab.AddButton "3", "2", "Nefrologica", 4, False, "Anamnesi Nefrologica"
   
    '# Repaint Ribbon
    ribTab.Refresh

    
    
    Exit Sub
gestione:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub mnuNoteRilascio_Click()
    frmNote.Show 1
End Sub

Private Sub mnuabout_Click()
   frmInfo.Show 1
End Sub

Private Sub mnulicenza_Click()
    On Error GoTo gestione
    Shell structApri.pathExe & "\SumatraPDF" & " " & structApri.pathExe & "\LicenzaISODIAL.pdf", vbMaximizedFocus
   
gestione:
    If Err.Number = 53 Then
       MsgBox "Si è verificato un errore aprendo il lettore PDF", vbCritical, "ATTENZIONE!!!"
    End If
End Sub

Private Sub mnuCaPAnnuale_Click()
    tStampa = tpCAPAnnuale
    frmStampaFiltri.Show 1
End Sub

Private Sub mnuPopUpInserisciEsame_Click()
    frmEsamiPeriodici.InserisciEsame_Click
End Sub

Private Sub mnuPopUpStampaPrescrizione_Click()
    frmEsamiPeriodici.StampaPrescrizione_Click
End Sub

Private Sub mnuPopUpStampaStandard_Click()
    frmEsamiPeriodici.StampaStandard_Click
End Sub

Private Sub mnuApriVassoio_Click()
    ' rende visibile il form
    frmMain.Visible = True
    Shell_NotifyIcon NIM_DELETE, myNID
    unHook
End Sub

'Private Sub mnuBarra_Click()
'    mnuBarra.Checked = Not mnuBarra.Checked
'    picContenitore.Visible = mnuBarra.Checked
'End Sub

Private Sub mnuCaricaPrescrizione_Click()
    frmPrescrizioni.Show
End Sub

Private Sub mnuEtichettePerProvetta_Click()
    frmEtichette.Show 1
End Sub

Private Sub mnuFattStampaFogli_Click()
    tStampa = tpFOGLIOVIAGGIO
    frmStampaFogliViaggio.Show 1
End Sub

Private Sub mnuGesPass_Click()
    frmGestisciPassword.Show 1
End Sub

Private Sub mnuGestioneFileC_Click()
    tFileRicette = tpFILEC
    frmGestioneFileRicette.Show
End Sub

Private Sub mnuGestioneFileXml_Click()
    tFileRicette = tpFILEXML
    frmGestioneFileRicette.Show
End Sub

Private Sub mnugestioneIndicatoriSotto_Click(Index As Integer)
    Select Case Index
        Case 1: frmKtv.Show
        Case 2: frmTsat.Show
        Case 3: frmProdottoCalcioFosforo.Show
        Case 4: frmEventi.Show
        Case 5: frmColture.Show
        Case 6: frmEpo.Show
        Case 7: frmMonitoraggio.Show
        Case 8: frmTrattamentoAcque.Show
        Case 9: frmTrapianti.Show
        Case 10: frmEsamiPeriodici.Show
        Case 11: frmSchedeSorveglianzaFAV.Show 1
    End Select
End Sub

Private Sub mnuImpegnativeDialisi_Click()
    tStampa = tpIMPEGNATIVE
    frmRichiestaImpegnativeDialisi.Show 1
End Sub

Private Sub mnuImpostaBackup_Click()
    frmBackup.Show
End Sub
Private Sub mnuEsportaDb_Click()
    ' zippa i db
    
    Dim MYUSER As ZIPUSERFUNCTIONS
    Dim retcode As Long

    MYUSER.DLLPrnt = Puntatore(AddressOf Stampa_messaggi_zip)
    MYUSER.DLLPASSWORD = 0&
    MYUSER.DLLCOMMENT = 0&
    MYUSER.DLLSERVICE = 0&
    retcode = ZpInit(MYUSER)

    Dim MYOPT As ZPOPT
    retcode = ZpSetOptions(MYOPT)

    Dim files As ZIPnames
    
    If MsgBox("I Databases verranno esportati sul desktop in un file compresso .zip - CONFERMI?", vbQuestion + vbYesNo + vbDefaultButton2, "ESPORTA DATABASES") = vbNo Then
        Exit Sub
    End If
    
    files.s(0) = (structApri.pathDB) & "\Centro.mdb"
    files.s(1) = (structApri.pathDB) & "\Connessioni.mdb"
    retcode = ZpArchive(2, Environ$("USERPROFILE") & "\Desktop\Db " & Left(structIntestazione.sRagione, 12) & " " & Day(date) & "_" & Month(date) & "_" & Year(date) & ".zip", files)
    
    MsgBox "Databases esportati correttamente", vbInformation, "Esporta DataBases"

End Sub

Private Sub mnuImpostaStampa_Click()
    frmIntestazioneCentro.Show 1
End Sub

Private Sub mnuIntestazioneFattura_Click()
    frmParametriFattura.Show 1
End Sub

Private Sub mnuKtvAnnuale_Click()
    tStampa = tpKTVANNUALE
    frmStampaFiltri.Show 1
End Sub

Private Sub mnuModuloFirmePaziente_Click()
    tStampa = tpMODULOFIRMEPAZIENTE
    frmStampaFogliViaggio.Show 1
End Sub

Private Sub mnuMostraFattElaborazione_Click()
    frmMostraElaborazioni.Show 1
End Sub

Private Sub mnuPthAnnuale_Click()
    tStampa = tpPTHAnnuale
    frmStampaFiltri.Show 1
End Sub

Private Sub mnuRimborsi_Click()
    frmRimborsiSpese.Show
End Sub

Private Sub mnuRipristina_Click()
    Dim lettera As String
    Dim numClient As Integer
    If VerificaDiscoRimovibile(lettera) = False Then
       MsgBox "Impossibile procedere al ripristino - CONNETTERE L'UNITA'", vbCritical, "UNITA' DI BACKUP NON PRESENTE"
    ElseIf nessunClient(numClient) = False Then
           If MsgBox("ATTENZIONE!!! Altri utenti sono connessi ad ISODIAL - Li disconnetto automaticamente?", vbQuestion + vbYesNo, "CONTROLLO UTENTI") = vbYes Then
             Call PulisciTabCLIENTI
             frmPeriferiche.Show 1
           Else
              MsgBox "Disconnettere TUTTI gli utenti e riavviare il ripristino", vbCritical, "RIPRISTINO ARCHIVIO"
           End If
    Else
        frmPeriferiche.Show 1
    End If
End Sub

Private Sub mnuSchedaDialiticaSettimanale_Click()
    tStampa = tpSCHEDADIALITICASETTIMANALE
    frmStampaFiltri.Show 1
End Sub

Private Sub mnuSottoApparati_Click(Index As Integer)
    Select Case Index
        Case 0: frmApparati.Show 1
        Case 1: Call StampaRegistroApparati
        Case 2: tTabelle = tpRENI
                Unload frmTabelle
                frmTabelle.Show
    End Select
End Sub

Private Sub StampaRegistroApparati()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(50) AS TIPO_APPARATO, " & _
                "       NEW adVarChar(50) AS MODELLO, " & _
                "       NEW adVarChar(50) AS MATRICOLA, " & _
                "       NEW adVarChar(50) AS PRODUTTORE, " & _
                "       NEW adVarChar(10) AS MODALITA_ACQUISIZIONE, " & _
                "       NEW adInteger AS PERIODO_AMMORTAMENTO, " & _
                "       NEW adVarChar(11) AS DATA_ACQUISIZIONE, " & _
                "       NEW adVarChar(11) AS DATA_COLLAUDO, " & _
                "       NEW adVarChar(11) AS DATA_DISMISSIONE, " & _
                "       NEW adVarChar(11) AS DATA_ROTTAMAZIONE, " & _
                "       NEW adVarChar(11) AS PROXREVSIC, " & _
                "       NEW adVarChar(11) AS PROXREVFUN "
                
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    
    rsDataset.Open "SELECT * FROM APPARATI ORDER BY NUMERO_INVENTARIO", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            Do While Not rsDataset.EOF
                .AddNew
                .Fields("TIPO_APPARATO") = rsDataset("TIPO_APPARATO")
                .Fields("MODELLO") = rsDataset("MODELLO")
                .Fields("MATRICOLA") = rsDataset("MATRICOLA")
                .Fields("PRODUTTORE") = rsDataset("PRODUTTORE")
                .Fields("MODALITA_ACQUISIZIONE") = rsDataset("MODALITA_ACQUISIZIONE")
                .Fields("PERIODO_AMMORTAMENTO") = rsDataset("PERIODO_AMMORTAMENTO")
                .Fields("DATA_ACQUISIZIONE") = rsDataset("DATA_ACQUISIZIONE") & ""
                .Fields("DATA_COLLAUDO") = rsDataset("DATA_COLLAUDO") & ""
                .Fields("DATA_DISMISSIONE") = rsDataset("DATA_DISMISSIONE") & ""
                .Fields("DATA_ROTTAMAZIONE") = rsDataset("DATA_ROTTAMAZIONE") & ""
                .Fields("PROXREVSIC") = rsDataset("PROXREVSIC") & ""
                .Fields("PROXREVFUN") = rsDataset("PROXREVFUN") & ""
                rsDataset.MoveNext
            Loop
        End With
    End If
        
    Set rsDataset = Nothing
    
    Set rptRegistroApparati.DataSource = rsMain
    rptRegistroApparati.Orientation = rptOrientLandscape
    rptRegistroApparati.TopMargin = 0
    rptRegistroApparati.BottomMargin = 1000
    rptRegistroApparati.RightMargin = 0
    rptRegistroApparati.LeftMargin = 0
    'rptRegistroApparati.Sections("Intestazione").Controls("lblElenco").Caption = TipoElenco
    rptRegistroApparati.PrintReport True, rptRangeAllPages

End Sub

Private Sub mnuSottoDialisi_Click(Index As Integer)
    Select Case Index
        Case 3: frmSchedaStraordinaria.Show
        Case 4: frmPianoLavoro.Show
        Case 5: frmConsumiPrevisioni.Show
    End Select
End Sub

Private Sub mnuSottoDialisiScheda_Click(Index As Integer)
    If Index = 1 Then
        frmSchedaDialitica.Show
    Else
        frmSchedaDialiticaPassate.Show
    End If
End Sub

Private Sub mnuSottoDialisiTurni_Click(Index As Integer)
    If Index = 1 Then
        frmTurni.Show
    Else
        Call StampaTurni
    End If
End Sub

Private Sub mnuSottoPaz_Click(Index As Integer)
    Select Case Index
        Case 1: frmPaziente.Show    ' info generali
        Case 5: frmAccessi.Show     ' accessi vascolari
        Case 6: frmDiario.Show      ' diario clinico
        Case 7: frmScanDocumenti.Show   ' scansione documenti paziente
    End Select
End Sub

Private Sub mnuSottoPazAne_Click(Index As Integer)
    Select Case Index
        Case 1: frmAnamnesiPat.Show     ' patologica remota e familiare
        Case 2: frmAnamnesiNefro.Show    ' nefrologica
        Case 3: frmAnamnesiDialitica.Show  ' dialitica
    End Select
End Sub

Private Sub mnuSottoPazEsami_Click(Index As Integer)
    If Index = 1 Then
        frmEsamiStrumentali.Show
    End If
End Sub

Private Sub mnuSottoPazEsamiLab_Click(Index As Integer)
    Select Case Index
        Case 1: frmAnamnesiEsamiLab.Show
        Case 2: frmEsitoEsami.Show
        Case 3: frmRichiesteEsamiLab.Show
    End Select
End Sub

Private Sub mnuSottoPazTerapia_Click(Index As Integer)
    Select Case Index
        Case 1: frmTerapiaDialitica.Show
        Case 2: frmTerapiaDomiciliare.Show
        Case 3:
            If structIntestazione.sCodiceSTS = CODICESTS_HELIOS Then
                Call StampaRiepiloghiTerapieHelios
            Else
                Call StampaRiepiloghiTerapie
            End If
    End Select
End Sub

Private Sub mnuSottoTab_Click(Index As Integer)
    Select Case Index
        Case 1: Exit Sub
        Case 6, 15, 16
            Select Case Index
                Case 6: tTabelle = tpesame
                Case 15: tTabelle = tpEDTA
                Case 16: tTabelle = tpEDTA_MORTE
            End Select
            Unload frmTabelle
            frmTabelle.Show
        Case 2, 4
            Dim lfrmTabPersonaleElenco As New frmTabPersonaleElenco
            Select Case Index
                Case 2: lfrmTabPersonaleElenco.intTipoTabPersonale = enumTipoTabPersonale.MEDICI_REFERTANTI
                Case 4: lfrmTabPersonaleElenco.intTipoTabPersonale = enumTipoTabPersonale.PSICOLOGI
            End Select
            lfrmTabPersonaleElenco.Show
            Set lfrmTabPersonaleElenco = Nothing
        Case 3: frmMediciBase.Show
        Case 13: frmVociEsami.Show
        Case 14: frmTipiEsamiLab.Show
        Case Else
            Dim lfrmTabSingoloElenco As New frmTabSingoloElenco
            Select Case Index
                Case 5: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.ORGANO
                Case 7: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.filtro
                Case 8: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.LINEE
                Case 9: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.AGO
                Case 10: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.Medicinali
                Case 11: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.ANTICOAGULANTI
                Case 12: lfrmTabSingoloElenco.intTipoTabSingolo = enumTipoTabSingolo.TITOLIDIARIO
            End Select
            lfrmTabSingoloElenco.Show
            Set lfrmTabSingoloElenco = Nothing
    End Select
End Sub

Private Sub mnuSottoTabOrgan_Click(Index As Integer)
    Dim lfrmTabPersonaleElenco As New frmTabPersonaleElenco
        
    Select Case Index
        Case 2
            lfrmTabPersonaleElenco.intTipoTabPersonale = MEDICI_DIALISI
        Case 3
            lfrmTabPersonaleElenco.intTipoTabPersonale = INFERMIERI
        Case 1
            frmDirettoreSanitario.Show
            Exit Sub
    End Select
    lfrmTabPersonaleElenco.Show
    Set lfrmTabPersonaleElenco = Nothing
End Sub

Private Sub StampaRiepiloghiTerapieHelios()
    On Error GoTo gestione
    Dim intSessione As enumSessioni
    Dim strNomeSessione As String
    Dim strPeriodo As String
    
    Dim rsDataset As New Recordset
    Dim rsAppo As New Recordset
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsFiglio2 As Recordset
    

    Dim strSqlStampa As String
    Dim strSql As String
    Dim strExtraWhereTurno As String
    Dim strMedicinaliEsclusi As String
    Dim intDifferenzaTurno As Integer
    Dim rigaCodiceMedicinale() As Integer
    Dim indicePosizioneMedicinale As Integer
           
    frmPannelloSceltaTurno.Show 1
    intSessione = frmPannelloSceltaTurno.GetSessione
    Unload frmPannelloSceltaTurno

    If intSessione = tpNoneSession Then Exit Sub
    
    Select Case intSessione
        Case tpPariMattina: strNomeSessione = "Pari Mattina"
        Case tpPariPomeriggio: strNomeSessione = "Pari Pomeriggio"
        Case tpPariSera: strNomeSessione = "Pari Sera"
        Case tpDispariMattina: strNomeSessione = "Dispari Mattina"
        Case tpDispariPomeriggio: strNomeSessione = "Dispari Pomeriggio"
        Case tpDispariSera: strNomeSessione = "Dispari Sera"
    End Select
    
    Select Case intSessione
        Case tpPariMattina, tpDispariMattina: strPeriodo = "AM_INIZIO"
        Case tpPariPomeriggio, tpDispariPomeriggio: strPeriodo = "PM_INIZIO"
        Case tpPariSera, tpDispariSera: strPeriodo = "SR_INIZIO"
    End Select
    
    If intSessione = tpDispariMattina Or intSessione = tpDispariPomeriggio Or intSessione = tpDispariSera Then
        intDifferenzaTurno = 0
    Else
        intDifferenzaTurno = 1
    End If
    'seleziona i campi(farmaci) i cui valori non sono vuoti
    strExtraWhereTurno = " AND (" & strPeriodo & 1 + intDifferenzaTurno & "<>'' OR " & strPeriodo & 3 + intDifferenzaTurno & "<>'' OR " & strPeriodo & 5 + intDifferenzaTurno & "<>'') " & _
                         " AND (GIORNO" & 1 + intDifferenzaTurno & "=TRUE OR GIORNO" & 3 + intDifferenzaTurno & "=TRUE OR GIORNO" & 5 + intDifferenzaTurno & "=TRUE OR TUTTI_GIORNI=TRUE" & _
                         " OR NOT ISNULL(DATA_1) OR NOT ISNULL(DATA_2) OR NOT ISNULL(DATA_3))"
    

    strSqlStampa = "SHAPE APPEND " & _
            "       NEW adVarChar(10) AS LINKSUPERIORE, " & _
            "       ((SHAPE APPEND " & _
            "           NEW adVarChar(10) AS LINKSUPERIORE, " & _
            "           NEW adLongVarChar AS PAZIENTE, " & _
            "           NEW adVarChar(10) AS LINK1, " & _
            "           (( SHAPE APPEND " & _
            "               NEW adVarChar(10) AS LINK1, " & _
            "               NEW adVarChar(15) AS GIORNO, " & _
            "               NEW adLongVarChar AS MEDICINALE1, " & _
            "               NEW adLongVarChar AS MEDICINALE2, " & _
            "               NEW adLongVarChar AS MEDICINALE3, " & _
            "               NEW adLongVarChar AS MEDICINALE4, " & _
            "               NEW adLongVarChar AS MEDICINALE5, " & _
            "               NEW adLongVarChar AS MEDICINALE6, " & _
            "               NEW adLongVarChar AS MEDICINALE7, " & _
            "               NEW adLongVarChar AS MEDICINALE8, " & _
            "               NEW adLongVarChar AS MEDICINALE9, " & _
            "               NEW adLongVarChar AS MEDICINALE10, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE1, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE2, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE3, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE4, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE5, "
    strSqlStampa = strSqlStampa & _
            "               NEW adLongVarChar AS POSOLOGIANOTE6, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE7, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE8, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE9, " & _
            "               NEW adLongVarChar AS POSOLOGIANOTE10 " & _
            "       ) RELATE LINK1 TO LINK1 ) AS RES1" & _
            "       ) RELATE LINKSUPERIORE TO LINKSUPERIORE ) AS RES2"

    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open strSqlStampa, cnConn, adOpenStatic, adLockOptimistic
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intIndiceRiga As Integer
    Dim intNumeroMedicinali As Integer
    Const intNumeroMaxMedicinali As Integer = 7
    
    Dim intNumPagCorrente As Integer
    Dim intPuntiCorrente As Integer
    
    Const intPuntiPrimoLivello As Integer = 4
    Const intPuntiSecondoLivello As Integer = 7
    Const intPuntiTotali As Integer = 120 'determina il salto pagina - valore max 130
    
    strSql = "" & _
            " FROM      (((TERAPIE_DIALITICHE " & _
            "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=TERAPIE_DIALITICHE.CODICE_PAZIENTE) " & _
            "           INNER JOIN TURNI ON PAZIENTI.KEY=TURNI.CODICE_PAZIENTE) " & _
            "           INNER JOIN MEDICINALI ON TERAPIE_DIALITICHE.CODICE_MEDICINALE=MEDICINALI.KEY) " & _
            " WHERE     SOSPESA=FALSE AND POSOLOGIA<>'' " '& strExtraWhereTurno

    intPuntiCorrente = intPuntiTotali
    rsDataset.Open "SELECT DISTINCT PAZIENTI.KEY, PAZIENTI.COGNOME, PAZIENTI.NOME " & strSql & strExtraWhereTurno & " ORDER BY PAZIENTI.COGNOME, PAZIENTI.NOME", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        ReDim rigaCodiceMedicinale(0)
        intPuntiCorrente = intPuntiCorrente + intPuntiPrimoLivello
        If intPuntiCorrente + intPuntiSecondoLivello > intPuntiTotali Then
            rsMain.AddNew
            intPuntiCorrente = intPuntiPrimoLivello
            intNumPagCorrente = intNumPagCorrente + 1
            rsMain.Fields("LINKSUPERIORE") = intNumPagCorrente
        End If
        Set rsFiglio = rsMain.Fields("Res2").Value
        rsFiglio.AddNew
        rsFiglio.Fields("LINKSUPERIORE") = intNumPagCorrente
        rsFiglio.Fields("LINK1") = rsDataset("KEY") & "-" & intNumPagCorrente & "-0"
        rsFiglio.Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
        
        'Per incolonnare i farmaci correttamente elimina i farmaci con le date attribuite
        'Determina l'ordine in colonna dei medicinali (ordinati alfabeticamente)
        strExtraWhereTurno = " AND (" & strPeriodo & 1 + intDifferenzaTurno & "<>'' OR " & strPeriodo & 3 + intDifferenzaTurno & "<>'' OR " & strPeriodo & 5 + intDifferenzaTurno & "<>'') " & _
                         " AND (GIORNO" & 1 + intDifferenzaTurno & "=TRUE OR GIORNO" & 3 + intDifferenzaTurno & "=TRUE OR GIORNO" & 5 + intDifferenzaTurno & "=TRUE OR TUTTI_GIORNI=TRUE)"
                     
        rsAppo.Open "SELECT MEDICINALI.KEY " & strSql & strExtraWhereTurno & " AND PAZIENTI.KEY=" & rsDataset("KEY") & " ORDER BY MEDICINALI.NOME", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
        Do While Not rsAppo.EOF
            ReDim Preserve rigaCodiceMedicinale(UBound(rigaCodiceMedicinale) + 1)
            rigaCodiceMedicinale(UBound(rigaCodiceMedicinale)) = rsAppo("KEY")
            rsAppo.MoveNext
        Loop
        rsAppo.Close
        
        For j = 1 To 3
            rsAppo.Open "SELECT MEDICINALI.*, TERAPIE_DIALITICHE.*, PAZIENTI.NOME, PAZIENTI.COGNOME " & strSql & strExtraWhereTurno & " AND PAZIENTI.KEY=" & rsDataset("KEY") & " AND (GIORNO" & Choose(j, 1, 3, 5) + intDifferenzaTurno & "=TRUE OR TUTTI_GIORNI=TRUE) ORDER BY MEDICINALI.NOME", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
            strMedicinaliEsclusi = ""
            If rsAppo.RecordCount <> 0 Then
                
                For i = 1 To Int(rsAppo.RecordCount / intNumeroMaxMedicinali) + 1
                    intPuntiCorrente = intPuntiCorrente + intPuntiSecondoLivello
                    If intPuntiCorrente > intPuntiTotali Then
                        intPuntiCorrente = intPuntiPrimoLivello + intPuntiSecondoLivello
                        intNumPagCorrente = intNumPagCorrente + 1
                        rsFiglio.Update
                        rsMain.Update
                        
                        rsMain.AddNew
                        intPuntiCorrente = intPuntiPrimoLivello + intPuntiSecondoLivello
                        intNumPagCorrente = intNumPagCorrente + 1
                        rsMain.Fields("LINKSUPERIORE") = intNumPagCorrente
                        Set rsFiglio = rsMain.Fields("Res2").Value
                        rsFiglio.AddNew
                        rsFiglio.Fields("LINKSUPERIORE") = intNumPagCorrente
                        rsFiglio.Fields("LINK1") = rsDataset("KEY") & "-" & intNumPagCorrente & "-0"
                        rsFiglio.Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    End If
                        
                    If i > 1 Then
                        rsFiglio.Update
                        intPuntiCorrente = intPuntiCorrente + intPuntiPrimoLivello
                        rsFiglio.AddNew
                        rsFiglio.Fields("LINK1") = rsDataset("KEY") & "-" & intNumPagCorrente & "-" & i
                        rsFiglio.Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    End If
                    rsAppo.Filter = strMedicinaliEsclusi
                    Set rsFiglio2 = rsFiglio.Fields("Res1").Value
                    With rsFiglio2
                        .AddNew
                        .Fields("LINK1") = rsFiglio.Fields("LINK1")
                        .Fields("GIORNO") = UCase(Mid(WeekdayName(Choose(j, 1, 3, 5) + intDifferenzaTurno, False, vbMonday), 1, 3))
                        If rsAppo.RecordCount > intNumeroMaxMedicinali Then
                            intNumeroMedicinali = intNumeroMaxMedicinali
                        Else
                            intNumeroMedicinali = rsAppo.RecordCount
                        End If
                        'stampa i farmaci sulla stessa riga
                        For k = 1 To intNumeroMedicinali
                            indicePosizioneMedicinale = 0
                            For intIndiceRiga = 1 To UBound(rigaCodiceMedicinale)
                                If rigaCodiceMedicinale(intIndiceRiga) = rsAppo("MEDICINALI.KEY") Then
                                    indicePosizioneMedicinale = intIndiceRiga
                                    Exit For
                                End If
                            Next
                            
                            .Fields("MEDICINALE" & indicePosizioneMedicinale) = UCase(rsAppo("MEDICINALI.NOME"))
                            .Fields("POSOLOGIANOTE" & indicePosizioneMedicinale) = "( " & rsAppo("POSOLOGIA") & IIf(rsAppo("NOTE") <> "", " - " & rsAppo("NOTE"), "") & " )"
                            strMedicinaliEsclusi = " NOT MEDICINALI.KEY=" & rsAppo("MEDICINALI.KEY") & " AND " & strMedicinaliEsclusi
                            rsAppo.MoveNext
                        Next k
                        strMedicinaliEsclusi = Mid(strMedicinaliEsclusi, 1, Len(strMedicinaliEsclusi) - 4)
                        .Update
                    End With
                Next i
            End If
            rsAppo.Close
        Next j
        
' Stampa Farmaci x Data
'------------------------
   Dim MGiorno As String

    Dim anno As Integer
    Dim mm As Integer
    Dim gg As Integer
    'Determina l'ultimo giorno del mese
    anno = Year(date)
    mm = Month(date)
    gg = Day(DateSerial(anno, mm + 1, 0))
    
     rsAppo.Open "SELECT MEDICINALI.*, TERAPIE_DIALITICHE.*, PAZIENTI.NOME, PAZIENTI.COGNOME " & strSql & " AND " & _
      "PAZIENTI.KEY=" & rsDataset("KEY") & " AND (DATA_1 BETWEEN #" & mm & "/01/" & anno & "# AND #" & mm & "/" & gg & "/" & anno & "#" & _
      " OR DATA_2 BETWEEN #" & mm & "/01/" & anno & "# AND #" & mm & "/" & gg & "/" & anno & "#" & _
      " OR DATA_3 BETWEEN #" & mm & "/01/" & anno & "# AND #" & mm & "/" & gg & "/" & anno & "#" & _
      ") ORDER BY MEDICINALI.NOME", cnPrinc, adOpenKeyset, adLockReadOnly, adCmdText
         '  strMedicinaliEsclusi = ""
            If rsAppo.RecordCount <> 0 Then
               Do While Not rsAppo.EOF
                'For i = 1 To Int(rsAppo.RecordCount / intNumeroMaxMedicinali) + 1
                    intPuntiCorrente = intPuntiCorrente + intPuntiSecondoLivello
                    If intPuntiCorrente > intPuntiTotali Then
                        intPuntiCorrente = intPuntiPrimoLivello + intPuntiSecondoLivello
                        intNumPagCorrente = intNumPagCorrente + 1
                        rsFiglio.Update
                        rsMain.Update
                        
                        rsMain.AddNew
                        intPuntiCorrente = intPuntiPrimoLivello + intPuntiSecondoLivello
                        intNumPagCorrente = intNumPagCorrente + 1
                        rsMain.Fields("LINKSUPERIORE") = intNumPagCorrente
                        Set rsFiglio = rsMain.Fields("Res2").Value
                        rsFiglio.AddNew
                        rsFiglio.Fields("LINKSUPERIORE") = intNumPagCorrente
                        rsFiglio.Fields("LINK1") = rsDataset("KEY") & "-" & intNumPagCorrente & "-0"
                        rsFiglio.Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    End If
                        
                 '   If i > 1 Then
                 '       rsFiglio.Update
                 '       intPuntiCorrente = intPuntiCorrente + intPuntiPrimoLivello
                 '       rsFiglio.AddNew
                 '       rsFiglio.Fields("LINK1") = rsDataset("KEY") & "-" & intNumPagCorrente & "-" & i
                 '       rsFiglio.Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                 '   End If
                 '   rsAppo.Filter = strMedicinaliEsclusi
                 '   Set rsFiglio2 = rsFiglio.Fields("Res1").Value
                    indicePosizioneMedicinale = 1
                    With rsFiglio2
                        .AddNew
                        .Fields("LINK1") = rsFiglio.Fields("LINK1")
                        
                        If rsAppo("DATA_1") <> "" And Month(rsAppo("DATA_1").Value) = Month(date) Then
                            MGiorno = CStr(rsAppo("DATA_1"))
                        ElseIf rsAppo("DATA_2") <> "" And Month(rsAppo("DATA_2").Value) = Month(date) Then
                            MGiorno = CStr(rsAppo("DATA_2"))
                        ElseIf rsAppo("DATA_3") <> "" And Month(rsAppo("DATA_3").Value) = Month(date) Then
                            MGiorno = CStr(rsAppo("DATA_3"))
                        End If
                        
                        .Fields("GIORNO") = MGiorno
                        .Fields("MEDICINALE" & indicePosizioneMedicinale) = UCase(rsAppo("MEDICINALI.NOME"))
                        .Fields("POSOLOGIANOTE" & indicePosizioneMedicinale) = "( " & rsAppo("POSOLOGIA") & IIf(rsAppo("NOTE") <> "", " - " & rsAppo("NOTE"), "") & " )"
                        .Update
                    End With
                rsAppo.MoveNext
                Loop
                ' Next i
            End If
            
        rsAppo.Close
        rsFiglio.Update
       
    rsDataset.MoveNext
    Loop
    rsDataset.Close
    
    If rsMain.RecordCount <> 0 Then
        Set rptFoglioTerapiaPerMedicinale = Nothing
        Set rptFoglioTerapiaPerMedicinale.DataSource = rsMain
        rptFoglioTerapiaPerMedicinale.LeftMargin = 400
        rptFoglioTerapiaPerMedicinale.RightMargin = 0
        rptFoglioTerapiaPerMedicinale.TopMargin = 0
        rptFoglioTerapiaPerMedicinale.BottomMargin = 0
        rptFoglioTerapiaPerMedicinale.Orientation = rptOrientLandscape
        rptFoglioTerapiaPerMedicinale.Sections("intestazione").Controls("lblTurno").Caption = strNomeSessione
        rptFoglioTerapiaPerMedicinale.PrintReport True, rptRangeAllPages
    Else
        MsgBox "Nessuna terapia trovata", vbInformation, "Stampa Terapie"
        Call StampaRiepiloghiTerapieHelios
    End If
    
    Set rsAppo = Nothing
    Set rsDataset = Nothing
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub

Private Sub StampaRiepiloghiTerapie()
    On Error GoTo gestione
    Dim intSessione As enumSessioni
    Dim strNomeSessione As String
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    Dim rsTerapia As Recordset
    Dim rsDataset As Recordset
    
    Dim strPeriodo As String
    Dim strSql As String
    Dim giorni() As Variant
    Dim i As Integer
    Dim condizione As String
    Dim somma As Integer
    Dim continua As Boolean
    
    frmPannelloSceltaTurno.Show 1
    intSessione = frmPannelloSceltaTurno.GetSessione
    Unload frmPannelloSceltaTurno
    
    If intSessione = tpNoneSession Then Exit Sub
    
    Select Case intSessione
        Case tpPariMattina, tpDispariMattina: strPeriodo = "AM_INIZIO"
        Case tpPariPomeriggio, tpDispariPomeriggio: strPeriodo = "PM_INIZIO"
        Case tpPariSera, tpDispariSera: strPeriodo = "SR_INIZIO"
    End Select
    
    Select Case intSessione
        Case tpPariMattina, tpPariPomeriggio, tpPariSera:  giorni = Array(2, 4, 6)
        Case tpDispariMattina, tpDispariPomeriggio, tpDispariSera: giorni = Array(1, 3, 5)
    End Select
    
    Select Case intSessione
        Case tpPariMattina: strNomeSessione = "Pari Mattina"
        Case tpPariPomeriggio: strNomeSessione = "Pari Pomeriggio"
        Case tpPariSera: strNomeSessione = "Pari Sera"
        Case tpDispariMattina: strNomeSessione = "Dispari Mattina"
        Case tpDispariPomeriggio: strNomeSessione = "Dispari Pomeriggio"
        Case tpDispariSera: strNomeSessione = "Dispari Sera"
    End Select
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(70) AS PAZIENTE, " & _
                "       NEW adVarChar(10) AS DATA_NASCITA, " & _
                "       NEW adVarChar(2) AS ANNI, " & _
                "       NEW adInteger AS LINK1, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS LINK1, " & _
                "           NEW adDate AS DATA, " & _
                "           NEW adVarChar(50) AS MEDICINALE, " & _
                "           NEW adLongVarChar AS POSOLOGIAENOTE, " & _
                "           NEW adInteger AS SOMMINISTRAZIONE, " & _
                "           NEW adLongVarChar as GIORNI " & _
                "       ) RELATE LINK1 TO LINK1 " & _
                "       ) AS Res1 "

    
    For i = 0 To UBound(giorni)
        condizione = condizione & strPeriodo & giorni(i) & "<>"""" OR "
    Next i
    condizione = Left(condizione, Len(condizione) - 4)
    
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    ' terapie non sospese
    strSql = "SELECT    PAZIENTI.KEY, COGNOME, NOME, DATA_NASCITA, STATO, MAX(DATA_ARRIVO) AS MAX_ARRIVO, MAX(DATA_PARTENZA) AS MAX_PARTENZA " & _
            " FROM      ((TURNI " & _
            "           INNER JOIN PAZIENTI ON PAZIENTI.KEY=TURNI.CODICE_PAZIENTE) " & _
            "           LEFT OUTER JOIN PAZIENTI_OSPITI ON PAZIENTI_OSPITI.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
            "           WHERE (" & condizione & ") AND " & _
            "           (STATO=0 OR STATO=4) " & _
            "GROUP BY   PAZIENTI.KEY, COGNOME, NOME, DATA_NASCITA, STATO " & _
            "ORDER BY   COGNOME, NOME"
    Set rsTerapia = New Recordset
    Set rsDataset = New Recordset
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not rsDataset.EOF
        If rsDataset("STATO") = 4 Then
            If Not (rsDataset("MAX_ARRIVO") <= date And rsDataset("MAX_PARTENZA") >= date) Then
                continua = False
            End If
        Else
            continua = True
        End If
        If continua Then
            With rsMain
                strSql = "SELECT    * " & _
                        "FROM       (TERAPIE_DIALITICHE " & _
                        "           INNER JOIN MEDICINALI ON MEDICINALI.KEY=TERAPIE_DIALITICHE.CODICE_MEDICINALE) " & _
                        "WHERE      CODICE_PAZIENTE=" & rsDataset("KEY") & " AND " & _
                        "           SOSPESA=FALSE " & _
                        "ORDER BY   DATA DESC"
                rsTerapia.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not (rsTerapia.EOF And rsTerapia.BOF) Then
                    .AddNew
                    .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    .Fields("DATA_NASCITA") = Day(rsDataset("DATA_NASCITA")) & "/" & Month(rsDataset("DATA_NASCITA")) & "/" & Year(rsDataset("DATA_NASCITA"))
                    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
                        somma = -1
                    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
                        somma = -1
                    Else
                        somma = 0
                    End If
                    .Fields("ANNI") = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
                    .Fields("LINK1") = rsDataset("KEY")
                    Do While Not rsTerapia.EOF
                    'Se al farmaco non è associato nessun giorno e nessuna data non deve uscire
                    If rsTerapia("GIORNO1").Value = 0 And rsTerapia("GIORNO2").Value = 0 And rsTerapia("GIORNO3").Value = 0 And rsTerapia("GIORNO4").Value = 0 And rsTerapia("GIORNO5").Value = 0 And rsTerapia("GIORNO6").Value = 0 And rsTerapia("GIORNO7").Value = 0 And rsTerapia("TUTTI_GIORNI").Value = 0 And IsNull(rsTerapia("DATA_1")) And IsNull(rsTerapia("DATA_2")) And IsNull(rsTerapia("DATA_3")) Then
                    
                    Else
                    'Se il mese del farmaco corrisponde a quello di sistema lo stampa
                    If Month(rsTerapia("DATA_1").Value) = Month(date) Or IsNull(rsTerapia("DATA_1")) Or (Month(rsTerapia("DATA_2").Value) = Month(date) Or IsNull(rsTerapia("DATA_2"))) Or (Month(rsTerapia("DATA_3").Value) = Month(date) Or IsNull(rsTerapia("DATA_3"))) Then
                         Set rsFiglio = .Fields("Res1").Value
                        With rsFiglio
                            .AddNew
                            .Fields("LINK1") = rsDataset("KEY")
                            .Fields("DATA") = rsTerapia("DATA")
                            .Fields("MEDICINALE") = rsTerapia("NOME")
                            'viene fatto un' ulteriore controllo, nel caso in cui un farmaco dovesse avere + di una data, e una della quale non corrisponde al mese corrente
                            If Month(rsTerapia("DATA_1").Value) = Month(date) Then
                                .Fields("GIORNI") = CStr(rsTerapia("DATA_1"))
                            End If
                            If Month(rsTerapia("DATA_2").Value) = Month(date) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & "  " & CStr(rsTerapia("DATA_2"))
                            End If
                            If Month(rsTerapia("DATA_3").Value) = Month(date) Then
                                .Fields("GIORNI") = .Fields("GIORNI") & "  " & CStr(rsTerapia("DATA_3"))
                            End If
                            .Fields("POSOLOGIAENOTE") = rsTerapia("POSOLOGIA") & "-" & rsTerapia("NOTE")
                            .Fields("SOMMINISTRAZIONE") = rsTerapia("SOMMINISTRAZIONE")
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
                    End If
                         End If ' questo
                        rsTerapia.MoveNext
                    Loop
                    .Update
                Else
                    .AddNew
                    .Fields("PAZIENTE") = rsDataset("COGNOME") & " " & rsDataset("NOME")
                    .Fields("DATA_NASCITA") = Day(rsDataset("DATA_NASCITA")) & "/" & Month(rsDataset("DATA_NASCITA")) & "/" & Year(rsDataset("DATA_NASCITA"))
                    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
                        somma = -1
                    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
                        somma = -1
                    Else
                        somma = 0
                    End If
                    .Fields("ANNI") = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
                    .Fields("LINK1") = rsDataset("KEY")
                    Set rsFiglio = .Fields("Res1").Value
                    With rsFiglio
                        .AddNew
                        .Fields("LINK1") = rsDataset("KEY")
                        .Fields("DATA") = Null
                        .Fields("MEDICINALE") = ""
                        .Fields("POSOLOGIAENOTE") = "NESSUNA TERAPIA"
                        .Fields("SOMMINISTRAZIONE") = 0
                        .Fields("GIORNI") = ""
                        .Update
                    End With
                    .Update
                End If
                rsTerapia.Close
            End With
        End If
        rsDataset.MoveNext
    Loop
    rsDataset.Close
    Set rsDataset = Nothing
    Set rsTerapia = Nothing
    
    If rsMain.RecordCount <> 0 Then
        Set rptStampaRiepiloghiTerapie.DataSource = rsMain
        rptStampaRiepiloghiTerapie.LeftMargin = 0
        rptStampaRiepiloghiTerapie.RightMargin = 0
        rptStampaRiepiloghiTerapie.Sections("intestazione").Controls("lblTurno").Caption = "Terapia Turno " & strNomeSessione
        rptStampaRiepiloghiTerapie.PrintReport True, rptRangeAllPages
    Else
        MsgBox "Nessuna terapia trovata", vbInformation, "Stampa Riepiloghi"
    End If
    
    Exit Sub
gestione:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox "Descrizione: " & Err.Description, vbCritical, "Errore n# " & Err.Number
    End If
End Sub
Private Sub StampaTurni()
 
    Dim intSessione As enumSessioni
    Dim strNomeSessione As String
    Dim SQLString As String
    Dim rsDataset As New Recordset
    Dim numPazienti As Integer
    Dim strGiorni As String
    Dim strGiornoIni(3) As Variant
    Dim strGiornoFin(3) As Variant
    Dim strPeriodo As String
    Dim strPeriodoFin As String
    Dim giorni() As Variant
    Dim i As Integer
    Dim condizione As String
    
    frmPannelloSceltaTurno.Show 1
    intSessione = frmPannelloSceltaTurno.GetSessione
    Unload frmPannelloSceltaTurno
    
    If intSessione = tpNoneSession Then Exit Sub
    
    Select Case intSessione
        Case tpPariMattina, tpDispariMattina: strPeriodo = "AM_INIZIO"
            strPeriodoFin = "AM_FINE"
        Case tpPariPomeriggio, tpDispariPomeriggio: strPeriodo = "PM_INIZIO"
            strPeriodoFin = "PM_FINE"
        Case tpPariSera, tpDispariSera: strPeriodo = "SR_INIZIO"
            strPeriodoFin = "SR_FINE"
    End Select
    
    Select Case intSessione
        Case tpPariMattina, tpPariPomeriggio, tpPariSera:  strGiorni = "Martedì             Giovedì             Sabato"
            giorni = Array(2, 4, 6)
        Case tpDispariMattina, tpDispariPomeriggio, tpDispariSera: strGiorni = "Lunedì             Mercoledì            Venerdì"
            giorni = Array(1, 3, 5)
    End Select
    
    Select Case intSessione
        Case tpPariMattina: strNomeSessione = "Turno Pari Mattina"
        Case tpPariPomeriggio: strNomeSessione = "Turno Pari Pomeriggio"
        Case tpPariSera: strNomeSessione = "Turno Pari Sera"
        Case tpDispariMattina: strNomeSessione = "Turno Dispari Mattina"
        Case tpDispariPomeriggio: strNomeSessione = "Turno Dispari Pomeriggio"
        Case tpDispariSera: strNomeSessione = "Turno Dispari Sera"
    End Select
    
    For i = 0 To UBound(giorni)
        condizione = condizione & strPeriodo & giorni(i) & "<>"""" OR "
        strGiornoIni(i) = strPeriodo & giorni(i)
        strGiornoFin(i) = strPeriodoFin & giorni(i)
    Next i
    condizione = Left(condizione, Len(condizione) - 4)
    
SQLString = "SELECT COGNOME, NOME, TURNI." & strGiornoIni(0) & " AS GGINI1, TURNI." & strGiornoIni(1) & " AS GGINI2,TURNI." & strGiornoIni(2) & " AS GGINI3,TURNI." & strGiornoFin(0) & " AS GGFIN1, TURNI." & strGiornoFin(1) & " AS GGFIN2,TURNI." & strGiornoFin(2) & " AS GGFIN3,APPARATI.POSTAZIONE, APPARATI.NUMERO_APPARATO, APPARATI.MODELLO AS MONITOR, APPARATI.TIPO " & _
             "FROM ((PAZIENTI " & _
             "INNER JOIN TURNI ON PAZIENTI.KEY = TURNI.CODICE_PAZIENTE ) " & _
             "INNER JOIN APPARATI  ON TURNI.CODICE_RENE= APPARATI.KEY ) " & _
             "WHERE (" & condizione & ") AND (PAZIENTI.STATO = 0) " & _
             "ORDER BY  PAZIENTI.COGNOME, PAZIENTI.NOME"
      
    rsDataset.Open SQLString, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
   
If rsDataset.RecordCount <> 0 Then
        numPazienti = rsDataset.RecordCount
        
        Set rptStampaTurni.DataSource = rsDataset
'        rptStampaTurni.Orientation = rptOrientLandscape
        rptStampaTurni.LeftMargin = 0
        rptStampaTurni.RightMargin = 0
        rptStampaTurni.TopMargin = 0
        rptStampaTurni.Sections("intestazione").Controls("lblTurno").Caption = strNomeSessione
        rptStampaTurni.Sections("intestazione").Controls("lblGiorni").Caption = strGiorni
        rptStampaTurni.Sections("Section5").Controls.Item("lblPazienti").Caption = numPazienti
        rptStampaTurni.PrintReport True, rptRangeAllPages
    Else
        MsgBox "Non risultano assegnati turni ai pazienti", vbInformation, "Stampa Riepiloghi"
End If
      
rsDataset.Close
Set rsDataset = Nothing
End Sub

Private Sub mnuStampaMediciBase_Click()
    ' stampa il report
    Dim SQLString As String
    Dim rsMedici As New Recordset
    Dim rsPazienti As New Recordset
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsFiglio As Recordset
    
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(35) AS COGNOME, " & _
                "       NEW adVarChar(35) AS NOME, " & _
                "       NEW adVarChar(35) AS INDIRIZZO, " & _
                "       NEW adVarChar(35) AS COMUNE, " & _
                "       NEW adVarChar(5) AS CAP, " & _
                "       NEW adVarChar(10) AS PROV, " & _
                "       NEW adVarChar(35) AS TELEFONO, " & _
                "       NEW adVarChar(35) AS CELLULARE, " & _
                "       NEW adVarChar(35) AS STUDIO, " & _
                "       NEW adVarChar(35) AS EMAIL, " & _
                "       NEW adInteger AS CODICE_MEDICO, " & _
                "       NEW adVarChar(7) AS CODICE, " & _
                "       (( SHAPE APPEND " & _
                "           NEW adInteger AS CODICE_MEDICO, " & _
                "           NEW adVarChar(100) AS PAZIENTE " & _
                "       ) RELATE CODICE_MEDICO TO CODICE_MEDICO " & _
                "       ) AS Res1 "
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic

    
    Set rsMedici = New Recordset
    rsMedici.Open "SELECT * FROM MEDICI_BASE ORDER BY COGNOME, NOME", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsPazienti.Open "SELECT KEY,COGNOME,NOME,CODICE_MEDICO FROM PAZIENTI", cnPrinc, adOpenKeyset, adLockOptimistic, adCmdText
    Do While Not rsMedici.EOF
        With rsMain
            .AddNew
            .Fields("COGNOME") = rsMedici("COGNOME")
            .Fields("NOME") = rsMedici("NOME")
            .Fields("INDIRIZZO") = rsMedici("INDIRIZZO")
            .Fields("COMUNE") = rsMedici("COMUNE")
            .Fields("CAP") = rsMedici("CAP")
            .Fields("PROV") = rsMedici("PROV")
            .Fields("TELEFONO") = rsMedici("TELEFONO")
            .Fields("CELLULARE") = rsMedici("CELLULARE")
            .Fields("STUDIO") = rsMedici("STUDIO")
            .Fields("EMAIL") = rsMedici("EMAIL")
            .Fields("CODICE_MEDICO") = rsMedici("KEY")
            .Fields("CODICE") = rsMedici("CODICE")
            Set rsFiglio = .Fields("Res1").Value
            With rsFiglio
                rsPazienti.Filter = "CODICE_MEDICO=" & rsMedici("KEY")
                Do While Not rsPazienti.EOF
                    .AddNew
                    .Fields("CODICE_MEDICO") = rsMedici("KEY")
                    .Fields("PAZIENTE") = "Paziente associato:  " & rsPazienti("COGNOME") & "  " & rsPazienti("NOME")
                    .Update
                    rsPazienti.MoveNext
                Loop
            End With
            .Update
        End With
        rsMedici.MoveNext
    Loop
    Set rsMedici = Nothing
    Set rsPazienti = Nothing
    
    Set rptMedici.DataSource = rsMain
    rptMedici.Orientation = rptOrientLandscape
    rptMedici.PrintReport True, rptRangeAllPages

End Sub

Private Sub mnuStampaPaz_Click()
    Dim dataDal As String
    Dim dataAl As String
    Dim strSql As String
    Dim numPazienti As Integer
    Dim strNomeStatoPaziente As String
    
    frmPannelloFiltroStato.Show 1
    
    If tFiltroStato.statoPaziente = tpNoneStatoPaziente Then Exit Sub
    
    If Not tFiltroStato.isTutteLeDate Then
       dataDal = DateValue(tFiltroStato.dataDal)
       dataAl = DateValue(tFiltroStato.dataAl)
'      dataDal = DateValue(Month(tFiltroStato.dataDal) & "/" & Day(tFiltroStato.dataDal) & "/" & Year(tFiltroStato.dataDal))
'      dataAl = DateValue(Month(tFiltroStato.dataAl) & "/" & Day(tFiltroStato.dataAl) & "/" & Year(tFiltroStato.dataAl))
    End If
    
    Select Case tFiltroStato.statoPaziente
        Case tpAMBULATORIALE: strNomeStatoPaziente = "Ambulatoriale"
        Case TPDECEDUTO: strNomeStatoPaziente = "Deceduto"
        Case tpDIALISI: strNomeStatoPaziente = "In Dialisi"
        Case TPOSPITE: strNomeStatoPaziente = "Ospite"
        Case TPTRAPIANTO: strNomeStatoPaziente = "Trapiantato"
        Case TPTRASFERITO: strNomeStatoPaziente = "Trasferito"
    End Select
    
    
    If tFiltroStato.statoPaziente = TPOSPITE Then
        strSql = "  SELECT      COGNOME, PAZIENTI.NOME as PAZIENTINOME, DATA_NASCITA, CODICE_FISCALE, INDIRIZZO, COMUNI.NOME as COMUNINOME, CODICE_DOCUMENTO, TELEFONO, CELLULARE, TIPO_DOCUMENTO, MAX(DATA_PARTENZA) AS DATA_P, MAX(DATA_ARRIVO) AS DATA_A " & _
                 "  FROM        ((PAZIENTI " & _
                 "              LEFT OUTER JOIN PAZIENTI_OSPITI ON PAZIENTI_OSPITI.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
                 "              INNER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) " & _
                 "  WHERE       STATO=4 " & _
                 "  GROUP BY    COGNOME, PAZIENTI.NOME, CODICE_FISCALE, INDIRIZZO, COMUNI.NOME, CODICE_DOCUMENTO, TELEFONO, CELLULARE, TIPO_DOCUMENTO"
        If Not tFiltroStato.isTutteLeDate Then
            strSql = strSql & _
                 "  HAVING (MAX(DATA_ARRIVO) BETWEEN #" & dataDal & "# AND #" & dataAl & "#) "
        End If
    ElseIf tFiltroStato.statoPaziente = tpDIALISI Then
        ' In dialisi
        strSql = "  SELECT  COGNOME, PAZIENTI.NOME as PAZIENTINOME, DATA_NASCITA, CODICE_FISCALE, INDIRIZZO, COMUNI.NOME as COMUNINOME, CODICE_DOCUMENTO, TELEFONO, CELLULARE, TIPO_DOCUMENTO, DATA_INIZIO, DATA1 " & _
                 "  FROM    ((PAZIENTI " & _
                 "          LEFT OUTER JOIN ANAMNESI_NEFROLOGICHE ON ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
                 "          INNER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) "
        If Not tFiltroStato.isTutteLeDate Then
            strSql = strSql & _
                 "  WHERE   (((DATA_INIZIO < #" & dataAl & "#)) AND ((DATA_FINE> #" & dataDal & "#) or (DATA_FINE is NULL)))"
        End If
        strSql = strSql & _
                 "  ORDER BY COGNOME"
    Else
        ' Deceduti, trapiantati, trasferiti, ambulatoriale
        strSql = "  SELECT  COGNOME, PAZIENTI.NOME as PAZIENTINOME, DATA_NASCITA, CODICE_FISCALE, INDIRIZZO, COMUNI.NOME as COMUNINOME, CODICE_DOCUMENTO, TELEFONO, CELLULARE, TIPO_DOCUMENTO, STATODATA, DATA1 " & _
                 "  FROM    ((PAZIENTI " & _
                 "          LEFT OUTER JOIN ANAMNESI_NEFROLOGICHE ON ANAMNESI_NEFROLOGICHE.CODICE_PAZIENTE=PAZIENTI.KEY) " & _
                 "          INNER JOIN COMUNI ON COMUNI.KEY=PAZIENTI.CODICE_COMUNE_RESIDENZA) " & _
                 "  WHERE   STATO=" & tFiltroStato.statoPaziente
        If Not tFiltroStato.isTutteLeDate Then
            Select Case tFiltroStato.statoPaziente
                Case 5:
                    strSql = strSql & _
                     "          AND (DATA_INIZIO BETWEEN #" & dataDal & "# AND #" & dataAl & "#) "
                Case 1, 2, 3:
                    strSql = strSql & _
                     "          AND (STATODATA BETWEEN #" & dataDal & "# AND #" & dataAl & "#) "
            End Select
        End If
        strSql = strSql & _
                 "  ORDER BY COGNOME"
    End If
    
    Dim rsPazienti As New Recordset
'    Debug.Print strSql
    rsPazienti.Open strSql, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Not (rsPazienti.EOF And rsPazienti.BOF) Then
        numPazienti = rsPazienti.RecordCount

        Set rptPazienti = Nothing
        Set rptPazienti.DataSource = rsPazienti
        rptPazienti.Orientation = rptOrientLandscape
        rptPazienti.Sections("intestazione").Controls.Item("lblTitolo").Caption = "LISTA PAZIENTI " & Choose(tFiltroStato.statoPaziente + 1, "IN DIALISI", "DECEDUTI", "TRASFERITI", "TRAPIANTATI", "OSPITI", "AMBULATORIALI")
        If dataDal = "" Then
            rptPazienti.Sections("intestazione").Controls.Item("lblDal").Caption = "Da Tutte le date"
        Else
            rptPazienti.Sections("intestazione").Controls.Item("lblDal").Caption = "Dal " & dataDal
        End If
        If dataAl = "" Then
            rptPazienti.Sections("intestazione").Controls.Item("lblAl").Caption = "A Tutte le date"
        Else
            rptPazienti.Sections("intestazione").Controls.Item("lblAl").Caption = "al " & dataAl
        End If
        rptPazienti.Sections("intestazione").Controls.Item("lblDataPrima").Caption = "Data " & Choose(tFiltroStato.statoPaziente + 1, "inizio dialisi in sede", "di   decesso", "di trasferimento", "di trapianto", "di     arrivo", "inizio dialisi in sede")
        rptPazienti.Sections("intestazione").Controls.Item("lblDataSeconda").Caption = "Data " & IIf(tFiltroStato.statoPaziente = 4, "di partenza", "inizio dialisi")
        rptPazienti.Sections("Section5").Controls.Item("lblPazienti").Caption = numPazienti
        rptPazienti.LeftMargin = 0
        rptPazienti.RightMargin = 0.1
        
        rptPazienti.TopMargin = 0
        rptPazienti.PrintReport True, rptRangeAllPages
    Else
        MsgBox "Nessun paziente trovato con lo stato: " & strNomeStatoPaziente, vbInformation, "Stampa pazienti"
    End If
    rsPazienti.Close
    Set rsPazienti = Nothing
End Sub

Private Sub mnuStampaRiepilogo_Click(Index As Integer)
    Select Case Index
        Case 0: tStampeRiepilogo = tpFATTURA
        Case 1: tStampeRiepilogo = tpXPAZIENTE
        Case 2: tStampeRiepilogo = tpXTOTALIPERPRESTAZIONE
        Case 3: tStampeRiepilogo = tpXTOTALIPERASL
        Case 4: tStampeRiepilogo = tpXMAZZETTEDISTRETTI
        Case 5: tStampeRiepilogo = TPXMAZZETTEMENSILI
        Case 6: tStampeRiepilogo = tpXMAZZETTASINGOLA
        Case 7: tStampeRiepilogo = tpXASLDISTRETTI
        Case 8: tStampeRiepilogo = tpXIMPEGNATIVE
    End Select
    frmStampaFattureRiepilogo.Show 1
End Sub

Private Sub mnuTabFatt_Click(Index As Integer)
    Select Case Index
        Case 0: tTabelle = tpRegioni
        Case 1: tTabelle = tpCOMUNI
        Case 2: tTabelle = tpasl
        Case 3: tTabelle = tpDISTRETTI
        Case 4: tTabelle = tpESENZIONI
        Case 5: tTabelle = tpNOMENCLATORE
        Case 6: frmAccompagnatori.Show
                Exit Sub
    End Select
    ' chiude il form se caricato
    Unload frmTabelle
    frmTabelle.Show
End Sub

Private Sub mnuTsatAnnuale_Click()
    tStampa = tpTSATANNUALE
    frmStampaFiltri.Show 1
End Sub

Private Sub cmdToolbar_Click(Index As Integer)
    Me.SetFocus
    Select Case Index
        Case 0
            mnuSottoPaz_Click (1)
        Case 1
            mnuSottoPazAne_Click (1)
        Case 2
            mnuSottoPazAne_Click (2)
        Case 3
            mnuSottoPazAne_Click (3)
        Case 4
            mnuSottoPazEsami_Click (1)
        Case 5
            mnuSottoPazEsamiLab_Click (1)
        Case 6
            mnuSottoPazEsamiLab_Click (2)
        Case 7
            mnuSottoPazTerapia_Click (2)
        Case 8
            mnuSottoPazTerapia_Click (1)
        Case 9
            mnuSottoPaz_Click (5)
        Case 10
            mnuSottoPaz_Click (6)
        Case 11
            mnuSottoDialisiTurni_Click (1)
        Case 12
            mnuSottoDialisiScheda_Click (1)
        Case 13
            mnuSottoDialisi_Click (4)
        Case 14
            mnuCaricaPrescrizione_Click
        Case 15
            mnuGestioneFileC_Click
        Case 16
            mnuGestioneFileXml_Click
        Case 17
            Call CloseAllForm
    End Select
End Sub

Private Sub CloseAllForm()
    Dim lForm As Form

    oPazientiKey.listFormAperti.Refresh
    For Each lForm In Forms
        If lForm.Name = "frmAccessi" Or _
            lForm.Name = "frmAnamnesiDialitica" Or _
            lForm.Name = "frmDiario" Or _
            lForm.Name = "frmAnamnesiEsamiLab" Or _
            lForm.Name = "frmAnamnesiNefro" Or _
            lForm.Name = "frmAnamnesiPat" Or _
            lForm.Name = "frmEsamiStrumentali" Or _
            lForm.Name = "frmEsitoEsami" Or _
            lForm.Name = "frmPaziente" Or _
            lForm.Name = "frmTerapiaDialitica" Or _
            lForm.Name = "frmTerapiaDomiciliare" Or _
            lForm.Name = "frmMediciBase" Or _
            lForm.Name = "frmTabelle" Or _
            lForm.Name = "frmTabPersonaleElenco" Or _
            lForm.Name = "frmTabSingoloElenco" Or _
            lForm.Name = "frmTipiEsamiLab" Or _
            lForm.Name = "frmVociEsami" Then
            Unload lForm
        End If
    Next
End Sub

Public Sub SubClassMenuXP()

    '/ this code is made by MenuCreator add-in

    '/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
    mnuPaziente.Caption = "&Gestione pazienti"
          mnuSottoPaz(1).Caption = "&Anagrafica Generale"
          mnuSottoPaz(2).Caption = "&Anamnesi"
          mnuSottoPazAne(1).Caption = "&Patologica Remota e Familiare"
          mnuSottoPazAne(2).Caption = "&Nefrologica"
          mnuSottoPazAne(3).Caption = "&Dialitica"
          mnuSottoPaz(3).Caption = "&Esami"
          mnuSottoPazEsami(1).Caption = "&Strumentali"
          mnuSottoPazEsami(2).Caption = "&Laboratorio"
          mnuSottoPazEsamiLab(1).Caption = "&Registrazione"
          mnuSottoPazEsamiLab(2).Caption = "&Consultazione"
          mnuSottoPazEsamiLab(3).Caption = "&Prescrizione"
          mnuSottoPaz(4).Caption = "&Terapia"
          mnuSottoPazTerapia(1).Caption = "&Dialitica"
          mnuSottoPazTerapia(2).Caption = "&Domiciliare"
          mnuSottoPazTerapia(3).Caption = "&Stampa Riepiloghi"
          mnuSottoPaz(5).Caption = "&Accessi Vascolari"
          mnuSottoPaz(6).Caption = "&Diario Clinico"
          mnuSottoPaz(7).Caption = "&Scansione Documenti Pazienti"
          
    mnuDialisi.Caption = "Gestione &Emodialisi"
          mnuSottoDialisi(1).Caption = "&Turni Pazienti"
          mnuSottoDialisiTurni(1).Caption = "&Associa Turni/Reni"
          mnuSottoDialisi(2).Caption = "&Seduta Dialitica Giornaliera"
          mnuSottoDialisiScheda(1).Caption = "&Compilazione"
          mnuSottoDialisiScheda(2).Caption = "C&onsultazione"
          mnuSottoDialisi(3).Caption = "S&eduta Supplementare"
          mnuSottoDialisi(4).Caption = "&Piano di Lavoro"
          mnuSottoDialisi(5).Caption = "C&onsumi e Previsioni"
          
    mnuGestioneIndicatori.Caption = "Gestione &Indicatori"
          mnuGestioneIndicatoriSotto(1).Caption = "&Calcolo Kt/V"
          mnuGestioneIndicatoriSotto(2).Caption = "C&alcolo TSAT %"
          mnuGestioneIndicatoriSotto(3).Caption = "Calcolo Pr&odotto Ca / P"
          mnuGestioneIndicatoriSotto(4).Caption = "&Eventi"
          mnuGestioneIndicatoriSotto(5).Caption = "&Colture"
          mnuGestioneIndicatoriSotto(6).Caption = "E&ritropoietina per Paziente"
          mnuGestioneIndicatoriSotto(7).Caption = "&Monitoraggi"
          mnuGestioneIndicatoriSotto(8).Caption = "&Trattamento Acque"
          mnuGestioneIndicatoriSotto(9).Caption = "&Pazienti Candidati al Trapianto"
          mnuGestioneIndicatoriSotto(10).Caption = "Esami Periodici in &ED"
          mnuGestioneIndicatoriSotto(11).Caption = "&Scheda Rilevazione FAV"
          
    mnuArchivi.Caption = "&Setup Tabelle"
          mnuSottoTab(1).Caption = "&Organigramma"
          mnuSottoTabOrgan(2).Caption = "&Medici in Dialisi"
          mnuSottoTabOrgan(3).Caption = "&Infermieri"
          mnuSottoTabOrgan(1).Caption = "&Direttore Sanitario"
          mnuSottoTab(2).Caption = "Medici &Refertanti"
          mnuSottoTab(3).Caption = "Medici di &Base"
          mnuSottoTab(4).Caption = "&Psicologi"
          mnuSottoTab(5).Caption = "Or&gano/apparato"
          mnuSottoTab(6).Caption = "Esami &Strumentali per Organo/Apparato"
          mnuSottoTab(7).Caption = "&Filtri"
          mnuSottoTab(8).Caption = "&Linee"
          mnuSottoTab(9).Caption = "&Aghi"
          mnuSottoTab(10).Caption = "Far&maci in uso"
          mnuSottoTab(11).Caption = "A&nticoagulanti"
          mnuSottoTab(12).Caption = "&Titoli Diario Clinico"
          mnuSottoTab(13).Caption = "&Esami di Laboratorio"
          mnuSottoTab(14).Caption = "Raggr&uppamento Esami di Laboratorio"
          mnuSottoTab(15).Caption = "Co&dici Era - E.D.T.A."
          mnuSottoTab(16).Caption = "Codici E.D.T.A. - Causa Morte"
          
    mnuStrumenti.Caption = "&Strumenti"
          mnuGesPass.Caption = "&Gestione Utenti"
          mnuImpostaStampa.Caption = "&Intestazione Centro"
          mnuIntestazioneFattura.Caption = "Parametri Fattura"
          mnuRipristina.Caption = "&Ripristino Archivi"
          mnuImpostaBackup.Caption = "&N° Backup"
          mnuEsportaDb.Caption = "&Esporta Database"
     '     mnuBarra.Caption = "&Barra degli Strumenti"
     
    mnuStampe.Caption = "Stam&pe"
          mnuStampaPaz.Caption = "Lista &Pazienti"
          mnuStampaMediciBase.Caption = "Lista &Medici di Base"
          mnuMostraFattElaborazione.Caption = "&Visualizza Giorni Dialisi"
          mnuImpegnativeDialisi.Caption = "Richieste &Impegnative Dialisi"
          mnuEtichettePerProvetta.Caption = "&Etichette per Provette"
          mnuModuloFirmePaziente.Caption = "&Modulo Firme Pazienti"
          mnuKtvAnnuale.Caption = "&KT/V Annuale"
          mnuTsatAnnuale.Caption = "&TSAT %  Annuale"
          mnuSchedaDialiticaSettimanale.Caption = "&Scheda Dialitica Settimanale"
          
    mnuFatturazione.Caption = "&Fatturazione"
        mnuTabelleFatturazione.Caption = "&Tabelle"
        mnuTabFatt(0).Caption = "&Regioni"
        mnuTabFatt(1).Caption = "&Comuni"
        mnuTabFatt(2).Caption = "&Asl"
        mnuTabFatt(3).Caption = "&Distretti"
        mnuTabFatt(4).Caption = "&Esenzioni"
        mnuTabFatt(5).Caption = "&Nomenclatore Tariffario"
        mnuTabFatt(6).Caption = "Acc&ompagnatori"
        mnuCaricaPrescrizione.Caption = "&Gestione Ricette"
        mnuGestioneFileC.Caption = "&Genera File C"
        mnuGestioneFileXml.Caption = "&Genera File XML"
        mnuFattStampaFogli.Caption = "S&tampa Fogli di Viaggio"
        mnuRimborsi.Caption = "St&ampa Rimborsi"
        mnuStampaRiepiloghi.Caption = "&Stampa Riepiloghi"
        mnuStampaRiepilogo(0).Caption = "&Fattura"
        mnuStampaRiepilogo(1).Caption = "x &Paziente"
        mnuStampaRiepilogo(2).Caption = "x &Totali - Prestazioni"
        mnuStampaRiepilogo(3).Caption = "x Totali - A&sl"
        mnuStampaRiepilogo(4).Caption = "x T&otali - Mazzette x Distretti"
        mnuStampaRiepilogo(5).Caption = "x &Mazzette - Mensili"
        mnuStampaRiepilogo(6).Caption = "x Ma&zzetta - Singola"
        mnuStampaRiepilogo(7).Caption = "x &Asl - Distretti"
        mnuStampaRiepilogo(8).Caption = "x &Impegnative"
        
    mnuVassoio.Caption = "PopupVassoio"
        mnuApriVassoio.Caption = "Ripristina"
        
    mnuApparati.Caption = "Apparati"
        mnuSottoApparati(0).Caption = "Gestione &Apparati"
        mnuSottoApparati(1).Caption = "Stampa Re&gistro"
        mnuSottoApparati(2).Caption = "Parco Ren&i Artificiali"
        
    mnu1.Caption = "&?"
        mnunoterilascio.Caption = "&Note di Rilascio"
        mnulicenza.Caption = "&Licenza d'uso"
        mnuabout.Caption = "&Informazioni su Isodial..."
          
    '/ Subclassing menu. Warning! Don't remove this comment!!!

    Set MenuEvents = New CEvents
    Set objMenuEx = New cMenuEx
    Call objMenuEx.Install(Me.hWnd, App.Path & "\" & Me.Name, ImageList1, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
    '/ Open Menu Designer tool
    objMenuEx.MenuDesigner Me.hWnd
End Sub

