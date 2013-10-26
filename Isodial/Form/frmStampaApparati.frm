VERSION 5.00
Begin VB.Form frmStampaApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Apparati "
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStampa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdTrovaCategoria 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   1395
         Picture         =   "frmStampaApparati.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   450
      End
      Begin VB.OptionButton optTuttiProduttori 
         Caption         =   "Elenco per tutti i produttori"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CommandButton cmdTrova 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Left            =   1400
         Picture         =   "frmStampaApparati.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   450
      End
      Begin VB.OptionButton optApparatiRottamati 
         Caption         =   "Elenco apparati rottamati"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   3615
      End
      Begin VB.OptionButton optTipoApparato 
         Caption         =   "Elenco per tutte le categorie"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   3375
      End
      Begin VB.OptionButton optInventario 
         Caption         =   "Elenco per n° d'inventario"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label lblNomeCategoria 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Elenco per Categoria"
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
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Elenco per Produttore"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblNomeProduttore 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   5415
      Begin VB.CommandButton cmdEsci 
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
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdAvanti 
         Cancel          =   -1  'True
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
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmStampaApparati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql As String
Dim TipoElenco As String

Private Sub cmdTrovaCategoria_Click()
    
    optApparatiRottamati.Value = Unchecked
    optInventario.Value = Unchecked
    optTipoApparato.Value = Unchecked
    optTuttiProduttori.Value = Unchecked
    lblNomeProduttore.Caption = ""

    'Azzero la variabile per evitare di ricaricare lo stesso dato
    tTrova.NomeStriga = ""
    
    tTrova.Tipo = tpAPPARATI_TIPO
    tTrova.condizione = ""
    tTrova.condStato = ""
    Unload frmTrova
    frmTrova.Show 1
    lblNomeCategoria.Caption = tTrova.NomeStriga

End Sub

Private Sub optTuttiProduttori_Click()
    lblNomeProduttore.Caption = ""
    lblNomeCategoria.Caption = ""
End Sub

Private Sub cmdAvanti_Click()
Dim NomeProduttore As String
Dim NomeCategoria As String
    
    If optInventario.Value = False And optTuttiProduttori.Value = Unchecked And optTipoApparato.Value = False And optApparatiRottamati.Value = False And lblNomeProduttore.Caption = "" And lblNomeCategoria.Caption = "" Then
        MsgBox "Selezionare il tipo di elenco da stampare", vbInformation, "Informazione"
        Exit Sub
    End If
    
    If optInventario.Value = True Then
        strSql = "SELECT * FROM APPARATI ORDER BY NUMERO_INVENTARIO"
        TipoElenco = "Elenco Apparati per N° Inventario"
        Call StampaApparato
        
    ElseIf optApparatiRottamati.Value = True Then
        strSql = "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE Is Not Null ORDER BY NUMERO_INVENTARIO"
        TipoElenco = "Elenco Apparati Rottamati"
        Call StampaApparato
        
    ' Stampe per Produttore
    ElseIf optTuttiProduttori.Value = True Then
        strSql = "SELECT * FROM APPARATI ORDER BY PRODUTTORE"
        TipoElenco = "Elenco Apparati per Produttore"
        Call StampaApparato
        
    ElseIf lblNomeProduttore.Caption <> "" Then
        NomeProduttore = lblNomeProduttore.Caption
        strSql = "SELECT * FROM APPARATI WHERE PRODUTTORE='" & NomeProduttore & "'" & "ORDER BY NUMERO_INVENTARIO"
        TipoElenco = "Elenco Apparati per Produttore"
        Call StampaApparato
        
    ' Stampe per Categoria
    ElseIf optTipoApparato.Value = True Then
        strSql = "SELECT * FROM APPARATI ORDER BY TIPO_APPARATO"
        TipoElenco = "Elenco per Categoria Apparati"
        Call StampaApparato
        
    ElseIf lblNomeCategoria.Caption <> "" Then
        NomeCategoria = lblNomeCategoria.Caption
        strSql = "SELECT * FROM APPARATI WHERE TIPO_APPARATO='" & NomeCategoria & "'" & "ORDER BY NUMERO_INVENTARIO"
        TipoElenco = "Elenco per Categoria Apparati"
        Call StampaApparato
        
    End If
    
End Sub

Private Sub StampaApparato()
    Dim SQLString As String
    Dim cnConn As Connection        ' connessione per lo shape
    Dim rsMain As Recordset         ' recordset padre per lo shape
    Dim rsDataset As Recordset
    Dim TotaleReni As Integer
    
    SQLString = "SHAPE APPEND " & _
                "       NEW adVarChar(4) AS NUMERO_INVENTARIO, " & _
                "       NEW adVarChar(4) AS NUMERO_APPARATO, " & _
                "       NEW adVarChar(4) AS POSTAZIONE, " & _
                "       NEW adVarChar(50) AS TIPO_APPARATO, " & _
                "       NEW adVarChar(50) AS MODELLO, " & _
                "       NEW adVarChar(50) AS MATRICOLA, " & _
                "       NEW adVarChar(50) AS PRODUTTORE, " & _
                "       NEW adDate AS DATA_FABBRICAZIONE, " & _
                "       NEW adDate AS DATA_COLLAUDO, " & _
                "       NEW adDate AS DATA_ROTTAMAZIONE "
                
        
    ' apre la connessione per lo shape
    Set cnConn = New ADODB.Connection
    cnConn.Open "Data Provider=NONE; Provider=MSDataShape"
    Set rsMain = New ADODB.Recordset
    rsMain.Open SQLString, cnConn, adOpenStatic, adLockOptimistic
    
    Set rsDataset = New Recordset
    
    rsDataset.Open strSql, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsDataset.EOF And rsDataset.BOF) Then
        With rsMain
            Do While Not rsDataset.EOF
                .AddNew
                .Fields("NUMERO_INVENTARIO") = rsDataset("NUMERO_INVENTARIO")
                .Fields("NUMERO_APPARATO") = rsDataset("NUMERO_APPARATO")
                .Fields("POSTAZIONE") = rsDataset("POSTAZIONE")
                .Fields("TIPO_APPARATO") = rsDataset("TIPO_APPARATO")
                .Fields("MODELLO") = rsDataset("MODELLO")
                .Fields("MATRICOLA") = rsDataset("MATRICOLA")
                .Fields("PRODUTTORE") = rsDataset("PRODUTTORE")
                .Fields("DATA_FABBRICAZIONE") = rsDataset("DATA_FABBRICAZIONE")
                .Fields("DATA_COLLAUDO") = rsDataset("DATA_COLLAUDO")
                .Fields("DATA_ROTTAMAZIONE") = rsDataset("DATA_ROTTAMAZIONE")
                rsDataset.MoveNext
            Loop
        End With
    End If
    
    If rsDataset.RecordCount > 0 Then
            TotaleReni = rsDataset.RecordCount
        Else
            MsgBox "Non sono presenti Apparati da Stampare", vbInformation, "Informazione"
            Exit Sub
    End If
    
    Set rsDataset = Nothing
    
    Set rptStampaApparati.DataSource = rsMain
    rptStampaApparati.Orientation = rptOrientLandscape
    rptStampaApparati.TopMargin = 0
    rptStampaApparati.RightMargin = 0
    rptStampaApparati.LeftMargin = 0
    rptStampaApparati.Sections("Intestazione").Controls("lblElenco").Caption = TipoElenco
    rptStampaApparati.Sections("Section5").Controls.Item("lblTotaleReni").Caption = TotaleReni
    rptStampaApparati.PrintReport True, rptRangeAllPages

End Sub

Private Sub cmdEsci_Click()
    Unload frmStampaApparati
End Sub

Private Sub cmdTrova_Click()
    
    optApparatiRottamati.Value = Unchecked
    optInventario.Value = Unchecked
    optTipoApparato.Value = Unchecked
    optTuttiProduttori.Value = Unchecked
    lblNomeCategoria.Caption = ""
    
    'La variabile StampaApparati è vera in modo tale che
    'quando carico il formTrova mi carica
    '1)il cmdNuovo e cmdModifica non visibili
    '2)il nome della stringa selezionata
    StampaApparati = True
        
    'Azzero la variabile per evitare di ricaricare lo stesso dato
    tTrova.NomeStriga = ""
    
    tTrova.Tipo = tpPRODUTTORE_MANUTENTORE
    tTrova.condizione = ""
    tTrova.condStato = ""
    Unload frmTrova
    frmTrova.Show 1
    lblNomeProduttore.Caption = tTrova.NomeStriga
    
    StampaApparati = False

End Sub

Private Sub optApparatiRottamati_GotFocus()
    lblNomeProduttore.Caption = ""
    lblNomeCategoria.Caption = ""
End Sub

Private Sub optInventario_GotFocus()
    lblNomeProduttore.Caption = ""
    lblNomeCategoria.Caption = ""
End Sub

Private Sub optTipoApparato_GotFocus()
    lblNomeProduttore.Caption = ""
    lblNomeCategoria.Caption = ""
End Sub
