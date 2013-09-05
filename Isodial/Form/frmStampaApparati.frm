VERSION 5.00
Begin VB.Form frmStampaApparati 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stampa Apparati "
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5595
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CheckBox chkIncludiApparatiRottamati 
         Caption         =   "Includi Apparati Rottamati"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   1455
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
         TabIndex        =   7
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton optTipoApparato 
         Caption         =   "Elenco per categoria apparato"
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
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton optProduttore 
         Caption         =   "Elenco per produttore"
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
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optInventario 
         Caption         =   "Elenco per n° inventario"
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
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraPulsanti 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2040
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

Private Sub cmdAvanti_Click()
    
    If optInventario.Value = False And optProduttore.Value = False And optTipoApparato.Value = False And optApparatiRottamati.Value = False Then
        MsgBox "Selezionare il tipo di elenco da stampare", vbInformation, "Informazione"
        Exit Sub
    End If
    
    If optInventario.Value = True Then
        If chkIncludiApparatiRottamati.Value = Checked Then
            strSql = "SELECT * FROM APPARATI ORDER BY NUMERO_INVENTARIO"
        Else
            strSql = "SELECT * FROM APPARATI WHERE IsNull(DATA_ROTTAMAZIONE) ORDER BY NUMERO_INVENTARIO"
        End If
        TipoElenco = "Elenco Apparati per N° Inventario:"
        Call StampaApparato
    
    ElseIf optProduttore.Value = True Then
        If chkIncludiApparatiRottamati.Value = Checked Then
            strSql = "SELECT * FROM APPARATI ORDER BY PRODUTTORE"
        Else
            strSql = "SELECT * FROM APPARATI WHERE IsNull(DATA_ROTTAMAZIONE) ORDER BY PRODUTTORE"
        End If
        TipoElenco = "Elenco Apparati per Produttore:"
        Call StampaApparato
        
    ElseIf optTipoApparato.Value = True Then
        If chkIncludiApparatiRottamati.Value = Checked Then
            strSql = "SELECT * FROM APPARATI ORDER BY TIPO_APPARATO"
        Else
            strSql = "SELECT * FROM APPARATI WHERE IsNull(DATA_ROTTAMAZIONE) ORDER BY TIPO_APPARATO"
        End If
        TipoElenco = "Elenco Apparati per Tipo Apparato:"
        Call StampaApparato
        
    ElseIf optApparatiRottamati.Value = True Then
        strSql = "SELECT * FROM APPARATI WHERE DATA_ROTTAMAZIONE Is Not Null ORDER BY KEY"
        TipoElenco = "Elenco Apparati Rottamati:"
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

Private Sub optApparatiRottamati_Click()
    chkIncludiApparatiRottamati.Value = Unchecked
End Sub
