VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmModuliWord 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Moduli prestampati"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid flxGriglia 
         Height          =   1935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         _Version        =   393216
         FormatString    =   "| Nome file                                                                                                          "
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6255
      Begin VB.CommandButton cmdWord 
         Caption         =   "&Word"
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
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdModifica 
         Caption         =   "&Modifica"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdStampa 
         Caption         =   "&Associa e Stampa"
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1620
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
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmModuliWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsModuli As Recordset
Dim codicePaziente As Integer

Public Property Get getCodicePaziente() As Integer
    getCodicePaziente = codicePaziente
End Property

Public Property Let LetCodicePaziente(ByVal vCodicePaziente As Integer)
    codicePaziente = vCodicePaziente
End Property

Private Sub CalcolaEta(codice As Integer)
    Dim rsDataset As New Recordset
    Dim somma As Integer
    
    rsDataset.Open "SELECT * FROM PAZIENTI WHERE KEY=" & codice, cnPrinc, adOpenKeyset, adLockPessimistic, adCmdText
    If Month(rsDataset("DATA_NASCITA")) > Month(date) Then
        somma = -1
    ElseIf Month(rsDataset("DATA_NASCITA")) = Month(date) And Day(rsDataset("DATA_NASCITA")) > Day(date) Then
        somma = -1
    Else
        somma = 0
    End If
    rsDataset("ETA") = Year(date) - Year(rsDataset("DATA_NASCITA")) + somma
    rsDataset.Update
    rsDataset.Close
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdModifica_Click()
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare il modulo da modificare", vbCritical, "Attenzione"
    Else
        ShellExecute Me.hWnd, "open", structApri.pathExe & "\Moduli\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc", "", "", 5
    End If
End Sub

Private Sub cmdStampa_Click()
    Dim i As Integer
    Dim nome As String
    Dim pos As String
    
    Dim wApp As Object
    Dim wDoc As Object
    
    If flxGriglia.Row = 0 Then
        MsgBox "Selezionare il modulo da stampare", vbCritical, "Attenzione"
    Else
        Call CalcolaEta(codicePaziente)
        
        Set rsModuli = New Recordset
        rsModuli.Open "SELECT * FROM " & structApri.strFromModuliWord & " WHERE P.KEY=" & codicePaziente, cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not (rsModuli.EOF And rsModuli.BOF) Then
            
            Set wApp = CreateObject("Word.Application")
            Set wDoc = wApp.Documents.Open(structApri.pathExe & "\Moduli\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc")
            For i = 1 To wDoc.FormFields.count
                nome = wDoc.FormFields(i).Name
                pos = InStr(1, nome, "9")
                If pos Then
                    Mid(nome, pos, 1) = "."
                End If
                wDoc.FormFields(i).Result = rsModuli(nome)
            Next i
            
            If Dir(Environ("TEMP") & "\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc") <> "" Then
                Kill Environ("TEMP") & "\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc"
            End If
            wDoc.SaveAs Environ("TEMP") & "\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc"
            
            ShellExecute Me.hWnd, "open", Environ("TEMP") & "\" & flxGriglia.TextMatrix(flxGriglia.Row, 1) & ".doc", "", "", 5
        Else
            MsgBox "Errore nel caricamento dei dati", vbCritical, "Errore"
            Exit Sub
        End If
        rsModuli.Close
        
        Set rsModuli = Nothing
    End If
End Sub

Private Sub cmdWord_Click()
    Dim w As Object
    Set w = CreateObject("Word.Application")
    If Dir(w.Path & "\winword.exe") <> "" Then
        ShellExecute Me.hWnd, "open", w.Path & "\winword.exe", "", "", 5
    ElseIf Dir(w.Path & "\word.exe") <> "" Then
        ShellExecute Me.hWnd, "open", w.Path & "\word.exe", "", "", 5
    Else
        MsgBox "Software Microsoft Office Word non trovato", vbCritical, "Attenzione"
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

Private Sub Form_Load()
    Set rsModuli = New Recordset
    
    With flxGriglia
        .Rows = 1
        .ColWidth(0) = 0
        .Row = 0
        .MousePointer = flexCustom
        .Col = 1
        .CellFontBold = True
    End With
    
    rsModuli.Open "MODULI_PRESTAMPATI", cnPrinc, adOpenForwardOnly, adLockReadOnly, adCmdTable
    Do While Not rsModuli.EOF
        With flxGriglia
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rsModuli("KEY")
            .TextMatrix(.Rows - 1, 1) = rsModuli("NOME_FILE")
        End With
        rsModuli.MoveNext
    Loop
    rsModuli.Close
    Set rsModuli = Nothing
End Sub
