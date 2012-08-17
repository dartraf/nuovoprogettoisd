VERSION 5.00
Begin VB.Form frmStatoPaz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stato Paziente: "
   ClientHeight    =   7500
   ClientLeft      =   2040
   ClientTop       =   1260
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOspiti 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   7695
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   6
         Left            =   7200
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   29
         ToolTipText     =   "Cerca data"
         Top             =   1545
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   5
         Left            =   7200
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   28
         ToolTipText     =   "Cerca data"
         Top             =   1035
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   4
         Left            =   7200
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   27
         ToolTipText     =   "Cerca data"
         Top             =   525
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   3
         Left            =   1440
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   26
         ToolTipText     =   "Cerca data"
         Top             =   1545
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   2
         Left            =   1440
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   25
         ToolTipText     =   "Cerca data"
         Top             =   1035
         Width           =   360
      End
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   1440
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   24
         ToolTipText     =   "Cerca data"
         Top             =   525
         Width           =   360
      End
      Begin VB.ComboBox cboCentroProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cboCentroProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ComboBox cboCentroProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label lblDataPa 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Arrivo"
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
         TabIndex        =   21
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveniente Da"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Partenza"
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
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblDataAr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDataAr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblDataAr 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblDataPa 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDataPa 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame fraData 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5055
      Begin VB.PictureBox picData 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   4320
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   23
         ToolTipText     =   "Cerca data"
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblDataNome 
         Caption         =   "Data "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2145
      End
   End
   Begin VB.Frame fraDonatore 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   5055
      Begin VB.OptionButton optDonatore 
         Caption         =   "Vivente"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optDonatore 
         Caption         =   "Cadavere"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Donatore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   2025
      End
   End
   Begin VB.Frame fraOpzioni 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5040
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
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdConferma 
         Caption         =   "&Conferma"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmStatoPaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    
    lblData.BackColor = vbWhite
    For i = 0 To 2
        lblDataAr(i).BackColor = vbWhite
        lblDataPa(i).BackColor = vbWhite
    Next i
    For i = 0 To 6
        picData(i).Picture = LoadResPicture("cal1", 0)
    Next i
    
    Select Case statoPaziente.statoPaz
        Case TPDECEDUTO
            fraData.Top = 0
            fraOpzioni.Top = 600
            fraOpzioni.ZOrder 1
            Me.Height = fraOpzioni.Height + fraOpzioni.Top + 450
            Me.Width = fraData.Width + 300
            Me.Caption = Me.Caption & "deceduto"
            lblDataNome = lblDataNome & "decesso"
        Case TPTRAPIANTO
            fraData.Top = 0
            fraDonatore.Top = 600
            fraOpzioni.Top = 1200
            fraOpzioni.ZOrder 1
            Me.Height = fraOpzioni.Height + fraOpzioni.Top + 450
            Me.Width = fraData.Width + 300
            Me.Caption = Me.Caption & "trapiantato"
            lblDataNome = lblDataNome & "trapiantato"
        Case TPOSPITE
            fraOspiti.Top = 0
            fraOpzioni.Top = 1920
            fraOpzioni.Width = fraOspiti.Width
            fraOpzioni.ZOrder 1
            Me.Height = fraOpzioni.Height + fraOpzioni.Top + 450
            Me.Width = fraOspiti.Width + 300
            Me.Caption = Me.Caption & "ospite"
            For i = 0 To 2
                Call RicaricaComboBox("CENTRI_PROVENIENZA", "NOME", cboCentroProv(i))
            Next
        Case TPTRASFERITO
            fraData.Top = 0
            fraOpzioni.Top = 600
            fraOpzioni.ZOrder 1
            Me.Height = fraOpzioni.Height + fraOpzioni.Top + 450
            Me.Width = fraData.Width + 300
            Me.Caption = Me.Caption & "trasferito"
            lblDataNome = lblDataNome & "trasferimento"
    End Select
    If frmPaziente.intPazientiKey <> 0 Then
        ' sta visualizzando le info su un paziente
        Select Case statoPaziente.statoPaz
            Case TPOSPITE
                For i = 0 To 2
                    lblDataAr(i) = statoPaziente.dataArrivi(i + 1)
                    lblDataPa(i) = statoPaziente.dataPartenza(i + 1)
                    cboCentroProv(i).ListIndex = GetCboListIndex(statoPaziente.centriProv(i + 1), cboCentroProv(i))
                Next i
            Case Else
                lblData = statoPaziente.dataStato
                If statoPaziente.donatore <> 2 Then
                    optDonatore(statoPaziente.donatore).Value = True
                End If
        End Select
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub cmdConferma_Click()
    Dim i As Integer
    If Completo Then
        For i = 0 To 2
            If cboCentroProv(i).Text <> "" Then
                Call GestisciNuovo("CENTRI_PROVENIENZA", cboCentroProv(i))
            End If
        Next i
        Call AnnullaVarStato
        Select Case statoPaziente.statoPaz
            Case TPOSPITE
                For i = 0 To 2
                    If lblDataAr(i) <> "" Then
                        statoPaziente.dataArrivi(i + 1) = lblDataAr(i)
                    End If
                    If lblDataPa(i) <> "" Then
                        statoPaziente.dataPartenza(i + 1) = lblDataPa(i)
                    End If
                    If cboCentroProv(i).ListIndex = -1 Then
                        statoPaziente.centriProv(i + 1) = -1
                    Else
                        statoPaziente.centriProv(i + 1) = cboCentroProv(i).ItemData(cboCentroProv(i).ListIndex)
                    End If
                Next i
            Case TPTRAPIANTO
                If lblData <> "" Then
                    statoPaziente.dataStato = lblData
                End If
                statoPaziente.donatore = IIf(optDonatore(0).Value = False And optDonatore(1).Value = False, 2, IIf(optDonatore(0).Value = True, 0, 1))
            Case Else
                If lblData <> "" Then
                    statoPaziente.dataStato = lblData
                End If
        End Select
        Unload Me
    End If
End Sub

Private Function Completo() As Boolean
    Dim i As Integer
    Completo = True
    For i = 0 To 2
        If lblDataAr(i) <> "" Or lblDataPa(i) <> "" Then
            If Not (lblDataAr(i) <> "" And lblDataPa(i) <> "" And cboCentroProv(i).Text <> "") Then
                Completo = False
                MsgBox "Inserire correttamente tutti i valori richiesti", vbCritical, "Attenzione"
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub lblData_Click()
    laData = ""
    lblData = ""
End Sub

Private Sub lblDataAr_Click(Index As Integer)
    laData = ""
    lblDataAr(Index) = ""
End Sub

Private Sub lblDataPa_Click(Index As Integer)
    laData = ""
    lblDataPa(Index) = ""
End Sub

Private Sub optDonatore_Click(Index As Integer)
    Call ColoraSel(optDonatore, Index, 2)
End Sub

Private Sub picData_Click(Index As Integer)
    frmCalendario.Show 1
    If laData <> "" Then
        If Index = 0 Then
            lblData = laData
        ElseIf Index >= 1 And Index <= 3 Then
            lblDataAr(Index - 1) = laData
        Else
            lblDataPa(Index - 4) = laData
        End If
    End If
End Sub

Private Sub picData_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal2", 0)
End Sub

Private Sub picData_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picData(Index).Picture = LoadResPicture("cal1", 0)
End Sub

