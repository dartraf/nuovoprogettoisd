VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'########################################################################################
'#                                                                                      #
'#    cTooltip.cls  ver 1.01                                                            #
'#    classe per la creazione e gestione di messaggi tooltip multilinea                 #
'#                                                                                      #
'########################################################################################


Option Explicit

Private Const WM_USER = &H400
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_SETTITLE = (WM_USER + 32)


Private Enum TT_FLAGS
    TTF_IDISHWND = &H1
    TTF_SUBCLASS = &H10
End Enum

Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_RESHOW = 1
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Enum TT_Msgs

    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_GETTOOLCOUNT = (WM_USER + 13)

    #If UNICODE Then
        TTM_ADDTOOL = (WM_USER + 50)
        TTM_DELTOOL = (WM_USER + 51)
        TTM_UPDATETIPTEXT = (WM_USER + 57)
        TTM_ENUMTOOLS = (WM_USER + 58)

    #Else
        TTM_ADDTOOL = (WM_USER + 4)
        TTM_DELTOOL = (WM_USER + 5)
        TTM_UPDATETIPTEXT = (WM_USER + 12)
        TTM_ENUMTOOLS = (WM_USER + 14)

    #End If
    TTM_SETTIPBKCOLOR = (WM_USER + 19)
    TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    TTM_GETDELAYTIME = (WM_USER + 21)
    TTM_GETTIPBKCOLOR = (WM_USER + 22)
    TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    TTM_GETMAXTIPWIDTH = (WM_USER + 25)

End Enum

Enum Style
    WinStandard = 0
    Fumetto = 1
End Enum

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As TT_FLAGS
    hwnd As Long
    uId As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type


Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long
                            
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
                            (ByVal dwExStyle As Long, ByVal lpClassName As String, _
                             ByVal lpWindowName As String, ByVal dwStyle As Long, _
                             ByVal x As Long, ByVal Y As Long, _
                             ByVal nWidth As Long, ByVal nHeight As Long, _
                             ByVal hwndParent As Long, ByVal hMenu As Long, _
                             ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                       (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
                       (Destination As Any, ByVal Length As Long)


Private m_hwndTT As Long
Private m_cMaxTip As Long
Private m_Creata As Boolean
Private m_Style As Style
Private m_Title As String
Private m_Icon As Long
Private m_DelayTime As Long
Private m_VisibleTime As Long
Private TI As TOOLINFO

'##########################################################################################
'#                                                                                        #
'#                                INTERFACCIA PUBBLICA                                    #
'#                                                                                        #
'#  METODI                                                                                #
'#     Create: Parametri: frm as Form: Identificativo Windows (Handle) del form in cui    #
'#                                     l'oggetto viene visualizzato                       #
'#             Valori rest. : Boolean True se l'oggetto viene creato                      #
'#     AddTool: Parametri: CTRL As Control: riferimento al controllo su cui l'oggetto     #
'#                                          verrà visualizzato                            #
'#                         [sText As String]: Opzionale testo del tooltip                 #
'#              Valori rest. : Boolean True se il tooltip viene aggiunto                  #
'#     RemoveTool: Parametri: CTRL As Control: riferimento al controllo su cui l'oggetto  #
'#                                             verrà visualizzato                         #
'#                 Valori rest. : Boolean True se il tooltip viene rimosso                #
'#  PROPRIETA'                                                                            #
'#     BackColor: imposta e restituisce il colore di sfondo del tooltip                   #
'#     DelayTime: imposta e restituisce il tempo che deve intercorrere prima che l'oggetto#
'#                venga visualizzato in millisecondi                                      #
'#     ForeColor: imposta e restituisce il colore del testo del tooltip                   #
'#     Hwnd: sola lettura Identificativo Windows (Handle) del tooltip                     #
'#     Icon: imposta e restituisce l'icona dell'oggetto                                   #
'#     MaxTipWidth: imposta e restituisce la larghezza massima in pixel                   #
'#     Title: imposta e restituisce il titolo dell'oggetto                                #
'#     ToolCount: sola lettura numero dei controlli a cui l'oggetto è associato           #
'#     ToolText: imposta e restituisce il testo dell'oggetto personalizzabile per ogni    #
'#               controllo a cui l'oggetto viene associato                                #
'#     VisibleTime: imposta e restituisce il tempo per il quale l'oggetto viene           #
'#                  visualizzato in millisecondi                                          #
'#                                                                                        #
'##########################################################################################

Private Sub Class_Initialize()
    '//prima di tutto pulisco la memoria riservata alla struttura
    ZeroMemory TI, Len(TI)
    If m_hwndTT <> 0 Then
        Call DestroyWindow(m_hwndTT)
    End If
    '//inizializzazione variabili
    m_Creata = False
    m_Style = Fumetto
    m_Title = " " '//importante che la stringa non sia vuota
    m_Icon = TTNoIcon
    m_DelayTime = 500
    m_VisibleTime = 20000
End Sub

Private Sub Class_Terminate()
Dim nItems  As Long
    If m_hwndTT Then
        nItems = ToolCount
        
        Do Until nItems = 0
            If SendMessage(m_hwndTT, TTM_ENUMTOOLS, (0), TI) Then
                RemoveToolHwnd TI.uId
            End If
            nItems = ToolCount
        Loop
        ZeroMemory TI, Len(TI)
        
        Call DestroyWindow(m_hwndTT)
    End If
End Sub

Public Function Create(frmhwnd As Long) As Boolean
On Error GoTo Hell
Dim lWinStyle As Long
Dim RES As Long

    RES = SendMessageLong(m_hwndTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, m_VisibleTime)
    RES = SendMessageLong(m_hwndTT, TTM_SETDELAYTIME, TTDT_INITIAL, m_DelayTime)

    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    If m_Style = Fumetto Then
        lWinStyle = lWinStyle Or TTS_BALLOON
    End If
    If (m_hwndTT = 0) Then
        Call InitCommonControls
        m_hwndTT = CreateWindowEx(0&, _
                        TOOLTIPS_CLASS, _
                        vbNullString, _
                        lWinStyle, _
                        0, _
                        0, _
                        0, _
                        0, _
                        frmhwnd, _
                        0, _
                        App.hInstance, _
                        0)
    End If
  
    m_Creata = CBool(m_hwndTT)
    If m_Creata = False Then
        Create = False
    Else
        Create = True
        RES = SendMessage(m_hwndTT, TTM_ADDTOOLA, 0&, TI)
        If Not m_Title = " " Then
            RES = SendMessage(m_hwndTT, TTM_SETTITLE, CLng(m_Icon), ByVal m_Title)
        End If
        
    End If
    
Exit Function
Hell:
    m_Creata = False
    Create = False
    Err.Clear
End Function

Public Function AddTool(CTRL As Control, frmhwnd As Long, Optional sText As String = "") As Boolean
Dim RES  As Long
    If m_Creata = False Then Exit Function
    If (m_hwndTT = 0) Then Exit Function
    If (GetToolInfo(CTRL.hwnd, TI) = False) Then
        
        With TI
            .cbSize = Len(TI)
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hwnd = frmhwnd
            .uId = CTRL.hwnd
            
            If Len(sText) Then
                .lpszText = sText
            Else
                .lpszText = ""
            End If
            m_cMaxTip = Max(m_cMaxTip, Len(.lpszText) + 1)
        End With
        AddTool = SendMessage(m_hwndTT, TTM_ADDTOOL, 0, TI)
        RES = SendMessageLong(m_hwndTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, m_VisibleTime)
        RES = SendMessageLong(m_hwndTT, TTM_SETDELAYTIME, TTDT_INITIAL, m_DelayTime)
    End If
    
End Function

Public Function RemoveTool(CTRL As Control) As Boolean
    
    If (m_hwndTT = 0) Then Exit Function
    If GetToolInfo(CTRL.hwnd, TI) Then
        Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, TI)
        RemoveTool = True
    End If

End Function


Public Property Let Title(ByVal vData As String)
    If vData = "" Then vData = " "
    m_Title = vData
    SendMessage m_hwndTT, TTM_SETTITLE, CLng(m_Icon), ByVal m_Title
End Property

Public Property Get Title() As String
   Title = m_Title
End Property

Public Property Get Icon() As ttIconType
    Icon = m_Icon
End Property

Public Property Let Icon(ByVal vData As ttIconType)
    m_Icon = vData
    If m_Title = "" Then m_Title = " "
    SendMessage m_hwndTT, TTM_SETTITLE, CLng(m_Icon), ByVal m_Title

End Property

Public Property Get BackColor() As OLE_COLOR
    If (m_hwndTT = 0) Then Exit Property
    BackColor = SendMessage(m_hwndTT, TTM_GETTIPBKCOLOR, 0, 0)
End Property

Public Property Let BackColor(clr As OLE_COLOR)
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_SETTIPBKCOLOR, clr, 0)
End Property

Public Property Get DelayTime() As Long
    DelayTime = m_DelayTime
End Property

Public Property Let DelayTime(dwMilliSecs As Long)
    If m_hwndTT = 0 Then
        m_DelayTime = dwMilliSecs
        Call SendMessage(m_hwndTT, TTM_SETDELAYTIME, TTDT_INITIAL, m_DelayTime)
    Else
        '// la finestra è gia stata creata non è possibile  modificare
        '// il tempo di delay
        
    End If
End Property

Public Property Get VisibleTime() As Long
    VisibleTime = m_VisibleTime
End Property

Public Property Let VisibleTime(dwMilliSecs As Long)

    m_VisibleTime = dwMilliSecs
    SendMessage m_hwndTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, m_DelayTime
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    If (m_hwndTT = 0) Then Exit Property
    ForeColor = SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)
End Property

Public Property Let ForeColor(clr As OLE_COLOR)
    If (m_hwndTT = 0) Then Exit Property
    Call SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)
End Property

Public Property Get hwnd() As Long   ' read-only
    hwnd = m_hwndTT
End Property

Public Property Get ToolCount() As Integer   ' read-only
    If (m_hwndTT = 0) Then Exit Property
    ToolCount = SendMessage(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Get MaxTipWidth() As Integer
    If (m_hwndTT = 0) Then Exit Property
    MaxTipWidth = LOWORD(SendMessage(m_hwndTT, TTM_GETMAXTIPWIDTH, 0, 0))
End Property

Public Property Let MaxTipWidth(ByVal cx As Integer)
    If (m_hwndTT = 0) Then Exit Property
    If (cx < 1) Then cx = -1
    Call SendMessage(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal CLng(cx))
End Property

Public Property Get ToolText(CTRL As Control) As String

    If (m_hwndTT = 0) Then Exit Property
    If GetToolInfo(CTRL.hwnd, TI, True) Then
        ToolText = GetStrFromBufferA(TI.lpszText)
    End If

End Property

Public Property Let ToolText(CTRL As Control, sText As String)
  
    If (m_hwndTT = 0) Then Exit Property
    If GetToolInfo(CTRL.hwnd, TI) Then
        TI.lpszText = sText
        m_cMaxTip = Max(m_cMaxTip, Len(sText) + 1)
        Call SendMessage(m_hwndTT, TTM_UPDATETIPTEXT, 0, TI)
    End If

End Property


'########################################################################################
'#                                                                                      #
'#                              INTERFACCIA PRIVATA                                     #
'#                                                                                      #
'########################################################################################

Private Function IsWindow(CTRL As Control) As Boolean
On Error GoTo Out
    IsWindow = CBool(CTRL.hwnd)
Out:
End Function

Private Function LOWORD(dwValue As Long) As Integer
    MoveMemory LOWORD, dwValue, 2
End Function

Private Function Max(param1 As Long, param2 As Long) As Long
    If param1 > param2 Then Max = param1 Else Max = param2
End Function

Private Function GetStrFromBufferA(szA As String) As String
    If InStr(szA, vbNullChar) Then
        GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
    Else
        GetStrFromBufferA = szA
    End If
End Function

Private Function GetToolInfo(hwndTool As Long, _
    TI As TOOLINFO, _
    Optional fGetText As Boolean = False) As Boolean
Dim nItems As Integer
Dim i As Integer
  
    TI.cbSize = Len(TI)
    If fGetText Then TI.lpszText = String$(m_cMaxTip, 0)
      
    nItems = ToolCount
    
    For i = 0 To nItems - 1
        If SendMessage(m_hwndTT, TTM_ENUMTOOLS, (i), TI) Then
            If (hwndTool = TI.uId) Then
                GetToolInfo = True
                Exit Function
            End If
        End If
    Next

End Function

Private Function RemoveToolHwnd(CTRLHwnd As Long) As Boolean
    If (m_hwndTT = 0) Then Exit Function
    If GetToolInfo(CTRLHwnd, TI) Then
        Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, TI)
        RemoveToolHwnd = True
    End If
End Function

