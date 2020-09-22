Attribute VB_Name = "NewMsgBox"
Option Explicit

Private Formhwnd As Long
Private Buttonhwnd As Long
Private PicHwnd As Long
Private TextHwnd As Long
Private Iconhandle As Long
Private DlgX As Long
Private DLgY As Long
Private MyHook As Long
Private Nr As Long
Private WndStatic(1) As Long
Private WindowStil As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd&, ByVal lpClassName$, ByVal nMaxCount&)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const HCBT_CREATEWND = 3
Private Const HCBT_ACTIVATE = 5
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const DS_NOIDLEMSG = &H100
Private Const DS_ABSALIGN = &H1&
Private Const WM_GETTEXT = &HD
Public Const WS_SYSMENU = &H80000

Public Sub AboutBox(ByVal parent_hwnd As Long, ByVal Headertext As String, ByVal Message As String, ByVal icon As Long, Optional MidOfScreen As Boolean = True, Optional Stil As Long = 0, Optional x As Long = 0, Optional y As Long = 0)
Iconhandle = icon
Select Case MidOfScreen
Case True
DlgX = (Screen.Width / Screen.TwipsPerPixelX / 2) - 263
DLgY = (Screen.Height / Screen.TwipsPerPixelY / 2) - 67
Case False
DlgX = x
DLgY = y
End Select
WindowStil = Stil
Nr = 0
MyHook = SetWindowsHookEx(&H5, AddressOf MsgBoxHookProc, App.hInstance, App.ThreadID)
MsgBox Message, vbInformation, Headertext
If MyHook Then UnhookWindowsHookEx MyHook
End Sub

Private Function MsgBoxHookProc(ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Classname As String
Dim i As Long
Dim ii As Long
Dim Text As String
Dim Back As Long
Dim Capheight As Long
Dim PicFeld As String
Dim Stilbit As Long
Dim ExStilBit As Long
Dim Enabled As Boolean

PicFeld = Chr(255) & Chr(4) & Chr(127)

MsgBoxHookProc = CallNextHookEx(MyHook, CodeNo, wParam, lParam)
Select Case CodeNo
    Case HCBT_CREATEWND
        Classname = String$(255, 0)
        GetClassName wParam, Classname, 256
        i = InStr(Classname, vbNullChar)
        If i Then Classname = Left$(Classname, i - 1)
        Select Case UCase(Classname)
        Case "#32770"
            Formhwnd = wParam
        Case "BUTTON"
            Buttonhwnd = wParam
        Case "STATIC"
        WndStatic(Nr) = wParam
        Nr = Nr + 1
End Select

    Case HCBT_ACTIVATE
    MoveWindow Formhwnd, DlgX, DLgY, 526, 135, 1
    Capheight = GetSystemMetrics(33)
    Capheight = Capheight + GetSystemMetrics(4)
    For i = 0 To 1
      Text = Space$(10)
      Back = SendMessage(WndStatic(i), WM_GETTEXT, 10, ByVal Text)
        Text = LCase(Left(Text, Back))
        Select Case Text
        Case LCase("ActiveX-S")
            TextHwnd = WndStatic(i)
        Case LCase(PicFeld), ""
            PicHwnd = WndStatic(i)
        End Select
    Next i
    MoveWindow Buttonhwnd, 443, 42 - Capheight, 64, 28, 1
    MoveWindow PicHwnd, 23, 42 - Capheight, 32, 32, 1
    MoveWindow TextHwnd, 83, 43 - Capheight, 350, 85, 1
    Select Case WindowStil
        Case 0, 3
        Enabled = True
        Case 1
        Enabled = False
    End Select
    Select Case WindowStil
    Case 0, 1
    EnableCloseButton Formhwnd, Enabled
    End Select
    Stilbit = GetWindowLong(Formhwnd, GWL_STYLE)
    Stilbit = Stilbit And Not DS_NOIDLEMSG And Not DS_ABSALIGN
    Select Case WindowStil
    Case 2
    Stilbit = Stilbit And Not WS_SYSMENU
    End Select
    SetWindowLong Formhwnd, GWL_STYLE, Stilbit
    ExStilBit = GetWindowLong(Formhwnd, GWL_EXSTYLE)
    SetWindowLong Formhwnd, GWL_EXSTYLE, ExStilBit
    SendMessage PicHwnd, &H170, Iconhandle, ByVal 0&
    UnhookWindowsHookEx MyHook
    MyHook = 0
End Select
End Function

Public Sub EnableCloseButton(ByVal hWnd As Long, Enabled As Boolean)
Dim hMenu As Long
Dim lpmim As MENUITEMINFO
Dim Anz As Long
Dim Version As Long
Dim Fl As Long

hMenu = GetSystemMenu(hWnd, 0)
Anz = GetMenuItemCount(hMenu)
If Anz < 2 Then
    Version = 0
    Else
    Fl = GetMenuState(hMenu, &HF060&, 0)
    Select Case Fl
        Case 0
        Version = 0
        Case Else
        Version = 1
    End Select
End If
Select Case Version
Case 0 'Win98
    If Enabled = True Then
    With lpmim
        .cbSize = Len(lpmim)
        .fMask = 50
        .wID = &HF060&
    End With
    InsertMenuItem hMenu, 1, 1, lpmim
    SendMessage hWnd, &H86, True, 0
    Else
    'Do Nothing
    End If
Case 1 'XP
    If Enabled = True Then
    'Do Nothing
    Else
    RemoveMenu hMenu, &HF060&, 0
    End If
End Select
End Sub

