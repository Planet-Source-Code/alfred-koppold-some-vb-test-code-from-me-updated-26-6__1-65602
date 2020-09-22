Attribute VB_Name = "NewMsgBox"
Option Explicit
Public bHOOK As Long
Attribute bHOOK.VB_VarMemberFlags = "440"
Public wHOOK As Long
Attribute wHOOK.VB_VarMemberFlags = "440"
Public Testnr As Long
Attribute Testnr.VB_VarMemberFlags = "440"
Public HasFocus As Boolean
Attribute HasFocus.VB_VarMemberFlags = "440"
Public Helpbuttonhwnd As Long
Attribute Helpbuttonhwnd.VB_VarMemberFlags = "440"
Public Subclassed As Boolean
Attribute Subclassed.VB_VarMemberFlags = "440"
Private Formhwnd As Long
Public Buthwnd As Long
Attribute Buthwnd.VB_VarMemberFlags = "440"
Public Buttonhwnd(10) As Long
Attribute Buttonhwnd.VB_VarMemberFlags = "440"
Private PicHwnd As Long
Private TextHwnd As Long
Private Iconhandle As Long
Private DlgX As Long
Private DLgY As Long
Private MyHook As Long
Private Nr As Long
Private WndStatic(1) As Long
Private WindowStil As Long
Public Parenthwnd As Long
Attribute Parenthwnd.VB_VarMemberFlags = "440"
Public Const GWL_WNDPROC = (-4)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Attribute MoveWindow.VB_MemberFlags = "40"
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type

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
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    
Private Const HCBT_CREATEWND = 3
Private Const HCBT_ACTIVATE = 5
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const DS_NOIDLEMSG = &H100
Private Const DS_ABSALIGN = &H1&
Private Const WM_GETTEXT = &HD
Public Const WS_SYSMENU = &H80000
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Attribute GetParent.VB_MemberFlags = "40"
Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShowAboutBox(ByVal parent_hwnd As Long, ByVal HeaderText As String, ByVal Message As String, ByVal szOtherStuff As String, ByVal icon As Long, x As Long, y As Long, Optional Stil As Long = 0)
Attribute ShowAboutBox.VB_MemberFlags = "40"
Iconhandle = icon
DlgX = x
DLgY = y
WindowStil = Stil
Nr = 0
MyHook = SetWindowsHookEx(&H5, AddressOf MsgBoxHookProc, App.hInstance, App.ThreadID)
MsgBox Message, vbInformation, HeaderText
If MyHook Then UnhookWindowsHookEx MyHook
End Sub

Private Function MsgBoxHookProc(ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ClassName As String
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
        ClassName = String$(255, 0)
        GetClassName wParam, ClassName, 256
        i = InStr(ClassName, vbNullChar)
        If i Then ClassName = Left$(ClassName, i - 1)
        Select Case UCase(ClassName)
        Case "#32770"
            Formhwnd = wParam
        Case "BUTTON"
            Buthwnd = wParam
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
    MoveWindow Buthwnd, 443, 42 - Capheight, 64, 28, 1
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

Public Sub EnableCloseButton(ByVal hwnd As Long, Enabled As Boolean)
Attribute EnableCloseButton.VB_MemberFlags = "40"
Dim hMenu As Long
Dim lpmim As MENUITEMINFO
Dim Anz As Long
Dim Version As Long
Dim Fl As Long

hMenu = GetSystemMenu(hwnd, 0)
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
    SendMessage hwnd, &H86, True, 0
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

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Attribute EnumChildProc.VB_MemberFlags = "40"
Dim test As Long
test = GetParent(hwnd)
If test = Parenthwnd Then
Buttonhwnd(Testnr) = hwnd
Testnr = Testnr + 1
If Testnr = 5 Then
Helpbuttonhwnd = hwnd
End If
End If
EnumChildProc = True
End Function

Public Function HButtonProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute HButtonProc.VB_MemberFlags = "40"
Select Case hwnd
Case Parenthwnd
If uMsg = &H111 And lParam = Helpbuttonhwnd And wParam = 9 Then
OpenHelp
Exit Function
End If
    HButtonProc = CallWindowProc(wHOOK, hwnd, uMsg, wParam, lParam)
Case Helpbuttonhwnd
    Select Case uMsg
    Case &H202
    OpenHelp
    Exit Function
    Case &H100
    If wParam = 32 Then
    OpenHelp
    Exit Function
    End If
    Case &H203
    Exit Function
    Case &H8
    HasFocus = False
    Case &H7
    HasFocus = True
    End Select
    HButtonProc = CallWindowProc(bHOOK, hwnd, uMsg, wParam, lParam)
End Select
End Function

Private Sub OpenHelp()
Dim Deskhwnd As Long
Dim Result As Long
Dim rc As RECT
Dim Filename As String
SystemParametersInfo 48, 0, rc, 0
    Filename = Space(255)
    Result = GetWindowsDirectory(Filename, 255)
    Filename = Left(Filename, Result)
    Filename = Filename & "\HELP\Vbcmn98.chm"
    Result = HTMLHelp(0, Filename, 0, "/html/vbPropropertyPagesActiveXControls.htm")
    MoveWindow Result, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 1
End Sub
