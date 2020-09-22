Attribute VB_Name = "modList"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private GoToLB As Long
Private LBAnz As Long
Private SelItem As Long
Private Init As Boolean
Public First As Long
Private LostFocus As Boolean
Private Selected As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  hdc As Long
  rcItem As RECT
  ItemData As Long
End Type

Private Const GWL_WNDPROC = (-4&)
Private Const COLOR_HIGHLIGHT = &HD
Private Const COLOR_BTNFACE = 15
Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_HSCROLL = &H100000
Private Const WS_TABSTOP = &H10000
Private Const LBS_OWNERDRAWFIXED = &H10&
Private Const LBS_NOINTEGRALHEIGHT = &H100&
Private Const LBS_MULTICOLUMN = &H200&
Private Const LBS_DISABLENOSCROLL = &H1000&
Public Const LB_ADDSTRING = &H180
Private Const LB_SETCOLUMNWIDTH = &H195
Private Const ODS_SELECTED = &H1
Private Const WM_DRAWITEM = &H2B
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const LB_SETITEMHEIGHT = &H1A0
Private Const WM_KILLFOCUS = &H8
Private Const WM_SETFOCUS = &H7
Private Const WM_KEYDOWN = &H100
Private Const VK_TAB = &H9
Private Const VK_SHIFT = &H10
Public Const LB_GETCOUNT = &H18B
Public Const LB_SETCURSEL = &H186
Private Const VK_LEFT = &H25
Private Const VK_RIGHT = &H27
Private Const VK_UP = &H26
Private Const VK_DOWN = &H28
Private Const LBS_NOTIFY = &H1&
Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LB_DELETESTRING = &H182
Private Const VK_ALT = &H12

Private Type ODLBTYPE
  Forecolor As Long
  BackColor As Long
  ForeColorSelected As Long
  BackColorSelected As Long
  BorderStyle As Long
  DrawStyle As Long
  ItemHeight As Long
  hwnd As Long
  Left As Long
  Top  As Long
  Width As Long
  Height As Long
  hFont As Long
  Parenthwnd As Long
End Type

Private MyListBox As ODLBTYPE
Attribute MyListBox.VB_VarMemberFlags = "440"
Private PrevWndProc&, PrevWndProcLB&, PrevWndProcAfter&, PrevWndProcBefore&, PrevWndProcButton&
Private TeBo As TextBox
Private hwndB As Long
Private hwndA As Long
Private hwndBut As Long
Private KeyJump As Long

Private Sub SubClass(hwnd&)
  PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
  PrevWndProcAfter = SetWindowLong(hwndA, GWL_WNDPROC, AddressOf WndProc)
  PrevWndProcBefore = SetWindowLong(hwndB, GWL_WNDPROC, AddressOf WndProc)
  PrevWndProcButton = SetWindowLong(hwndBut, GWL_WNDPROC, AddressOf WndProc)
  PrevWndProcLB = SetWindowLong(MyListBox.hwnd, GWL_WNDPROC, AddressOf WndProcLB)
End Sub

Private Sub UnSubClass(hwnd&)
  Call SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
  Call SetWindowLong(hwndA, GWL_WNDPROC, PrevWndProcAfter)
  Call SetWindowLong(hwndB, GWL_WNDPROC, PrevWndProcBefore)
  Call SetWindowLong(hwndBut, GWL_WNDPROC, PrevWndProcButton)
  Call SetWindowLong(MyListBox.hwnd, GWL_WNDPROC, PrevWndProcLB)
End Sub

Private Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Back As Long

Select Case hwnd
Case hwndA
        Select Case Msg
        Case &H100E
        Exit Function
             Case WM_KILLFOCUS
             If GoToLB = 1 Then Exit Function
             Back = GetWindowLong(hwnd, (-16))
            Back = Back And Not 1
            SetWindowLong hwnd, (-16), Back

            Case WM_SETFOCUS
                
                If wParam = hwndB And KeyJump = 0 Then
                    Back = GetKeyState(VK_TAB) And &HF0000000
                        If Back <> 0 Then
                        Back = GetKeyState(VK_SHIFT) And &HF0000000
                            Select Case Back
                            Case 0
                             Back = SendMessage(MyListBox.hwnd, LB_GETCOUNT, 0, 0)
                                If Back > 0 Then
                                WndProc = 0
                                First = 1
                                GoToLB = 1
                                SetFocusA MyListBox.hwnd
                                Back = SendMessage(MyListBox.hwnd, LB_SETCURSEL, SelItem, 0)
                                GoToLB = 0
                                Exit Function
                                End If
                            End Select
                        End If
                End If
        End Select
        WndProc = CallWindowProc(PrevWndProcAfter, hwnd, Msg, wParam, lParam)

Case hwndB
    Select Case Msg
        Case WM_SETFOCUS
            If wParam = hwndA And KeyJump = 0 Then
            Back = GetKeyState(VK_TAB) And &HF0000000
            If Back <> 0 Then
            Back = GetKeyState(VK_SHIFT) And &HF0000000
            Select Case Back
            Case &HF0000000
                Back = SendMessage(MyListBox.hwnd, LB_GETCOUNT, 0, 0)
                 If Back > 0 Then
                 First = 1
                 SetFocusA MyListBox.hwnd
                 Back = SendMessage(MyListBox.hwnd, LB_SETCURSEL, SelItem, 0)
                 Exit Function
                 End If
            End Select
            End If
            End If
    End Select
                
    WndProc = CallWindowProc(PrevWndProcBefore, hwnd, Msg, wParam, lParam)
Case hwndBut
    Select Case Msg
        Case WM_KILLFOCUS
        Back = GetWindowLong(hwnd, (-16))
        Back = Back And Not 1
        SetWindowLong hwnd, (-16), Back
        Case WM_SETFOCUS
        End Select
    WndProc = CallWindowProc(PrevWndProcButton, hwnd, Msg, wParam, lParam)
Case MyListBox.Parenthwnd
  Select Case Msg
    Case WM_CTLCOLORLISTBOX
        WndProc = MyListBox.DrawStyle
    Case WM_DRAWITEM
        DrawItem (lParam)
        WndProc = 1
    Case Else
        WndProc = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
  End Select
End Select
End Function

Private Function WndProcLB(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim Back As Long
  Dim wohin As Long
  
  Select Case Msg
  Case LB_ADDSTRING
  LBAnz = LBAnz + 1
  Case LB_DELETESTRING
  LBAnz = LBAnz - 1
  Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
  If Init = False Then
  Exit Function
  Else
  First = 1
  End If
  Case WM_KILLFOCUS
         Back = GetKeyState(VK_LEFT) And &HF0000000
         If Back = &HF0000000 Then wohin = 1
         Back = GetKeyState(VK_DOWN) And &HF0000000
         If Back = &HF0000000 Then wohin = 1
         Back = GetKeyState(VK_RIGHT) And &HF0000000
         If Back = &HF0000000 Then wohin = 2
         Back = GetKeyState(VK_UP) And &HF0000000
         If Back = &HF0000000 Then wohin = 2
        Select Case wohin
        Case 1
        If SelItem > 0 Then SelItem = SelItem - 1
        SetFocusA MyListBox.hwnd
        Back = SendMessage(MyListBox.hwnd, LB_SETCURSEL, SelItem, 0)
        Exit Function
        Case 2
        If SelItem < LBAnz - 1 Then SelItem = SelItem + 1
        SetFocusA MyListBox.hwnd
        Back = SendMessage(MyListBox.hwnd, LB_SETCURSEL, SelItem, 0)
        Exit Function
        End Select
    LostFocus = True
       Back = GetKeyState(VK_TAB) And &HF0000000
       If Back <> 0 Then
         Back = GetKeyState(VK_SHIFT) And &HF0000000
          Select Case Back
          Case 0
          KeyJump = 2
          wParam = hwndB
          SetFocusA hwndB
          WndProcLB = CallWindowProc(PrevWndProcLB, hwnd, Msg, wParam, lParam)
          SetFocusA hwndA
          KeyJump = 0
          Case Else
          KeyJump = 1
          wParam = hwndA
          WndProcLB = CallWindowProc(PrevWndProcLB, hwnd, Msg, wParam, lParam)
          SetFocusA hwndB
          KeyJump = 0
          End Select
        Exit Function
        End If
  Case WM_SETFOCUS
  LostFocus = False
  End Select
  WndProcLB = CallWindowProc(PrevWndProcLB, hwnd, Msg, wParam, lParam)
End Function

Public Function InitListBox(Parenthwnd As Long, Left As Long, Top As Long, Width As Long, Height As Long, tBox As TextBox, hwndBefore As Long, hwndAfter As Long, hwndButton As Long) As Long
Attribute InitListBox.VB_MemberFlags = "40"
  Dim x&, LStyle&
  KeyJump = 0
  GoToLB = 0
  LBAnz = 0
  Set TeBo = tBox
  hwndBut = hwndButton
  Init = False
  LostFocus = False
  First = 0
  hwndA = hwndAfter
  hwndB = hwndBefore
    With MyListBox
      .BackColor = vbWhite
      .BorderStyle = 5
      .Parenthwnd = Parenthwnd
      .Left = Left
      .Top = Top
      .Width = Width
      .Height = Height
      LStyle = LBS_OWNERDRAWFIXED Or LBS_NOINTEGRALHEIGHT Or LBS_MULTICOLUMN Or LBS_DISABLENOSCROLL Or WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_TABSTOP Or WS_MAXIMIZEBOX
      .hwnd = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_NOPARENTNOTIFY, "LISTBOX", vbNullString, LStyle, .Left, .Top, .Width, .Height, .Parenthwnd, 0, App.hInstance, ByVal 0&)
      .ItemHeight = 54
      .Width = 54
      Call SendMessage(.hwnd, LB_SETITEMHEIGHT, 0&, ByVal .ItemHeight)
                       
          Call SendMessage(.hwnd, LB_SETCOLUMNWIDTH, ByVal 54, 0)

    
      .DrawStyle = CreateSolidBrush(.BackColor)
      
      
      Call SubClass(.Parenthwnd)
    End With
    InitListBox = MyListBox.hwnd
End Function

Public Sub ExitListBox()
Attribute ExitListBox.VB_MemberFlags = "40"
  Call DestroyWindow(MyListBox.hwnd)
  Call DeleteObject(MyListBox.DrawStyle)
  UnSubClass (MyListBox.Parenthwnd)
End Sub

Private Sub DrawItem(lParam&)
Dim Pictureid As Long
Dim BColor&, FColor&, hBrush&, l&
Dim rc As RECT
Dim DI As DRAWITEMSTRUCT

Init = True
    Call CopyMemory(DI, ByVal lParam, Len(DI))
    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
    With DI
      If .itemState And ODS_SELECTED Then
      SelItem = .itemID
      TeBo.Text = SelItem + 1
      FillRect .hdc, .rcItem, hBrush
      DrawFrameControl .hdc, .rcItem, &H4, &H10
      rc.Left = .rcItem.Left + 3
      rc.Right = .rcItem.Right - 3
      rc.Top = .rcItem.Top + 3
      rc.Bottom = .rcItem.Bottom - 3
      If LostFocus = False Then
If First = 1 Then DrawEdge .hdc, rc, 6, 15 'Focus
End If
      Else
       FillRect .hdc, .rcItem, hBrush
      End If
        DrawImagelist Imagelist, ImgArr(.itemID + 1), .hdc, .rcItem.Left + 5, .rcItem.Top + 5, 43, 43
    End With
    
    Call DeleteObject(hBrush)

    Call CopyMemory(ByVal lParam, DI, Len(DI))
End Sub
