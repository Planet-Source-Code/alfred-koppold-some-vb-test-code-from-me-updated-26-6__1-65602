Attribute VB_Name = "NewMsgBox"
Option Explicit

Private Const WM_DESTROY = &H2
Private Const GWL_WNDPROC = (-4)
Private Const WM_PAINT = &HF&
Private Const HCBT_CREATEWND = 3
Private Const HCBT_ACTIVATE = 5
Private Const EDGE_RAISED = 5
Private Const EDGE_SUNKEN = 10
Private Const EDGE_ETCHED = 6
Private Const EDGE_BUMP = 9
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_NOTIFY = &H4E
Private Const WM_COMMAND = &H111
Private Const EN_SETFOCUS = &H100

Private Type POINTAPI
x As Long
y As Long
End Type

Private Type Size
cx As Long
cy As Long
End Type

Private Type SizePriv
Height As Long
Width As Long
Left As Long
Top As Long
End Type

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Type WindowList
ClassName As String
Hwnd As Long
WindowRect As RECT
Text As String
End Type

Private Type PICTDESC
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GUID
    Part1 As Long
    Part2 As Integer
    Part3 As Integer
    Part4 As Integer
    Part5(1 To 6) As Byte
End Type

Dim SesEnd As Long
Private ClickedHwnd As Long
Private SysLVHwnd As Long
Private IsXP As Boolean
Private LVisFlat As Boolean
Private Iscaptured As Boolean
Private MClicked As Boolean
Private B1Hwnd As Long
Private B2Hwnd As Long
Private B1Stil As Long
Private B2Stil As Long
Private EStil As Long
Private PicBoarder As Long
Private StilFlat As Boolean
Private Formbrush As Long
Private MoveDown As Long
Private MoveRight As Long
Private Anz As Long
Private Windows() As WindowList
Private CBinLB32 As Long
Private EnumNum As Long
Private BorderWidthX As Long
Private BorderwidthY As Long
Private WithStretch As Boolean
Private Controlhwnd() As Long
Private FormHwnd As Long
Private DlgX As Long
Private DLgY As Long
Private Nr As Long
Private WHook As Long
Private PBLeftHook As Long
Private PBTopHook As Long
Private PBRightHook As Long
Private PBBottomHook As Long
Private B1Hook As Long
Private B2Hook As Long
Dim LVHook As Long
Private EHook As Long
Private First As Long
Private OldHeight As Long
Private OldWidth As Long
Private NewWidth As Long
Private NewHeight As Long
Private NewPicSize() As SizePriv
Private PBSize() As SizePriv
Private OldX As Long
Private OldY As Long
Private HasPicture() As Boolean
Private PicSize() As SizePriv
Private PicsDC() As Long
Private hPicBox() As Long
Private Statichook As Long
Private Midscreen As Boolean
Private LeftBorder As Long
Private TopBorder As Long
Private CB1Hook As Long
Private CB2Hook As Long
Private CB3Hook As Long
Private CB1Hwnd As Long
Private CB2Hwnd As Long
Private CB3Hwnd As Long
Private CB1Stil As Long
Private CB2Stil As Long
Private CB3Stil As Long
Private CapWidth As Long
Private CapHeight As Long
Private EditHwnd As Long
Private hOldBmp() As Long

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "Olepro32" (ByRef pPictDesc As PICTDESC, ByRef RIID As GUID, ByVal fOwn As Long, ByRef ppvObj As Any) As Long
Private Declare Function IIDFromString Lib "OLE32" (ByVal lpsz As String, ByRef lpiid As GUID) As Long

Public Sub OwnOpen(ByVal parent_hwnd As Long, MyPictures() As StdPicture, Cmdialog As CommonDialog, Optional Stil As Long = 0, Optional Picstyle As Long = 1, Optional x As Long = 0, Optional y As Long = 0)
Dim Screendc As Long
Dim i As Long
Dim Kon As Long
ReDim HasPicture(4)
ReDim Pics(4)
ReDim PicSize(4)
ReDim PicsDC(4)
ReDim hOldBmp(4)
First = 0
SesEnd = 0

PicBoarder = 0
Kon = Picstyle And 8
If Kon = 8 Then PicBoarder = 9
Kon = Picstyle And 4
If Kon = 4 Then PicBoarder = 6
Kon = Picstyle And 2
If Kon = 2 Then PicBoarder = 5
Kon = Picstyle And 1
If Kon = 1 Then PicBoarder = 10

Kon = Stil And 1
Select Case Kon
Case 1
Midscreen = True
Case Else
Midscreen = False
End Select
Kon = Stil And 2
Select Case Kon
Case 2
StilFlat = False
Case Else
StilFlat = True
End Select

'WithStretch = WithStretching

For i = 0 To 3
If IsNothing(MyPictures(i)) = False Then
HasPicture(i) = True
PicSize(i).Height = MyPictures(i).Height * 0.5669 / Screen.TwipsPerPixelY
PicSize(i).Width = MyPictures(i).Width * 0.5669 / Screen.TwipsPerPixelX
Screendc = GetDC(0)
PicsDC(i) = CreateCompatibleDC(Screendc)
ReleaseDC 0, Screendc
hOldBmp(i) = SelectObject(PicsDC(i), MyPictures(i).Handle)
End If
Next i
If IsNothing(MyPictures(4)) = False Then
Formbrush = CreatePatternBrush(MyPictures(4).Handle) 'CreateSolidBrush(vbBlue)
End If
Kon = Stil And 1
Select Case Kon
Case 1
Midscreen = True
Case Else
Midscreen = False
End Select
Kon = Stil And 2
Select Case Kon
Case 2
StilFlat = False
Case Else
StilFlat = True
End Select

DlgX = x
DLgY = y
Nr = 0
ReDim Controlhwnd(20)
WHook = SetWindowsHookEx(&H5, AddressOf MsgBoxHookProc, App.hInstance, App.ThreadID)
Cmdialog.ShowOpen
ClearUp
End Sub

Private Function MsgBoxHookProc(ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ClassName As String
Dim Newhwnd As Long
Dim Temp As Long
Dim i As Long
Dim Back As Long
Dim Style As Long
Dim rc As RECT
Dim h As Long
Dim w As Long
Dim Bild1 As Long
Dim Bild2 As Long
Dim Bild3 As Long
Dim Bild4 As Long

ReDim hPicBox(3)
MsgBoxHookProc = CallNextHookEx(WHook, CodeNo, wParam, lParam)
Select Case CodeNo
    Case HCBT_CREATEWND
    ClassName = String$(255, 0)
    GetClassName wParam, ClassName, 256
    i = InStr(ClassName, vbNullChar)
    If i Then ClassName = Left$(ClassName, i - 1)
    Select Case UCase(ClassName)
        Case "#32770"
        FormHwnd = wParam
        Style = GetWindowLong(FormHwnd, (-16))
        Style = Style And Not &H40000
        SetWindowLong FormHwnd, (-16), Style
        End Select
    Case HCBT_ACTIVATE
    If wParam = FormHwnd Then
    GetClientRect FormHwnd, rc
    w = rc.Right - rc.Left
    h = rc.Bottom - rc.Top
    If HasPicture(0) Then Bild1 = CreateWindowEx(0, "Static", "", &H50000000, 0, 0, 10, 10, FormHwnd, 0, 0, ByVal 0)
    If HasPicture(1) Then Bild2 = CreateWindowEx(0, "Static", "", &H50000000, 0, 0, 10, 10, FormHwnd, 0, 0, ByVal 0)
    If HasPicture(2) Then Bild3 = CreateWindowEx(0, "Static", "", &H50000000, 0, 0, 10, 10, FormHwnd, 0, 0, ByVal 0)
    If HasPicture(3) Then Bild4 = CreateWindowEx(0, "Static", "", &H50000000, 0, 0, 10, 10, FormHwnd, 0, 0, ByVal 0)
    hPicBox(0) = Bild1
    hPicBox(1) = Bild2
    hPicBox(2) = Bild3
    hPicBox(3) = Bild4
    UnhookWindowsHookEx WHook
    WHook = 0
    Berechnungen
    WHook = SetWindowLong(FormHwnd, GWL_WNDPROC, AddressOf FrmWndProc)
    If HasPicture(0) Then PBLeftHook = SetWindowLong(hPicBox(0), GWL_WNDPROC, AddressOf FrmWndProc)
    If HasPicture(1) Then PBRightHook = SetWindowLong(hPicBox(1), GWL_WNDPROC, AddressOf FrmWndProc)
    If HasPicture(2) Then PBTopHook = SetWindowLong(hPicBox(2), GWL_WNDPROC, AddressOf FrmWndProc)
    If HasPicture(3) Then PBBottomHook = SetWindowLong(hPicBox(3), GWL_WNDPROC, AddressOf FrmWndProc)
    End If
End Select
End Function

Private Sub PaintBMP(Number As Long)
Dim Back As Long
Dim dc As Long
Dim rc As RECT

dc = GetDC(hPicBox(Number))
rc.Bottom = PBSize(Number).Height
rc.Right = PBSize(Number).Width
Back = StretchBlt(dc, BorderWidthX, BorderwidthY, NewPicSize(Number).Width, NewPicSize(Number).Height, PicsDC(Number), 0, 0, PicSize(Number).Width, PicSize(Number).Height, vbSrcCopy)
If PicBoarder <> 0 Then DrawEdge dc, rc, PicBoarder, 15
UpdateWindow hPicBox(Number)
ReleaseDC hPicBox(Number), dc
End Sub


Private Function GetAddressOf(ByVal Address As Long)
GetAddressOf = Address
End Function

Private Function FrmWndProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim rc As RECT
Dim Menge As Long
Dim x As Integer
Dim Stil As Long
Dim Test As Long
Dim OldHwnd As Long
Dim y As Integer
Dim Hi As Integer
Select Case Hwnd
    Case FormHwnd
    Select Case uMsg
    Case &H111
    If wParam = 2 Then
    SesEnd = 1
    End If
        Case &H138, &H136
            Call SetTextColor(wParam, vbBlack)
            SetBkMode wParam, 1
            If Formbrush <> 0 Then FrmWndProc = Formbrush
            Exit Function
        Case WM_PAINT
            If First = 0 Then
            Firstpainting
            First = 1
            End If
            
        Case &H52E
        If LVisFlat = False Then
        If SysLVHwnd <> 0 Then
        OldHwnd = SysLVHwnd
        EnumChildWindows FormHwnd, AddressOf EnumChildProc, 12
        If OldHwnd <> SysLVHwnd Then
        SetWindowLong OldHwnd, GWL_WNDPROC, LVHook
        SetHook
        End If
        End If
        End If
        Case WM_DESTROY
        MoveWindow FormHwnd, OldX, OldY, OldWidth, OldHeight, 1
        End Select
FrmWndProc = CallWindowProc(WHook, Hwnd, uMsg, wParam, lParam)
Case hPicBox(0)
    FrmWndProc = CallWindowProc(PBLeftHook, Hwnd, uMsg, wParam, lParam)
    If uMsg = WM_PAINT Or uMsg = &HA Then PaintBMP 0
Case hPicBox(1)
    FrmWndProc = CallWindowProc(PBRightHook, Hwnd, uMsg, wParam, lParam)
    If uMsg = WM_PAINT Or uMsg = &HA Then PaintBMP 1
Case hPicBox(2)
    FrmWndProc = CallWindowProc(PBTopHook, Hwnd, uMsg, wParam, lParam)
    If uMsg = WM_PAINT Or uMsg = &HA Then PaintBMP 2
Case hPicBox(3)
    FrmWndProc = CallWindowProc(PBBottomHook, Hwnd, uMsg, wParam, lParam)
    If uMsg = WM_PAINT Or uMsg = &HA Then PaintBMP 3
Case B1Hwnd
    If StilFlat And uMsg = WM_PAINT Then SendMessage B1Hwnd, &HF4, 0, 0
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(B1Hook, Hwnd, uMsg, wParam, lParam)
    If StilFlat Then MsgControl uMsg, Hwnd, B1Stil, lParam, wParam, 0
    If uMsg = WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(B1Hook, Hwnd, uMsg, wParam, lParam)
Case B2Hwnd
    If StilFlat And uMsg = WM_PAINT Then SendMessage B2Hwnd, &HF4, 0, 0
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(B2Hook, Hwnd, uMsg, wParam, lParam)
    If StilFlat Then MsgControl uMsg, Hwnd, B2Stil, lParam, wParam, 0
    If uMsg = WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(B2Hook, Hwnd, uMsg, wParam, lParam)
Case CB1Hwnd
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(CB1Hook, Hwnd, uMsg, wParam, lParam)
    If StilFlat Then Menge = MsgControl(uMsg, Hwnd, CB1Stil, lParam, wParam, 9)
    If uMsg = WM_LBUTTONDOWN And Menge > 0 Then FrmWndProc = CallWindowProc(CB1Hook, Hwnd, uMsg, wParam, lParam)
Case CB2Hwnd
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(CB2Hook, Hwnd, uMsg, wParam, lParam)
    If StilFlat Then Menge = MsgControl(uMsg, Hwnd, CB2Stil, lParam, wParam, 9)
    If uMsg = WM_LBUTTONDOWN And Menge > 0 Then FrmWndProc = CallWindowProc(CB2Hook, Hwnd, uMsg, wParam, lParam)
Case CB3Hwnd
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(CB3Hook, Hwnd, uMsg, wParam, lParam)
    If StilFlat Then Menge = MsgControl(uMsg, Hwnd, CB3Stil, lParam, wParam, 9)
    If uMsg = WM_LBUTTONDOWN And Menge > 0 Then FrmWndProc = CallWindowProc(CB3Hook, Hwnd, uMsg, wParam, lParam)
Case EditHwnd
    If uMsg <> WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(EHook, Hwnd, uMsg, wParam, lParam)
    If StilFlat And IsXP = False Then MsgControl uMsg, Hwnd, EStil, lParam, wParam, 1
    If uMsg = WM_LBUTTONDOWN Then FrmWndProc = CallWindowProc(EHook, Hwnd, uMsg, wParam, lParam)
Case SysLVHwnd
    Select Case uMsg
    Case WM_PAINT
    If LVisFlat = False Then
    Debug.Print "Jetzt flach zeichnen"
    Debug.Print Hex(Hwnd)
    SetLVStyle Hwnd, 1
    LVisFlat = True
    End If
    Case WM_DESTROY
    If LVHook <> 0 And SesEnd = 0 Then
    SetWindowLong SysLVHwnd, GWL_WNDPROC, LVHook
    SysLVHwnd = -1
    LVisFlat = False
    End If
    End Select
    FrmWndProc = CallWindowProc(LVHook, Hwnd, uMsg, wParam, lParam)
End Select

End Function

Private Sub SetHook()
        LVHook = SetWindowLong(SysLVHwnd, GWL_WNDPROC, AddressOf FrmWndProc)
End Sub

Private Function MsgControl(uMsg As Long, Hwnd As Long, Stil As Long, lParam As Long, wParam As Long, Typ As Long) As Long
Dim rc As RECT
Dim x As Integer
Dim y As Integer
Dim Menge As Long
Dim DoIt As Boolean

Select Case uMsg
Case WM_MOUSEMOVE
    If StilFlat And MClicked = False Then
    If Iscaptured = False Then
    If Stil = 0 Then Stil = 3
    GetWindowRect Hwnd, rc
    CapWidth = rc.Right - rc.Left
    CapHeight = rc.Bottom - rc.Top
    Select Case Typ
    Case 9
    DrawCombo Hwnd, Stil
    Case Else
    DrawCRect Hwnd, Stil, Typ
    End Select
    SetCapture Hwnd
    Iscaptured = True
    End If
    End If
    CopyMemory ByVal VarPtr(x), ByVal VarPtr(lParam), 2
    CopyMemory ByVal VarPtr(y), ByVal VarPtr(lParam) + 2, 2
    If x < 0 Or x > CapWidth Or y < 0 Or y > CapHeight Then
    If Iscaptured And MClicked = False Then
    ReleaseCapture
    Iscaptured = False
    Stil = 0
    End If
    Select Case Typ
    Case 9
    DrawCombo Hwnd, Stil
    Case Else
    DrawCRect Hwnd, Stil, Typ
    End Select
    End If
Case WM_LBUTTONDOWN
    If Typ = 9 Then
    Menge = SendMessage(Hwnd, &H146, 0, 0)
    End If
    If Typ = 9 And Menge = 0 Then
    DoIt = False
    Else
    DoIt = True
    End If
    If IsXP = False Then DoIt = True
    If DoIt Then
    MsgControl = 1
    If Iscaptured Then
    MClicked = True
    ClickedHwnd = Hwnd
    Stil = 0
    ReleaseCapture
    Iscaptured = False
    End If
    End If
Case &H202 'WM_LBUTTONUP
If Iscaptured = False And MClicked And Hwnd = ClickedHwnd Then
    MClicked = False
    Hwnd = 0
    Stil = 0
    MClicked = False
    End If
Case WM_PAINT
    Select Case Typ
    Case 9
    DrawCombo Hwnd, Stil
    Case Else
    DrawCRect Hwnd, Stil, Typ
    End Select
End Select
End Function
Private Sub Firstpainting()
Dim ButSize As SizePriv
Dim i As Long
ReDim Windows(0)
Dim Nr1 As Long
Dim Nr2 As Long
Dim NrScroll As Long
Dim NrShell As Long
Dim NrComboEx As Long
Dim NrList As Long
Dim NrEdit As Long
Dim MovXP As Long
Dim More As Long
Dim Style As Long
Dim sizeT As Size
Dim Back As Long
Dim dc As Long
Dim Text As String
Dim TWidth As Long
Dim anzahl As Long
Dim WinRect As RECT
Dim WidthZw As Long
Dim NrButton() As Long
Dim AnzButton As Long
Dim NrCombo() As Long
Dim AnzCombo As Long
Dim NrStatic() As Long
Dim AnzStatic As Long
Dim NrPicBox As Long
Dim NrToolB() As Long
Dim AnzToolB As Long
Dim z As Long
Dim Number As Long
Dim Screendc As Long
Dim dc1 As Long
Dim dc2 As Long
Dim Test As Long
IsXP = False
Iscaptured = False
MClicked = False
Anz = 0
EnumNum = 0
EnumChildWindows FormHwnd, AddressOf EnumChildProc, 12
For i = 1 To Anz
Select Case Windows(i).ClassName
Case "SYSLISTVIEW32"
SysLVHwnd = Windows(i).Hwnd
Case "SCROLLBAR"
NrScroll = i
Case "COMBOBOX"
AnzCombo = AnzCombo + 1
ReDim Preserve NrCombo(AnzCombo)
NrCombo(AnzCombo) = i
Case "EDIT"
NrEdit = i
Case "TOOLBARWINDOW32"
AnzToolB = AnzToolB + 1
ReDim Preserve NrToolB(AnzToolB)
NrToolB(AnzToolB) = i
Case "BUTTON"
AnzButton = AnzButton + 1
ReDim Preserve NrButton(AnzButton)
NrButton(AnzButton) = i
Case "STATIC"
Number = 0
For z = 0 To 3
If Windows(i).Hwnd = hPicBox(z) Then
Number = 1
End If
Next z
If Number = 0 Then
AnzStatic = AnzStatic + 1
ReDim Preserve NrStatic(AnzStatic)
NrStatic(AnzStatic) = i
End If
Case "COMBOBOXEX32"
NrComboEx = i
IsXP = True
Case "LISTBOX"
NrList = i
Case "SHELLDLL_DEFVIEW"
NrShell = i
End Select
Next i

'Scrollbar
DestroyWindow Windows(NrScroll).Hwnd
'Syslistview
If StilFlat = True Then
LVHook = SetWindowLong(SysLVHwnd, GWL_WNDPROC, AddressOf FrmWndProc)
End If
'Buttons
DestroyWindow Windows(NrButton(1)).Hwnd
B1Hwnd = Windows(NrButton(2)).Hwnd
B2Hwnd = Windows(NrButton(3)).Hwnd
MoveControls B1Hwnd, 447 - LeftBorder, 247 - TopBorder, 100, 28
MoveControls B2Hwnd, 447 - LeftBorder, 283 - TopBorder, 100, 28
DestroyWindow Windows(NrButton(4)).Hwnd
If StilFlat Then
B1Hook = SetWindowLong(B1Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
B2Hook = SetWindowLong(B2Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
SetWindowLong B1Hwnd, (-16), &H5401A000
SetWindowLong B2Hwnd, (-16), &H5401A000
End If
'Edit
EditHwnd = Windows(NrEdit).Hwnd
If IsXP = False Then
MoveControls Windows(NrEdit).Hwnd, 111 - LeftBorder, 249 - TopBorder, 310, 24
If StilFlat Then EHook = SetWindowLong(EditHwnd, GWL_WNDPROC, AddressOf FrmWndProc)
End If
'Static
If IsXP = True Then
SetWindowLong Windows(NrStatic(1)).Hwnd, (-20), 4
SetWindowLong Windows(NrStatic(1)).Hwnd, (-16), &H50020100
End If
dc = GetDC(Windows(NrStatic(1)).Hwnd)
Text = Replace(Windows(NrStatic(1)).Text, "&", "")
Back = GetTextExtentPoint32(dc, Text, Len(Text), sizeT)
ReleaseDC Windows(NrStatic(1)).Hwnd, dc
MoveControls Windows(NrStatic(1)).Hwnd, 11 - LeftBorder, 39 - TopBorder, sizeT.cx, 16
MoveControls Windows(NrStatic(2)).Hwnd, 11 + sizeT.cx + 264 - LeftBorder, 31 - TopBorder, 204, 34
MoveControls Windows(NrStatic(3)).Hwnd, 13 - LeftBorder, 251 - TopBorder, Windows(NrStatic(3)).WindowRect.Right - Windows(NrStatic(3)).WindowRect.Left, 16
MoveControls Windows(NrStatic(4)).Hwnd, 13 - LeftBorder, 287 - TopBorder, Windows(NrStatic(4)).WindowRect.Right - Windows(NrStatic(4)).WindowRect.Left, 16
'Combo
CB1Stil = 0
CB2Stil = 0
CB3Stil = 0
CB1Hwnd = Windows(NrCombo(1)).Hwnd
CB2Hwnd = Windows(NrCombo(2)).Hwnd
MoveControls CB1Hwnd, 11 + sizeT.cx - LeftBorder, 33 - TopBorder, 264, 25
If StilFlat Then CB1Hook = SetWindowLong(CB1Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
Select Case IsXP
Case False
MoveControls CB2Hwnd, 111 - LeftBorder, 285 - TopBorder, Windows(NrCombo(2)).WindowRect.Right - Windows(NrCombo(2)).WindowRect.Left, 24
If StilFlat Then CB2Hook = SetWindowLong(CB2Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
Case True
CB2Hook = SetWindowLong(CB2Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
CB3Hwnd = Windows(NrCombo(3)).Hwnd
MoveControls CB3Hwnd, 111 - LeftBorder, 285 - TopBorder, 310, 24
If StilFlat Then CB3Hook = SetWindowLong(CB3Hwnd, GWL_WNDPROC, AddressOf FrmWndProc)
End Select
'Toolbox
If AnzToolB = 2 Then DestroyWindow Windows(NrToolB(2)).Hwnd 'XP
MakeTBStyle Windows(NrToolB(1)).Hwnd, StilFlat
MoveControls Windows(NrToolB(1)).Hwnd, 11 + sizeT.cx + 264 - LeftBorder, 31 - TopBorder, 204, 34
ReleaseDC Windows(NrToolB(1)).Hwnd, dc2
'syslistview
MoveControls Windows(NrShell).Hwnd, 11 - LeftBorder, 67 - TopBorder, 544, 170
MoveControls Windows(NrList).Hwnd, 11 - LeftBorder, 67 - TopBorder, 544, 170
'ComboEx
If IsXP Then MoveControls Windows(NrComboEx).Hwnd, 111 - LeftBorder, 249 - TopBorder, 310, 24
'Picbox
For i = 0 To 3
If HasPicture(i) Then
MoveWindow hPicBox(i), PBSize(i).Left, PBSize(i).Top, PBSize(i).Width, PBSize(i).Height, 0
End If
Next i
MoveWindow FormHwnd, DlgX, DLgY, NewWidth, NewHeight, 1

End Sub

Private Sub Berechnungen()
Dim Windowrc As RECT
Dim Clientrc As RECT
Dim Bottomborder As Long
Dim i As Long
Dim newTop As Long

ReDim PBSize(3)
ReDim NewPicSize(3)
NewWidth = 0
NewHeight = 0
MoveRight = 0
MoveDown = 0
NewWidth = 582
NewHeight = 326
Select Case PicBoarder
Case 0
BorderwidthY = 0
BorderWidthX = 0
Case Else
BorderwidthY = 2
BorderWidthX = 2
End Select
LeftBorder = GetSystemMetrics(7)
TopBorder = GetSystemMetrics(8) + GetSystemMetrics(4)
Bottomborder = GetSystemMetrics(8)
GetWindowRect FormHwnd, Windowrc
GetClientRect FormHwnd, Clientrc
OldX = Windowrc.Left
OldY = Windowrc.Top
OldWidth = Windowrc.Right - Windowrc.Left
OldHeight = Windowrc.Bottom - Windowrc.Top
For i = 0 To 3
PBSize(i).Height = PicSize(i).Height + (2 * BorderwidthY)
PBSize(i).Width = PicSize(i).Width + (2 * BorderWidthX)
Next i

NewWidth = 582
NewHeight = 326
newTop = 0
If HasPicture(2) Then
    PBSize(2).Top = 2
    PBSize(2).Left = 2
    newTop = PBSize(2).Height + 2
    MoveDown = PBSize(2).Height + 2
    NewPicSize(2).Height = PicSize(2).Height
    NewPicSize(2).Width = PicSize(2).Width
    NewHeight = NewHeight + PBSize(2).Height + 2
End If

If HasPicture(0) Then
    NewWidth = NewWidth + PBSize(0).Width
    MoveRight = PBSize(0).Width + 2
    PBSize(0).Left = 2
    PBSize(0).Top = newTop
    NewPicSize(0).Height = PicSize(0).Height
    NewPicSize(0).Width = PicSize(0).Width
End If
If HasPicture(1) Then
    PBSize(1).Left = NewWidth - LeftBorder - LeftBorder
    PBSize(1).Top = newTop
    NewWidth = NewWidth + PBSize(1).Width
    NewPicSize(1).Height = PicSize(1).Height
    NewPicSize(1).Width = PicSize(1).Width
End If
If HasPicture(3) Then
    PBSize(3).Left = 2
    NewHeight = NewHeight + PBSize(3).Height + 2
    PBSize(3).Top = NewHeight - TopBorder - Bottomborder - PBSize(3).Height - 2
    NewPicSize(3).Height = PicSize(3).Height
    NewPicSize(3).Width = PicSize(3).Width
End If

If Midscreen = True Then
DlgX = ((Screen.Width / Screen.TwipsPerPixelX) - (NewWidth)) / 2
DLgY = ((Screen.Height / Screen.TwipsPerPixelY) - NewHeight) / 2
End If
End Sub
Private Function EnumChildProc(ByVal Hwnd As Long, ByVal lParam As Long) As Boolean
Dim ClassName As String * 256
Dim CurCls As String
Dim Back As Long
Dim Text As String

If (GetClassName(Hwnd, ClassName, 256) = 0) Then
EnumChildProc = True
Exit Function
End If
ClassName = Trim$(ClassName)
If InStr(ClassName, Chr$(0)) > 0 Then CurCls = Left$(ClassName, InStr(ClassName, Chr$(0)) - 1)
Text = Space(50)
Back = GetWindowText(Hwnd, Text, 50)
Text = Left(Text, Back)
Select Case EnumNum
Case 0
Anz = Anz + 1
ReDim Preserve Windows(Anz)
Windows(Anz).ClassName = UCase(CurCls)
If UCase(CurCls) = "SYSLISTVIEW32" Then SysLVHwnd = Hwnd
Windows(Anz).Hwnd = Hwnd
Windows(Anz).Text = Text
GetWindowRect Hwnd, Windows(Anz).WindowRect
ToClient Windows(Anz).WindowRect
Case 1
If UCase(CurCls) = "COMBOBOX" Then
CBinLB32 = Hwnd
End If
End Select
EnumChildProc = True
End Function

Private Sub ToClient(WRect As RECT)
Dim h As Long
Dim w As Long
Dim Pt As POINTAPI
w = WRect.Right - WRect.Left
h = WRect.Bottom - WRect.Top
Pt.x = WRect.Left
Pt.y = WRect.Top
ScreenToClient FormHwnd, Pt
WRect.Left = Pt.x
WRect.Top = Pt.y
WRect.Right = Pt.x + w
WRect.Bottom = Pt.y + h
End Sub

Private Sub MoveControls(Hwnd As Long, x As Long, y As Long, Width As Long, Height As Long)
x = x + MoveRight
y = y + MoveDown
MoveWindow Hwnd, x, y, Width, Height, 0
End Sub

Private Function IsNothing(Testobj As Object) As Boolean
Dim a As Long
On Error GoTo No
a = Testobj.Handle
IsNothing = False
Exit Function
No:
IsNothing = True
End Function

Public Sub MakeTBStyle(TBHwnd As Long, Tbstyle As Boolean)
Dim Style As Long
Select Case Tbstyle
Case True
Style = 1342409029
Case False
Style = 1342406981
End Select
SetWindowLong TBHwnd, (-16), Style
End Sub

Private Sub DrawCombo(Hwnd As Long, ByVal dwStyle As Long)
 Dim rct As RECT
 Dim cmbDC As Long
 GetClientRect Hwnd, rct
 cmbDC = GetDC(Hwnd)
 Select Case dwStyle
    Case 3
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight
    Case 0
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
    Case 1, 2
        DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
        InflateRect rct, -1, -1
        DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 End Select
 InflateRect rct, -1, -1
 rct.Left = rct.Right - GetSystemMetrics(10)
 DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 InflateRect rct, -1, -1
 DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
 Select Case dwStyle
    Case 0
        rct.Top = rct.Top - 1
        rct.Bottom = rct.Bottom + 1
        DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight
        rct.Left = rct.Left - 1
        rct.Right = rct.Left
        DrawRect cmbDC, rct, vbWindowBackground, &H0
    Case 1
        rct.Top = rct.Top - 1
        rct.Bottom = rct.Bottom + 1
        rct.Right = rct.Right + 1
        DrawRect cmbDC, rct, vb3DHighlight, vbButtonShadow
    Case 2
        rct.Left = rct.Left - 1
        rct.Top = rct.Top - 2
        OffsetRect rct, 1, 1
        DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
 End Select
 DeleteDC cmbDC
End Sub


Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Function DrawRect(ByVal hdc As Long, ByRef rct As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)
 Dim hPen As Long
 Dim hPenOld As Long
 Dim tP As POINTAPI
 hPen = CreatePen(0, 1, TranslateColor(oTopLeftColor))
 hPenOld = SelectObject(hdc, hPen)
 MoveToEx hdc, rct.Left, rct.Bottom - 1, tP
 LineTo hdc, rct.Left, rct.Top
 LineTo hdc, rct.Right - 1, rct.Top
 SelectObject hdc, hPenOld
 DeleteObject hPen
 If (rct.Left <> rct.Right) Then
    hPen = CreatePen(0, 1, TranslateColor(oBottomRightColor))
    hPenOld = SelectObject(hdc, hPen)
    LineTo hdc, rct.Right - 1, rct.Bottom - 1
    LineTo hdc, rct.Left, rct.Bottom - 1
    SelectObject hdc, hPenOld
    DeleteObject hPen
 End If
End Function

Private Sub DrawCRect(Hwnd As Long, Stil As Long, Typ As Long)
Dim i As Long
Dim Color As Long
Dim rc As RECT
Dim dc As Long

dc = GetDC(Hwnd)
Select Case Typ
Case 0
    Select Case Stil
    Case 0
    Color = vbButtonFace
    Case Else
    Color = vb3DHighlight
    End Select
    GetClientRect Hwnd, rc
    DrawRect dc, rc, Color, Color
Case 1
    GetClientRect Hwnd, rc
    For i = 1 To 2
    rc.Left = rc.Left - 1
    rc.Top = rc.Top - 1
    rc.Right = rc.Right + 1
    rc.Bottom = rc.Bottom + 1
    Select Case Stil
    Case 0
    DrawRect dc, rc, vbButtonFace, vbButtonFace
    Case 3
    If i = 1 Then
    DrawRect dc, rc, vb3DHighlight, vb3DHighlight
    End If
    End Select
    Next i
End Select
DeleteDC dc
End Sub

Private Sub ClearUp()
Dim i As Long
SetWindowLong FormHwnd, GWL_WNDPROC, WHook
WHook = 0
If HasPicture(0) Then SetWindowLong hPicBox(0), GWL_WNDPROC, PBLeftHook
PBLeftHook = 0
If HasPicture(1) Then SetWindowLong hPicBox(1), GWL_WNDPROC, PBRightHook
PBRightHook = 0
If HasPicture(2) Then SetWindowLong hPicBox(2), GWL_WNDPROC, PBTopHook
PBTopHook = 0
If HasPicture(3) Then SetWindowLong hPicBox(3), GWL_WNDPROC, PBBottomHook
PBBottomHook = 0
SetWindowLong B1Hwnd, GWL_WNDPROC, B1Hook
B1Hook = 0
SetWindowLong B2Hwnd, GWL_WNDPROC, B2Hook
B2Hook = 0
If LVHook <> 0 And SysLVHwnd <> -1 Then SetWindowLong SysLVHwnd, GWL_WNDPROC, LVHook
LVHook = 0
If IsXP Then SetWindowLong EditHwnd, GWL_WNDPROC, EHook
EHook = 0
SetWindowLong CB1Hwnd, GWL_WNDPROC, CB1Hook
CB1Hook = 0
If CB2Hook <> 0 Then SetWindowLong CB2Hwnd, GWL_WNDPROC, CB2Hook
CB2Hook = 0
If CB3Hook <> 0 Then SetWindowLong CB3Hwnd, GWL_WNDPROC, CB3Hook
CB3Hook = 0
For i = 0 To 4
If HasPicture(i) Then
DeleteDC PicsDC(i)
PicsDC(i) = 0
DeleteObject hOldBmp(i)
hOldBmp(i) = 0
End If
Next i
ReDim HasPicture(4)
If Formbrush <> 0 Then DeleteObject Formbrush
Formbrush = 0
End Sub

Private Function SetLVStyle(Hwnd As Long, IsFlat As Long) As Long
Dim lStyle As Long
   lStyle = SendMessageByLong(Hwnd, &H1037, 0, 0)
   Select Case IsFlat
    Case 0
      lStyle = lStyle And Not &H100
   Case 1
      lStyle = lStyle Or &H100
   End Select
   
   SendMessageByLong Hwnd, &H1036, 0, lStyle
End Function

