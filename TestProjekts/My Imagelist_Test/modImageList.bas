Attribute VB_Name = "modImageList"
Option Explicit
Public KeyArr() As String
Public TagArr() As String
Public ImgArr() As Long
Public Schange As Boolean
Public Ini As Boolean
Public PP As New PropertyBag
Public MaskColorIntern As Long
Public BackColorIntern As Long
Public cx As Long
Public cy As Long
Public Anzahl As Long
Public Const EMsg As String = "Die Eigenschaft ist schreibgeschützt, wenn die Abbildungsliste Abbildungen enthält."
Public Imagelist As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type IMAGEINFO
    hbmImage As Long
    hbmMask As Long
    Unused1 As Long
    Unused2 As Long
    rcImage As RECT
End Type
Private Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIml As Long, ByRef cx As Long, ByRef cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long) As Long
Attribute ImageList_Remove.VB_MemberFlags = "40"
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetImageInfo Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, pimageinfo As IMAGEINFO) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal iImageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fFlags As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const ILC_COLOR4 = &H5 'all with Mask(&H1)
Private Const ILC_COLOR8 = &H9
Private Const ILC_COLOR16 = &H11
Private Const ILC_COLOR24 = &H19
Private Const ILC_COLOR32 = &H21
Private Const LR_COPYRETURNORG = &H4
Private Const COLOR_BTNFACE = 15

Private hbmp As Long
Private BrushBG As Long
Private TempDC As Long
Private hbmpmono As Long
Private hbmpOld As Long

Public Function CreateImageList(Width As Long, Height As Long) As Long
Attribute CreateImageList.VB_MemberFlags = "40"
Dim BGColor As Long
Dim hIml As Long
BGColor = GetSysColor(COLOR_BTNFACE)
hIml = ImageList_Create(Width, Height, ILC_COLOR32, 10, 10)
CreateTempHdc Width, Height
BrushBG = CreateSolidBrush(BGColor)
CreateImageList = hIml
End Function

Public Function DestroyImageList(hIml As Long) As Long
DestroyImageList = ImageList_Destroy(hIml)
DeleteTempDC
DeleteObject BrushBG
End Function

Public Function ImageListAdd(hIml As Long, ImagePath As String, Width As Long, Height As Long, Optional Pic As Long = 0, Optional ES As Long = 0) As Long
Dim hbmp As Long
Dim Typ As Long
Dim clrMask As Long
Dim Back As Long
Dim IWidth As Long
Dim IHeight As Long
Dim oHeight As Long
Dim oWidth As Long
Dim NoW As Long
Dim NoH As Long

If Width = 0 Then NoW = 1
If Height = 0 Then NoH = 1
If ES Then
IHeight = 43
IWidth = 43
Else
IHeight = Height
IWidth = Width
End If
On Error GoTo ErrorMsg
clrMask = RGB(192, 192, 192)
Select Case ImagePath
Case ""
hbmp = Pic
Case Else
    hbmp = LoadImageFile(ImagePath, IWidth, IHeight, Typ, oWidth, oHeight)
    If NoW Then cx = oWidth
    If NoH Then cy = oHeight
End Select

If Imagelist = 0 Then
hIml = CreateImageList(IWidth, IHeight) ' erstellen
Imagelist = hIml
End If
If hbmp <> 0 Then
Select Case Typ
Case 0 'Bitmap
Back = ImageList_AddMasked(hIml, hbmp, clrMask)
Case 1, 2 'Icon Cursor
Back = ImageList_AddIcon(hIml, hbmp)
Case Else
Back = -1
End Select
If Back >= 0 Then
Call DeleteObject(hbmp)
ImageListAdd = Back
End If
End If
Exit Function
ErrorMsg:
MsgBox Err.Description, vbExclamation, "ImgListCtrl"
Err.Clear
Back = -1
End Function

Public Function ImagelistRemove(hIml As Long, IndexNr As Long)
Attribute ImagelistRemove.VB_MemberFlags = "40"
Dim Back As Long
Back = ImageList_Remove(hIml, IndexNr - 1)
End Function

Private Function LoadImageFile(ImagePath As String, Width As Long, Height As Long, Typ As Long, oWidth As Long, oHeight As Long) As Long
  Dim Back As Long
  Dim Pic As StdPicture
  Dim ScreenDc As Long
  Dim FirstDc As Long
  Dim secdc As Long
  Dim Old As Long
  Dim hbm As Long
  
  Set Pic = LoadPicture(ImagePath)
oHeight = Pic.Height * 0.5669 / Screen.TwipsPerPixelX
oWidth = Pic.Width * 0.5669 / Screen.TwipsPerPixelY
If Width = 0 Or Height = 0 Then
Width = oWidth
Height = oHeight
End If
ScreenDc = GetDC(0)
FirstDc = CreateCompatibleDC(ScreenDc)
Old = SelectObject(FirstDc, Pic.handle)
If Old = 0 Then
Typ = GetimageTyp(ImagePath, oWidth, oHeight)
If Width = 0 Or Height = 0 Then
Width = oWidth
Height = oHeight
End If
Select Case Typ
Case 1
LoadImageFile = LoadImage(0, ImagePath, 1, Width, Height, 16) 'Icon
Exit Function
Case 2
LoadImageFile = LoadImage(0, ImagePath, 2, Width, Height, 16) 'Cursor
Exit Function
Case -1
Exit Function 'Error
End Select
End If
StretchBlt FirstDc, 0, 0, oWidth, oHeight, FirstDc, 0, 0, oWidth, oHeight, vbSrcCopy
secdc = CreateCompatibleDC(ScreenDc)
hbmp = CreateCompatibleBitmap(ScreenDc, Width, Height)
Old = SelectObject(secdc, hbmp)
DrawBG secdc, Width, Height
Back = StretchBlt(secdc, 0, 0, Width, Height, FirstDc, 0, 0, oWidth, oHeight, vbSrcCopy)
SelectObject secdc, Old
LoadImageFile = hbmp
DeleteDC ScreenDc
DeleteObject Old
DeleteDC FirstDc
DeleteDC secdc
End Function

Public Function DrawImagelist(hIml As Long, PicNr As Long, PicHdc As Long, Left As Long, Top As Long, Width As Long, Height As Long) As Long
Attribute DrawImagelist.VB_MemberFlags = "40"
Dim Back As Long
Dim w As Long
Dim h As Long

ImageList_GetIconSize hIml, w, h
DrawBG TempDC, w, h
Back = ImageList_Draw(hIml, PicNr, TempDC, 0, 0, 0)
StretchBlt PicHdc, Left, Top, Width, Height, TempDC, 0, 0, w, h, vbSrcCopy
End Function

Private Sub CreateTempHdc(Width As Long, Height As Long)
Dim ScreenDc As Long

ScreenDc = GetDC(0)
TempDC = CreateCompatibleDC(ScreenDc)
hbmpmono = CreateCompatibleBitmap(ScreenDc, Width, Height)
hbmpOld = SelectObject(TempDC, hbmpmono)
DeleteDC ScreenDc
End Sub

Private Sub DeleteTempDC()
SelectObject TempDC, hbmpOld
DeleteDC TempDC
DeleteObject hbmpmono
DeleteObject hbmpOld
End Sub

Private Sub DrawBG(dc As Long, Width As Long, Height As Long)
Dim rc As RECT
rc.Right = Width
rc.Bottom = Height
FillRect dc, rc, BrushBG
End Sub

Private Function GetimageTyp(Path As String, oWidth As Long, oHeight As Long) As Long
Dim Filenr As Long
Dim Tempbyte As Byte
Dim Testint As Integer
Filenr = FreeFile
Open Path For Binary As Filenr
Get Filenr, 3, Testint
Select Case Testint
Case 1, 2
GetimageTyp = Testint
Get Filenr, 7, Tempbyte
oWidth = Tempbyte
Get Filenr, 8, Tempbyte
oHeight = Tempbyte
Case Else
GetimageTyp = -1
End Select
Close Filenr
If oWidth = 0 Or oHeight = 0 Then 'Standard
oWidth = 32
oHeight = 32
End If
End Function
