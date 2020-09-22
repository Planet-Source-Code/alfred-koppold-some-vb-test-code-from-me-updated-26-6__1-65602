VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ImageDrawConstants
imlNormal = 0
imlTransparent = 1
imlSelected = 2
imlFocus = 3
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PictureInfo
    cbSizeofStruct      As Long
    picType             As Long
    hImage              As Long
    xExt                As Long
    yExt                As Long
End Type

Public Index As Integer
Public Key As String
Public Picture As IPictureDisp
Public Tag As Variant
Private Const COLOR_HIGHLIGHT = 13

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictureInfo As PictureInfo, riid As GUID, ByVal fown As Long, ipic As IPicture) As Long

Public Sub Draw(hdc As OLE_HANDLE, Optional x As Long = 0, Optional y As Long = 0, Optional Style As ImageDrawConstants = imlNormal)
Dim PicHeight As Long
Dim Picwidth As Long
Dim Picdc As Long
Dim Transparentcolor As Long
Dim rc As RECT

PicHeight = Picture.Height * 0.5669 / Screen.TwipsPerPixelY
Picwidth = Picture.Width * 0.5669 / Screen.TwipsPerPixelX
Picdc = CreateCompatibleDC(hdc)
SelectObject Picdc, Picture.handle
Transparentcolor = RGB(255, 255, 255)

Select Case Style
Case 0
'imlNormal 0 (Voreinstellung) Normal. Zeichnet die Abbildung ohne Änderung.
StretchBlt hdc, x, y, Picwidth, PicHeight, Picdc, 0, 0, Picwidth, PicHeight, vbSrcCopy
Case 1
rc.Right = Picwidth
rc.Bottom = PicHeight
'imlTransparent 1 Transparent. Zeichnet die Abbildung mit der MaskColor-Eigenschaft, um zu ermitteln, welche Farbe der Abbildung transparent ist.
TransparentBlt hdc, hdc, Picdc, rc, 0, 0, Transparentcolor, 0
Case 2
'imlSelected 2 Ausgewählt. Zeichnet die Abbildung so, daß sie mit der Systemfarbe für Hervorhebungen gemischt wird.
CreateSelFoc hdc, Picwidth, PicHeight, 0, Picdc
Case 3
'imlFocus 3 Fokus. Zeichnet die Abbildung so, daß sie mit der Farbe für Hervorhebungen gemischt und mit einem Streifenmuster versehen wird, wodurch der Eindruck einer Schraffur entsteht, der anzeigt, daß die Abbildung den Fokus hat.
CreateSelFoc hdc, Picwidth, PicHeight, 1, Picdc
End Select
DeleteDC Picdc
End Sub

Public Function ExtractIcon() As IPictureDisp
Dim PicHeight As Long
Dim Picwidth As Long
Dim Transparentcolor As Long
Dim bm As Long
Dim ScreenDc As Long
Dim dc As Long
Dim oldBM As Long
Dim rc As RECT
Dim hBrush As Long
Dim il As Long
Dim Back As Long
Dim Picdc As Long
Dim hIcon As Long

PicHeight = Picture.Height * 0.5669 / Screen.TwipsPerPixelY
Picwidth = Picture.Width * 0.5669 / Screen.TwipsPerPixelX
hBrush = CreateSolidBrush(vbBlack)
Transparentcolor = vbWhite
ScreenDc = GetDC(0)
dc = CreateCompatibleDC(ScreenDc)
Picdc = CreateCompatibleDC(ScreenDc)
bm = CreateCompatibleBitmap(ScreenDc, 32, 32)
Back = SelectObject(dc, Picture.handle)
oldBM = SelectObject(Picdc, bm)
StretchBlt Picdc, 0, 0, 32, 32, dc, 0, 0, Picwidth, PicHeight, vbSrcCopy
Back = SelectObject(dc, Back)
bm = SelectObject(Picdc, oldBM)
il = CreateImageList(32, 32)
Back = ImageListAdd(il, "", Picwidth, PicHeight, bm)
hIcon = ImageList_GetIcon(il, 0, 1)
DestroyImageList il
Set ExtractIcon = CreateIconIPicture(hIcon)
DeleteObject hBrush
DeleteDC ScreenDc
DeleteDC dc
DeleteDC Picdc
DeleteObject bm
End Function

Private Sub TransparentBlt(OutDstDC As Long, DstDC As Long, SrcDC As Long, SrcRect As RECT, DstX As Long, DstY As Long, TransColor As Long, Typ As Long)
Dim nRet As Long, w As Integer, h As Integer
Dim MonoMaskDC As Long, hMonoMask As Long
Dim MonoInvDC As Long, hMonoInv As Long
Dim ResultDstDC As Long, hResultDst As Long
Dim ResultSrcDC As Long, hResultSrc As Long
Dim hPrevMask As Long, hPrevInv As Long
Dim hPrevSrc As Long, hPrevDst As Long
Dim Mask1 As Long
Dim Mask2 As Long

  w = SrcRect.Right - SrcRect.Left
  h = SrcRect.Bottom - SrcRect.Top
   
 'create monochrome mask and inverse masks
  MonoMaskDC = CreateCompatibleDC(DstDC)
  MonoInvDC = CreateCompatibleDC(DstDC)
  hMonoMask = CreateBitmap(w, h, 1, 1, ByVal 0&)
  hMonoInv = CreateBitmap(w, h, 1, 1, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
 'create keeper DCs and bitmaps
  ResultDstDC = CreateCompatibleDC(DstDC)
  ResultSrcDC = CreateCompatibleDC(DstDC)
  hResultDst = CreateCompatibleBitmap(DstDC, w, h)
  hResultSrc = CreateCompatibleBitmap(DstDC, w, h)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   
'copy src to monochrome mask
  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, TransColor)
  nRet = BitBlt(MonoMaskDC, 0, 0, w, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
  TransColor = SetBkColor(SrcDC, OldBC)
 'create inverse of mask
  nRet = BitBlt(MonoInvDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbNotSrcCopy)

 'get background
  nRet = BitBlt(ResultDstDC, 0, 0, w, h, DstDC, DstX, DstY, vbSrcCopy)
  Select Case Typ
  Case 0
  Mask1 = MonoMaskDC
  Mask2 = MonoInvDC
  Case 1, 2
  Mask1 = MonoInvDC
  Mask2 = MonoMaskDC
  End Select
  
 
Select Case Typ
Case 1
BitBlt ResultDstDC, 0, 0, w, h, SrcDC, 0, 0, vbSrcInvert
BitBlt ResultDstDC, 0, 0, w, h, Mask1, 0, 0, vbSrcAnd
BitBlt ResultDstDC, 0, 0, w, h, SrcDC, 0, 0, vbSrcInvert
OldBC = SetBkColor(SrcDC, TransColor)
nRet = BitBlt(MonoMaskDC, 0, 0, w, h, ResultDstDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
TransColor = SetBkColor(SrcDC, OldBC)
nRet = BitBlt(MonoInvDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbNotSrcCopy)
BitBlt ResultSrcDC, 0, 0, w, h, SrcDC, 0, 0, vbSrcCopy
BitBlt ResultDstDC, 0, 0, w, h, ResultSrcDC, 0, 0, vbSrcInvert
BitBlt ResultDstDC, 0, 0, w, h, Mask1, 0, 0, vbSrcAnd
BitBlt ResultDstDC, 0, 0, w, h, ResultSrcDC, 0, 0, vbSrcInvert
Case 0, 2
  nRet = BitBlt(ResultDstDC, 0, 0, w, h, Mask1, 0, 0, vbSrcAnd)
  nRet = BitBlt(ResultSrcDC, 0, 0, w, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
  nRet = BitBlt(ResultSrcDC, 0, 0, w, h, Mask2, 0, 0, vbSrcAnd)
  nRet = BitBlt(ResultDstDC, 0, 0, w, h, ResultSrcDC, 0, 0, vbSrcInvert)
End Select
 'output results
  nRet = BitBlt(OutDstDC, DstX, DstY, w, h, ResultDstDC, 0, 0, vbSrcCopy)
                
 'clean up
  hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
  DeleteObject hMonoMask
  hMonoInv = SelectObject(MonoInvDC, hPrevInv)
  DeleteObject hMonoInv
  hResultDst = SelectObject(ResultDstDC, hPrevDst)
  DeleteObject hResultDst
  hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
  DeleteObject hResultSrc
  DeleteDC MonoMaskDC
  DeleteDC MonoInvDC
  DeleteDC ResultDstDC
  DeleteDC ResultSrcDC

End Sub

Private Function CreateSelFoc(hdc As Long, Width As Long, Height As Long, Typ As Long, Picdc As Long) As IPictureDisp
Dim bm As Long
Dim BrushDc As Long
Dim dc As Long
Dim i As Long
Dim z As Long
Dim Brush As Long
Dim hOldbm As Long
Dim u As Long
Dim rc As RECT
Dim Colorbrush As Long
Dim HLColor As Long
Dim bmA As Long

HLColor = GetSysColor(COLOR_HIGHLIGHT)
bm = CreateCompatibleBitmap(hdc, 8, 8)
BrushDc = CreateCompatibleDC(hdc)
hOldbm = SelectObject(BrushDc, bm)
rc.Right = 8
rc.Bottom = 8
Colorbrush = CreateSolidBrush(vbWhite)
FillRect BrushDc, rc, Colorbrush
DeleteObject Colorbrush
u = 1
Select Case Typ
Case 0 'Select
For i = 0 To 7  'Height
For z = u To 7 Step 2  'Width
SetPixel BrushDc, z, i, HLColor
Next z
Select Case u
Case 0
u = 1
Case 1
u = 0
End Select
Next i
Case 1 'Focus
For i = 1 To 7 Step 2 'width
For z = 0 To 7 'height
SetPixel BrushDc, i, z, HLColor
Next z
Next i
u = 1
For i = 0 To 7 Step 2 'width
For z = u To 7 Step 2 'Height
SetPixel BrushDc, i, z, HLColor
Next z
Select Case u
Case 0
u = 1
Case 1
u = 0
End Select
Next i
End Select

bm = SelectObject(BrushDc, hOldbm)
Brush = CreatePatternBrush(bm)
rc.Right = Width
rc.Bottom = Height
bmA = CreateCompatibleBitmap(hdc, Width, Height)
dc = CreateCompatibleDC(hdc)
SelectObject dc, bmA
FillRect dc, rc, Brush
TransparentBlt hdc, dc, Picdc, rc, 0, 0, vbWhite, 1
DeleteObject Brush
DeleteDC dc
DeleteObject bmA
End Function

Private Function CreateIconIPicture(ByVal hIcon As Long) As IPictureDisp
Dim objPicture As IPictureDisp
Dim udtGuid  As GUID
Dim udtPictureInfo As PictureInfo

     With udtGuid
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
    End With
    With udtPictureInfo
        .cbSizeofStruct = Len(udtPictureInfo)
        .picType = 3
        .hImage = hIcon
    End With
    OleCreatePictureIndirect udtPictureInfo, udtGuid, 1, objPicture
    Set CreateIconIPicture = objPicture
    Set objPicture = Nothing
End Function


Private Sub Class_Initialize()

End Sub
