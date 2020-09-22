Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Const SRCCOPY = &HCC0020

Public Function LoadBitmapIntoMemory(P As StdPicture) As Long
 Dim hBmp As Long
 Dim hBmpOld As Long
 Dim hDCDesk As Long
 Dim hdcTemp As Long
  hBmp = P.Handle
  hDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
  If (hDCDesk <> 0) Then
   hdcTemp = CreateCompatibleDC(hDCDesk)
   If (hdcTemp <> 0) Then
    hBmpOld = SelectObject(hdcTemp, hBmp)
          LoadBitmapIntoMemory = hdcTemp
End If
End If
hBmp = DeleteObject(hBmp)
If hBmp = 1 Then hBmp = 0
hBmpOld = DeleteObject(hBmpOld)
If hBmpOld = 1 Then hBmpOld = 0
hDCDesk = DeleteDC(hDCDesk)
If hDCDesk = 1 Then hDCDesk = 0

End Function

Public Sub TransparentBlt(OutDstDC As Long, _
                           DstDC As Long, _
                           SrcDC As Long, _
                           SrcRect As RECT, _
                           DstX As Long, _
                           DstY As Long, _
                           TransColor As Long)
   
  'DstDC- Device context into which image must be
  'drawn transparently
  
  'OutDstDC- Device context into image is actually drawn,
  'even though it is made transparent in terms of DstDC

  'Src- Device context of source to be made transparent
  'in color TransColor

  'SrcRect- Rectangular region within SrcDC to be made
  'transparent in terms of DstDC, and drawn to OutDstDC

  'DstX, DstY - Coordinates in OutDstDC (and DstDC)
  'where the transparent bitmap must go. In most
  'cases, OutDstDC and DstDC will be the same
  Dim nRet As Long, W As Integer, H As Integer
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long
  Dim hPrevSrc As Long, hPrevDst As Long

  W = SrcRect.Right - SrcRect.Left + 1
  H = SrcRect.Bottom - SrcRect.Top + 1
   
 'create monochrome mask and inverse masks
  MonoMaskDC = CreateCompatibleDC(DstDC)
  MonoInvDC = CreateCompatibleDC(DstDC)
  hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
 'create keeper DCs and bitmaps
  ResultDstDC = CreateCompatibleDC(DstDC)
  ResultSrcDC = CreateCompatibleDC(DstDC)
  hResultDst = CreateCompatibleBitmap(DstDC, W, H)
  hResultSrc = CreateCompatibleBitmap(DstDC, W, H)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   
'copy src to monochrome mask
  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, TransColor)
  nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, _
                SrcRect.Left, SrcRect.Top, vbSrcCopy)
  TransColor = SetBkColor(SrcDC, OldBC)
   
 'create inverse of mask
  nRet = BitBlt(MonoInvDC, 0, 0, W, H, _
                MonoMaskDC, 0, 0, vbNotSrcCopy)
   
 'get background
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                DstDC, DstX, DstY, vbSrcCopy)
   
 'AND with Monochrome mask
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                MonoMaskDC, 0, 0, vbSrcAnd)
   
 'get overlapper
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, _
                SrcRect.Left, SrcRect.Top, vbSrcCopy)
   
 'AND with inverse monochrome mask
  nRet = BitBlt(ResultSrcDC, 0, 0, W, H, _
                MonoInvDC, 0, 0, vbSrcAnd)
   
'XOR these two
  nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
                ResultSrcDC, 0, 0, vbSrcInvert)
   
 'output results
  nRet = BitBlt(OutDstDC, DstX, DstY, W, H, _
                ResultDstDC, 0, 0, vbSrcCopy)
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


