VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Palette8Bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum PaletteTypes
    PaletteSystemDefault
    PaletteHalftone
    PaletteWebSafe
    PaletteGray
    PaletteOptimal
    PaletteCustom
End Enum

Private Enum eColorDepth
    clrUnknown = 0
    clr1bpp = 1
    clr4bpp = 4
    clr8bpp = 8
    clr16bpp = 16
    clr24bpp = 24
    clr32bpp = 32
    clr48bpp = 48
    clr64bpp = 64
End Enum
Private Type tNode       ' Octree node struct
   bIsLeaf       As Long ' Leaf flag
   bAddedReduce  As Long ' Linked list flag
   cClrs         As Long ' Pixel count
   vRed          As Long ' Cumulative color values
   vGreen        As Long
   vBlue         As Long
   iChildren(7)  As Long ' Child nodes
   iNext         As Long ' Next reducable node
   iPaletteIndex As Long ' Stored palette index, for fast search
End Type

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BITMAPINFO 'I use Dynamic array which allow to work with ANY color depth
        bmiHeader As BITMAPINFOHEADER
        bmiColors() As RGBQUAD
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPal As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private cNodes     As Long  ' Total number of nodes
Private TopGarbage As Long  ' Index of next item to recycle
Private m_iPalette As Long  ' Number of palette entries
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private m_Err As Long
Private m_bOptimal As Boolean
Private aReduce(1 To 8) As Long    ' Reducable node for a given level
Private mask(7)         As Long    ' Bitmasks
Private Value(128)      As Byte    ' Precalculated values
Private aNodes()        As tNode   ' Array of nodes
Private m_tPal()        As RGBQUAD ' Array of colors
Private m_hPal          As Long    ' Handle to palette
Private Const ITER As Long = 1023  ' Create 1K nodes at a time

Public Property Get Count() As Long
   Count = m_iPalette
End Property

Public Property Let Count(vData As Long)
   m_iPalette = vData
   ReDim Preserve m_tPal(vData - 1)
End Property


Public Property Get IndexColor(ByVal iIndex As Long) As Long
   If iIndex > m_iPalette - 1 Then Exit Property
   IndexColor = RGB(m_tPal(iIndex).rgbBlue, m_tPal(iIndex).rgbGreen, m_tPal(iIndex).rgbRed)
End Property

Public Sub GetPaletteColors(ByVal lIndex As Long, ByRef red As Byte, ByRef green As Byte, ByRef blue As Byte)
   If lIndex > m_iPalette - 1 Then Exit Sub
   red = m_tPal(lIndex).rgbRed
   green = m_tPal(lIndex).rgbGreen
   blue = m_tPal(lIndex).rgbBlue
End Sub

Private Sub BuildPalette()
   Dim abPal() As Byte
   Dim lVer As Integer
   lVer = &H300
   ReDim abPal(m_iPalette * 4 + 3)
   CopyMemory abPal(0), lVer, 2
   CopyMemory abPal(2), m_iPalette, 2
   CopyMemory abPal(4), m_tPal(0), m_iPalette * 4
   m_hPal = CreatePalette(abPal(0))
End Sub

Private Sub AddColor(ByVal iBranch As Long, ByVal nLevel As Long, ByVal nDepth As Long, ByRef cClr As Long, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
  Dim iIndex As Long
  Dim lBGR As Long ' bugger! :P
  Dim lMask As Long
  
  lMask = mask(nLevel - 1)
  lBGR = (Value(b And lMask) * 4) Or (Value(g And lMask) * 2) Or (Value(r And lMask))
  
   If aNodes(iBranch).iChildren(lBGR) = 0 Then  'If no child here then...
      iIndex = GetFreeNode 'Get a new node index
      aNodes(iBranch).iChildren(lBGR) = iIndex
      aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
      With aNodes(iIndex)   'Clear/set data
         .bIsLeaf = (nLevel = nDepth)
         .cClrs = 0
         .vRed = 0
         .vGreen = 0
         .vBlue = 0
         .iNext = 0
      End With
   Else  'Has a child here
      iIndex = aNodes(iBranch).iChildren(lBGR)
   End If
   
  If aNodes(iIndex).bIsLeaf = False Then  'If it is not a leaf
    With aNodes(iIndex)
      If .bAddedReduce = False Then    'If 2 or more colors, add to reducable aNodes list
        .iNext = aReduce(nLevel)
        aReduce(nLevel) = iIndex
        .bAddedReduce = True
      End If
    End With
      
    'Search a level deeper
    AddColor iIndex, nLevel + 1, nDepth, cClr, r, g, b
  Else  'it's a leaf
    With aNodes(iIndex)
      If .cClrs = 0 Then cClr = cClr + 1
        .cClrs = .cClrs + 1
        .vRed = .vRed + r
        .vGreen = .vGreen + g
        .vBlue = .vBlue + b
    End With
  End If
End Sub

Public Function ClosestIndex(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Long
    ClosestIndex = GetNearestPaletteIndex(m_hPal, VBA.RGB(red, green, blue))
End Function

' Function used in getting a palette index for 2/4/8 bit palettes.
' A better algorithm could be devised to move up the tree,
' and search for a closer color.
Public Function ClosestIndexOptimal(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte) As Long
  Dim lLevel As Long
  Dim lShift As Long
  Dim lIndex As Long
  Dim lChild As Long
  Dim lPrevIndex As Long
  Dim i As Long
  Dim lMask As Long

  lIndex = 1 ' start at root

  Do
    lShift = 7 - lLevel
    lMask = mask(lLevel)
    lChild = (Value(blue And lMask) * 4) Or (Value(green And lMask) * 2) Or (Value(red And lMask))
    lPrevIndex = lIndex
    lIndex = aNodes(lIndex).iChildren(lChild) ' BGR
    Do While lIndex = 0 ' rarely happens
      lIndex = aNodes(lPrevIndex).iChildren(i) ' attempt to find a similar color
      i = i + 1
      If i = 8 Then Exit Function ' how about the first index?
    Loop
    lLevel = lLevel + 1
  Loop Until aNodes(lIndex).bIsLeaf
 
  ClosestIndexOptimal = aNodes(lIndex).iPaletteIndex
End Function

' Function to get the individual colors for a non-paletted image
' Returns colors ByRef
Public Sub ClosestIndexColors(ByRef red As Byte, ByRef green As Byte, ByRef blue As Byte)
  Dim lLevel As Long
  Dim lShift As Long
  Dim lIndex As Long
  Dim lChild As Long
  Dim lPrevIndex As Long
  Dim i As Long
  Dim lMask As Long
  
  lIndex = 1 ' start at root
  
  Do
    lShift = 7 - lLevel
    lMask = mask(lLevel)
    lChild = (Value(blue And lMask) * 4) Or (Value(green And lMask) * 2) Or (Value(red And lMask))

    lPrevIndex = lIndex
    lIndex = aNodes(lIndex).iChildren(lChild) ' BGR
    Do While lIndex = 0 ' rarely happens
      lIndex = aNodes(lPrevIndex).iChildren(i) ' attempt to find a similar color
      i = i + 1
      If i = 8 Then Exit Sub      ' how about black?
    Loop
    lLevel = lLevel + 1
  Loop While aNodes(lIndex).bIsLeaf = False
  
  i = aNodes(lIndex).iPaletteIndex
  red = m_tPal(i).rgbRed
  green = m_tPal(i).rgbGreen
  blue = m_tPal(i).rgbBlue
End Sub

Private Function CombineNodes(ByRef lColors As Long, ByVal nDepth As Long) As Long
  Dim i As Long, iIndex As Long
  Dim nR As Long, nG As Long, nB As Long, nPixel As Long
  Dim lChildren As Long
   
   For i = nDepth To 1 Step -1   'Find deepest reducable level
      If aReduce(i) <> 0 Then Exit For
   Next i
   
   If i = 0 Then Exit Function
   iIndex = aReduce(i)
   aReduce(i) = aNodes(iIndex).iNext
   
   For i = 0 To 7
      If aNodes(iIndex).iChildren(i) <> 0 Then 'If there is a child
         With aNodes(aNodes(iIndex).iChildren(i))
            'Add red, green, blue, and pixel count to running total
            nR = nR + .vRed
            nG = nG + .vGreen
            nB = nB + .vBlue
            nPixel = nPixel + .cClrs
            FreeNode aNodes(iIndex).iChildren(i) 'Free the node
            lChildren = lChildren + 1
         End With
         aNodes(iIndex).iChildren(i) = 0 'Clear the link
      End If
   Next i
   
   lColors = lColors - (lChildren - 1)
   
   'Set the new node data
   With aNodes(iIndex)
      .cClrs = nPixel
      .bIsLeaf = True
      .vRed = nR
      .vBlue = nB
      .vGreen = nG
   End With
   CombineNodes = True
End Function

Public Sub CreateOptimal(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lPad As Long, _
                         ByRef lImage() As Byte, Optional ByVal PalSize As Long = 236, _
                         Optional ByVal nLevels As Long = 6)
  Dim x         As Long
  Dim y         As Long
  Dim lNodeIdx  As Long
  Dim cClr      As Long
  
  Reset ' prepare tree for use

  y = (lWidth * (lHeight - 1)) - 1

  Do
    x = lWidth - lPad
    Do
      AddColor 1, 1, nLevels, cClr, lImage(x + y - 2), lImage(x + y - 1), lImage(x + y)

      Do While cClr > PalSize
        If CombineNodes(cClr, nLevels) = False Then Exit Do
      Loop
      x = x - 3
    Loop While x >= 0
    y = y - lWidth
  Loop While y >= 0
   
  If cClr > 16 Then ' allow for 4 or 8 bit palettes
    cClr = 256
  Else
    cClr = 16
  End If
   
  ReDim m_tPal(cClr - 1) As RGBQUAD
  lNodeIdx = 1 ' beginning of tree
  SetPalette aNodes, m_tPal, m_iPalette, lNodeIdx
  m_iPalette = m_iPalette - 1
  BuildPalette
End Sub

Private Sub FreeNode(ByVal iNode As Long)
   aNodes(iNode).iNext = TopGarbage
   TopGarbage = iNode
   aNodes(iNode).bIsLeaf = False 'Necessary for final loop through
   aNodes(iNode).bAddedReduce = False
   cNodes = cNodes - 1
End Sub

Private Function GetFreeNode() As Long
  Dim i As Long
  
  cNodes = cNodes + 1
  If TopGarbage = 0 Then
    If (cNodes And ITER) = 0 Then
      ReDim Preserve aNodes(1 To cNodes + ITER + 1)
    End If
    GetFreeNode = cNodes
  Else
    GetFreeNode = TopGarbage
    TopGarbage = aNodes(TopGarbage).iNext
    For i = 0 To 7
      aNodes(GetFreeNode).iChildren(i) = 0
    Next i
  End If
End Function

' Reset and create lookup tables
Public Sub Reset()
  Dim i As Long
  
  Erase aReduce
  ReDim aNodes(1 To ITER + 1) 'Allocates initial storage
  TopGarbage = 0
  m_iPalette = 0
  cNodes = 1
   
  For i = 0 To 7
    mask(i) = 2 ^ (7 - i)
  Next i
  
  For i = 0 To 7
    Value(2 ^ i) = 1
  Next i
End Sub

' Recursively grab palette entries, makes for mostly-sorted order
' Caches the palette index for faster lookups
Private Sub SetPalette(ByRef Node() As tNode, ByRef tPal() As RGBQUAD, ByRef lColorIndex As Long, ByRef lNodeIndex As Long)
  Dim i         As Long
  Dim lColors   As Long
  Dim lChild    As Long
  
  If Node(lNodeIndex).bIsLeaf Then
    With Node(lNodeIndex)
      lColors = .cClrs
      tPal(lColorIndex).rgbRed = .vRed \ lColors   ' integer division, baby!
      tPal(lColorIndex).rgbGreen = .vGreen \ lColors
      tPal(lColorIndex).rgbBlue = .vBlue \ lColors
      .iPaletteIndex = lColorIndex              ' cache the index into the color table
      lColorIndex = lColorIndex + 1
    End With
  Else
    For i = 0 To 7 ' reverse this for light->dark palette
      lChild = Node(lNodeIndex).iChildren(i)
      If lChild Then SetPalette Node, tPal, lColorIndex, lChild
    Next i
  End If
End Sub

Private Sub Class_Initialize()
   m_iPalette = 0
   ReDim m_tPal(0) As RGBQUAD
End Sub

Private Sub Class_Terminate()
   If m_hPal Then DeleteObject m_hPal
End Sub

Public Function Save8Bpp(pbSrc As PictureBox, ByVal sFileName As String)
   Dim abPicture_8() As Byte
   Dim abPicture_24() As Byte
   Dim bi As BITMAPINFO
   Dim bfh As BITMAPFILEHEADER
   Dim lHeight As Long
   Dim ScanLine_8 As Long, lPad_8 As Long
   Dim ScanLine_24 As Long, lPad_24 As Long
   Dim lCounter As Long
   Dim x As Long, y As Long
   
   If GetTrueBits(pbSrc, abPicture_24, bi) Then
      ScanLine_24 = BytesPerScanLine(bi.bmiHeader.biWidth, clr24bpp)
      lPad_24 = ScanLine_24 - bi.bmiHeader.biWidth * 3
      ScanLine_8 = BytesPerScanLine(bi.bmiHeader.biWidth, clr8bpp)
      lPad_8 = ScanLine_8 - bi.bmiHeader.biWidth
      lHeight = bi.bmiHeader.biHeight
      ReDim abPicture_8(ScanLine_8 * lHeight - 1)
               CreateNew 1

         'CreateOptimal ScanLine_24, lHeight, lPad_24, abPicture_24()
    ' process Palette
      lCounter = ScanLine_8 * lHeight - 1 - lPad_8
      y = (ScanLine_24 * (lHeight - 1)) - 1
      On Error Resume Next
      Do
        x = ScanLine_24 - lPad_24
        Do
          ' store pointers into the palette
          abPicture_8(lCounter) = ClosestIndex(abPicture_24(x + y), _
                                                    abPicture_24(x + y - 1), _
                                                    abPicture_24(x + y - 2))
          lCounter = lCounter - 1
          x = x - 3
        Loop While x > 0
        lCounter = lCounter - lPad_8
        y = y - ScanLine_24
      Loop While y > -ScanLine_24
      On Error GoTo 0
      With bi.bmiHeader
         .biSizeImage = .biWidth * .biHeight
         .biBitCount = 8
         .biClrImportant = 256
         .biClrUsed = 256
         .biCompression = BI_RGB
      End With
      ReDim bi.bmiColors(255)
      Dim i As Long
      For i = 0 To Count - 1
          With bi.bmiColors(i)
               'Palette colors store in reverse mode (BGR)
               GetPaletteColors i, .rgbBlue, .rgbGreen, .rgbRed
               .rgbReserved = 0
          End With
      Next i
      bfh.bfType = 19778
      bfh.bfOffBits = Len(bfh) + Len(bi.bmiHeader) + Len(bi.bmiColors(0)) * 256
      bfh.bfSize = bfh.bfOffBits + UBound(abPicture_8) + 1
      Dim nFile As Integer
      nFile = FreeFile
      Open sFileName For Binary As #nFile
         Put #nFile, , bfh
         Put #nFile, , bi.bmiHeader
         Put #nFile, , bi.bmiColors
         Put #nFile, , abPicture_8
      Close #nFile
   End If
End Function
Private Function GetTrueBits(pb As PictureBox, abPicture() As Byte, bi As BITMAPINFO) As Boolean
   Dim bmp As BITMAP
   Call GetObjectAPI(pb.Picture, Len(bmp), bmp)
   ReDim bi.bmiColors(0)
   With bi.bmiHeader
       .biSize = Len(bi.bmiHeader)
       .biWidth = bmp.bmWidth
       .biHeight = bmp.bmHeight
       .biPlanes = 1
       .biBitCount = 24
       .biCompression = BI_RGB
       .biSizeImage = BytesPerScanLine(.biWidth, clr24bpp) * .biHeight
       ReDim abPicture(.biSizeImage - 1)
   End With
   GetTrueBits = GetDIBits(pb.hdc, pb.Picture, 0, bi.bmiHeader.biHeight, abPicture(0), bi, DIB_RGB_COLORS)
End Function

Private Function BytesPerScanLine(ByVal lWidth As Long, bpp As eColorDepth) As Long
   Select Case bpp
       Case clr1bpp
            BytesPerScanLine = ScanAlign((lWidth + 7) \ 8)
       Case clr4bpp
            BytesPerScanLine = ScanAlign((lWidth + 1) \ 2)
       Case clr8bpp
            BytesPerScanLine = ScanAlign(lWidth)
       Case clr16bpp
            BytesPerScanLine = ScanAlign(lWidth * 2)
       Case clr24bpp
            BytesPerScanLine = ScanAlign(lWidth * 3)
       Case clr32bpp
            BytesPerScanLine = ScanAlign(lWidth * 4)
   End Select
End Function

Private Function ScanAlign(ByVal lWidth As Long) As Long
    ScanAlign = (lWidth + 3) And &HFFFFFFFC
End Function

Public Sub CreateNew(PalType As PaletteTypes, Optional nColors As Long = 256)
   m_bOptimal = False
   Select Case PalType
'       Case PaletteMono
'            m_iPalette = 2
'            CreateMono m_tPal()
'       Case Palette_16
'            m_iPalette = 16
'            Create16Colors m_tPal()
       Case PaletteHalftone
            m_iPalette = 256
            CreateHalftone m_tPal()
   End Select
   BuildPalette
End Sub

Private Sub CreateHalftone(m_tPal() As RGBQUAD)
   Dim lIndex As Long
   Dim r As Long, g As Long, b As Long
   Dim rA As Long, gA As Long, bA As Long
   ReDim m_tPal(255)
   ' Halftone 256 colour palette
   For b = 0 To &H100 Step &H40
      If b = &H100 Then bA = b - 1 Else bA = b
      For g = 0 To &H100 Step &H40
         If g = &H100 Then gA = g - 1 Else gA = g
         For r = 0 To &H100 Step &H40
            If r = &H100 Then rA = r - 1 Else rA = r
            With m_tPal(lIndex)
               .rgbRed = rA: .rgbGreen = gA: .rgbBlue = bA
               lIndex = lIndex + 1
            End With
         Next r
      Next g
   Next b
End Sub


