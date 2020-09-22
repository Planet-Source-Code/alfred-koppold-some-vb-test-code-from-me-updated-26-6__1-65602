VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SavePNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DIB_RGB_COLORS            As Long = 0
Const BMPBIT                    As Long = 50
Const BI_RGB                    As Long = 0
Const MinMatch                  As Long = 3

Public bkgdColorLong                As Long
Public HasbkgdColor             As Boolean
Public HasAlpha                 As Boolean
Public HasTrans                 As Boolean
Public Transparent24            As Long

Dim BitCnt                      As Long
Dim ArrUBnd                     As Long
Dim Bildhöhe                    As Long
Dim Bildbreite                  As Long
Dim BildbreiteRight             As Long
Dim CompPNG()                   As Byte
Dim Bitcbr                      As Long

Dim MaxMatch                    As Long
Dim MaxHistory                  As Long
Dim NxtFreeSlot                 As Long
Dim NxtFreeHashSlot1            As Long
Dim NxtFreeHashSlot2            As Long
Dim NxtFreeHash                 As Long
Dim Hash1()                     As Long
Dim HashSlot1()                 As HashSlotType
Dim hashSlot2()                 As HashSlotType
Dim Hash()                      As HashTableType
Dim Slot()                      As SlotType
Dim Udata                       As DataType
Dim IData                       As DataType
Dim GoodMatch                   As Long
Dim LazyMatch                   As Long
Dim NiceMatch                   As Long
Dim MaxChain                    As Long
Dim MatchL                      As Long
Dim MatchP                      As Long
Dim FileLength                  As Long
Dim LL(287)                     As Long
Dim DC(31)                      As Long
Dim Lens(31)                    As Long
Dim Lext(31)                    As Long
Dim Dist(31)                    As Long
Dim Dext(31)                    As Long
Dim Border(18)                  As Long
Dim LitLen(287)                 As CodesType
Dim Distance(31)                As CodesType
Dim Bl_Tree()                   As CodesType
Dim Max_BL                      As Long
Dim MaxLitLen                   As Long
Dim MaxDist                     As Long
Dim NumDist                     As Long
Dim m_AlphaSet                  As Long
Dim m_bkgdcolorLong                 As Long
Dim AlphaArray()                As Byte
Dim TransArray8() As Byte
Private Type IHDR
    Width                       As Long
    Height                      As Long
    bpp                         As Byte
    ColorType                   As Byte
    CompressionType             As Byte
    FilterType                  As Byte
    InterlaceFlag               As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize                      As Long
    biWidth                     As Long
    biHeight                    As Long
    biPlanes                    As Integer
    biBitCount                  As Integer
    biCompression               As Long
    biSizeImage                 As Long
    biXPelsPerMeter             As Long
    biYPelsPerMeter             As Long
    biClrUsed                   As Long
    biClrImportant              As Long
End Type

Private Type RGBTRE
    rgbBlue                     As Byte
    rgbGreen                    As Byte
    rgbRed                      As Byte
End Type

Private Type RGBQUAD
    rgbBlue                     As Byte
    rgbGreen                    As Byte
    rgbRed                      As Byte
    rgbReserved                 As Byte
End Type

Private Type BITMAPINFO8Bit
    bmiHeader                   As BITMAPINFOHEADER
    bmiColors(0 To 255)         As RGBQUAD
End Type

Private Type BITMAPINFO
    bmiHeader                   As BITMAPINFOHEADER
    bmiColors                   As RGBQUAD
End Type

'For Comression
Private Type HashSlotType
    Char                        As Byte
    ToHash                      As Long
    ToSlot                      As Long
End Type

Private Type HashTableType
    FirstSlot                   As Long
    LastSlot                    As Long
    Pointer                     As Long
    NumSlots                    As Long
End Type

Private Type SlotType
    Next                        As Long
    Prev                        As Long
    Winpos                      As Long
    Reserved                    As Long     ' adapt to 16 bytes
End Type

Private Type DataType
    bar()                       As Byte
    pos                         As Long
    bit                         As Long
    val                         As Long
End Type

Private Type CodesType
    Code                        As Long
    Length                      As Long
End Type


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Compress Lib "ZLIB.DLL" Alias "compress2" (ByRef DestinationArray As Byte, ByRef DestLen As Long, ByRef SourceArray As Byte, ByVal SourceLen As Long, ByVal CompressionLevel As Long) As Long
Private Declare Function CompressN Lib "ZLIB.DLL" Alias "compress" (ByRef DestinationArray As Byte, ByRef DestLen As Long, ByRef SourceArray As Byte, ByVal SourceLen As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Sub memCPY Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function Calc_Dynamic() As Long
 Dim x As Long
 Dim Temp As Long
 Dim Freq(18) As Long
 Call OptimizeHuffman(LitLen, LL, 287, 15)
 Call OptimizeHuffman(Distance, DC, NumDist, 15)
 For x = 0 To 256
 Temp = Temp + LL(x) * LitLen(x).Length
 Next
 For x = 257 To 287
 Temp = Temp + LL(x) * LitLen(x).Length + LL(x) * Lext(x - 257)
 Next
 For x = 0 To NumDist
 Temp = Temp + DC(x) * Distance(x).Length + DC(x) * Dext(x)
 Next
 ReDim Bl_Tree(18)
 For MaxLitLen = 287 To 0 Step -1
 If LitLen(MaxLitLen).Length <> 0 Then Exit For
 Next
 Call Scan_Tree(LitLen, 287)
 For MaxDist = NumDist To 0 Step -1
 If Distance(MaxDist).Length <> 0 Then Exit For
 Next
 Call Scan_Tree(Distance, NumDist)
 For x = 0 To 18
 Freq(x) = Bl_Tree(x).Code
 Next
 ReDim Bl_Tree(18)
 Call OptimizeHuffman(Bl_Tree, Freq, 18, 7)
 For Max_BL = 18 To 0 Step -1
 If Bl_Tree(Border(Max_BL)).Length <> 0 Then Exit For
 Next
 Temp = Temp + Max_BL * 3
 For x = 0 To 15
 Temp = Temp + Freq(x) * Bl_Tree(x).Length
 Next
 Temp = Temp + Freq(16) * (Bl_Tree(16).Length + 2)
 Temp = Temp + Freq(17) * (Bl_Tree(17).Length + 3)
 Temp = Temp + Freq(18) * (Bl_Tree(18).Length + 7)
 Calc_Dynamic = Fix((Temp + 3) / 8)
End Function

Private Function Calc_Static() As Long
 Dim x As Long
 Dim Temp As Long
 For x = 0 To 143
 Temp = Temp + LL(x) * 8
 Next
 For x = 144 To 255
 Temp = Temp + LL(x) * 9
 Next
 Temp = Temp + LL(256) * 7
 For x = 257 To 279
 Temp = Temp + LL(x) * 7 + LL(x) * Lext(x - 257)
 Next
 For x = 280 To 287
 Temp = Temp + LL(x) * 8 + LL(x) * Lext(x - 257)
 Next
 For x = 0 To 31
 Temp = Temp + DC(x) * 5 + DC(x) * Dext(x)
 Next
 Calc_Static = Fix((Temp + 3) / 8)
End Function

Public Sub Compress_with_VB(Bytearray() As Byte, CompLevel As Long, Optional Deflate64 As Boolean = False)
    Dim HoldPos     As Long
    Dim LitCount    As Long
    Dim PrevMatch   As Long
    Dim PrevPos     As Long
    Dim Char        As Long
    Dim BestStore   As Long
    Dim BestStatic  As Long
    Dim BestDynamic As Long
    Dim Bestcase    As Long
    Dim Header      As Byte
    Dim tb          As Byte
    Dim StLen       As Long
    Dim x           As Long
 
    Call INIT(CompLevel, Deflate64)
    Call init_Trees
    
    FileLength = UBound(Bytearray)
    IData.bar = Bytearray
    Erase Bytearray
    IData.pos = 0
    
    If CompLevel > 0 Then
        Do While IData.pos + MinMatch <= FileLength
            Call FindLongestMatch(IData.pos, GoodMatch, True)
            
            'If MatchL = GoodMatch Then
                PrevMatch = MatchL
                PrevPos = MatchP
                Call FindLongestMatch(IData.pos, LazyMatch, False)
                
                If MatchL <= PrevMatch Then
                    MatchL = PrevMatch
                    MatchP = PrevPos
                End If
            'End If
 
            If CompLevel < 4 Then
            'ElseIf MatchL = LazyMatch Then
                PrevMatch = MatchL
                PrevPos = MatchP
                Call FindLongestMatch(IData.pos + 1, NiceMatch, True)
                
                If MatchL <= PrevMatch Then
                    MatchL = PrevMatch
                    MatchP = PrevPos
                
                Else
                    Char = IData.bar(IData.pos)
                    If LitCount = 127 Then Udata.bar(HoldPos) = LitCount: LitCount = 0
                    If LitCount = 0 Then HoldPos = PutByte(0)
                    Call InsertSlot(IData.pos)
                    Call PutByte(Char)
                    LL(Char) = LL(Char) + 1
                    LitCount = LitCount + 1
                    IData.pos = IData.pos + 1
                End If
            End If
 
            If MatchL = MinMatch And IData.pos - MatchP > 4096 Then MatchL = 1
            If MatchL < MinMatch Then
                Char = IData.bar(IData.pos)
                If LitCount = 127 Then Udata.bar(HoldPos) = LitCount: LitCount = 0
                If LitCount = 0 Then HoldPos = PutByte(0)
                Call InsertSlot(IData.pos)
                Call PutByte(Char)
                LL(Char) = LL(Char) + 1
                LitCount = LitCount + 1
                IData.pos = IData.pos + 1
            
            Else
                If LitCount > 0 Then Udata.bar(HoldPos) = LitCount: LitCount = 0
                PutByte 128
                MatchP = IData.pos - MatchP
                Call PutByte(Fix(MatchP / 256))
                Call PutByte(MatchP And 255)
                x_lngINC MatchP, NumDist, Dist, DC, 0
                
                If MatchL > 257 And Deflate64 Then
                    Call PutByte(258 - 3)
                    Call PutByte(Fix((MatchL - 3) / 256))
                    Call PutByte((MatchL - 3) And 255)
                    x_lngINC 258, 28, Lens, LL, 257

                Else
                    Call PutByte(MatchL - 3)
                    x_lngINC MatchL, 28, Lens, LL, 257
                End If
 
                Do While MatchL > 0
                    Call InsertSlot(IData.pos)
                    If IData.pos = FileLength - 2 Then IData.pos = IData.pos + MatchL: Exit Do
                    IData.pos = IData.pos + 1
                    MatchL = MatchL - 1
                Loop
            End If
        Loop
 
        If IData.pos <= FileLength Then
            For x = IData.pos To FileLength
                If LitCount = 255 Then Udata.bar(HoldPos) = LitCount: LitCount = 0
                If LitCount = 0 Then HoldPos = PutByte(0)
                Call PutByte(CInt(IData.bar(x)))
                LL(IData.bar(x)) = LL(IData.bar(x))
                LitCount = LitCount + 1
            Next
        End If
 
        If LitCount > 0 Then Udata.bar(HoldPos) = LitCount
        Call PutByte(0)
        LL(256) = 1
        ReDim Preserve Udata.bar(Udata.pos - 1)
    End If
 
    If CompLevel <> 0 Then
        BestStore = UBound(IData.bar) + 5
        BestStatic = Calc_Static
        BestDynamic = Calc_Dynamic
        Bestcase = 0
        If BestStatic < BestStore Then Bestcase = 1
        If BestDynamic < BestStatic Then Bestcase = 2
    
    Else
        Bestcase = 0
    End If
 
    Call putbits(1, 1)
    Call putbits(Bestcase, 2)
    
    If Bestcase = 0 Then
        Do While Udata.bit <> 0
            Call putbits(0, 1)
        Loop
        
        StLen = FileLength + 1
        Call PutByte(StLen And &HFF)
        Call PutByte(Fix(StLen / 256) And &HFF)
        StLen = (Not StLen) And &HFFFF&
        Call PutByte(StLen And &HFF)
        Call PutByte(Fix(StLen / 256) And &HFF)
 
        For x = 0 To UBound(IData.bar)
            PutByte (IData.bar(x))
        Next
    
    Else
        IData.bar = Udata.bar
        IData.pos = 0
        Udata.pos = 0
        If Bestcase = 1 Then
            Erase LitLen
            Erase Distance
            For x = 0 To 143: LL(x) = 8: Next
            For x = 144 To 255: LL(x) = 9: Next
            For x = 256 To 279: LL(x) = 7: Next
            For x = 280 To 287: LL(x) = 8: Next
            For x = 0 To 31: DC(x) = 5: Next
            Call Create_Codes(LitLen, LL, 287)
            Call Create_Codes(Distance, DC, 31)
        
        Else
            Call putbits(MaxLitLen + 1 - 257, 5)
            Call putbits(MaxDist + 1 - 1, 5)
            Call putbits(Max_BL + 1 - 4, 4)
            
            For x = 0 To Max_BL
                Call putbits(Bl_Tree(Border(x)).Length, 3)
            Next
 
            Call Send_Tree(LitLen, MaxLitLen)
            Call Send_Tree(Distance, MaxDist)
        End If
 
        Do
            Header = IData.bar(IData.pos)
            IData.pos = IData.pos + 1
            
            If Header < 128 Then
                If Header = 0 Then
                    Call PutBitsRev(LitLen(256).Code, LitLen(256).Length)
                    Exit Do
                End If
 
                For x = 1 To Header
                    tb = IData.bar(IData.pos)
                    IData.pos = IData.pos + 1
                    Call PutBitsRev(LitLen(tb).Code, LitLen(tb).Length)
                Next
 
            Else
                MatchP = CLng(IData.bar(IData.pos)) * 256 + IData.bar(IData.pos + 1)
                IData.pos = IData.pos + 2
                MatchL = IData.bar(IData.pos) + 3
                IData.pos = IData.pos + 1
                x = 28
                
                Do While Lens(x) > MatchL
                    x = x - 1
                Loop
 
                Call PutBitsRev(LitLen(x + 257).Code, LitLen(x + 257).Length)
                Call putbits(MatchL - Lens(x), Lext(x))
 
                If MatchL = 258 And Deflate64 Then
                    MatchL = CLng(IData.bar(IData.pos)) * 256 + IData.bar(IData.pos + 1)
                    IData.pos = IData.pos + 2
                    Call putbits(MatchL, 16)
                End If
    
                x = NumDist
 
                Do While Dist(x) > MatchP
                    x = x - 1
                Loop
 
                Call PutBitsRev(Distance(x).Code, Distance(x).Length)
                Call putbits(MatchP - Dist(x), Dext(x))
            End If
        Loop
 
        Call Flush_Buffer
    End If
 
    ReDim Preserve Udata.bar(Udata.pos - 1)
    Bytearray = Udata.bar
    Erase Udata.bar
    Erase Hash
    Erase Hash1
    Erase HashSlot1
    Erase hashSlot2
    Erase Slot
End Sub

Private Function Compress_with_zLib(ByRef ArrayToCompress() As Byte, ByRef Return_Array() As Byte, ByVal CompressionLevel As Long) As Long
 Dim OrigSize As String
 Dim arrayLens As Long
 Dim arrayLenD As Long
 Dim CharCount As Long
 Dim MyCounter As Long
 Dim Return_ErrorCode As Long

 On Error GoTo ErrorTrap
 Return_ErrorCode = 0
 
 arrayLens = UBound(ArrayToCompress) + 1
 If arrayLens = 0 Then
 Compress_with_zLib = True
 Exit Function
 End If
 arrayLenD = arrayLens + ((arrayLens * 0.001) + 15) ' Extra 3 bytes added on for some extra padding (avoids errors)
 Erase Return_Array
 ReDim Return_Array(arrayLenD) As Byte
 'Return_ErrorCode = Compressn(Return_Array(0), arrayLenD, ArrayToCompress(0), arrayLens)
 Return_ErrorCode = Compress(Return_Array(0), arrayLenD, ArrayToCompress(0), arrayLens, CompressionLevel)
 ReDim Preserve Return_Array(arrayLenD - 1) As Byte
 Compress_with_zLib = Return_ErrorCode
 Exit Function
 
ErrorTrap:
 
 If Err.Number = 0 Then ' No Error
 Resume Next
 ElseIf Err.Number = 20 Then ' Resume Without Error
 Resume Next
 Else ' Unknown Error
 MsgBox Err.Source & " caused the following error :" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, " Error - " & Err.Description
 Compress_with_zLib = Err.Number
 Err.Clear
 End If
End Function

Private Sub CompressPNG(ArraytoComp() As Byte, withVB As Boolean)
Dim Fertigb() As Byte
Fertigb = ArraytoComp
Select Case withVB
Case 1 'VB slow
Compress_with_VB Fertigb, 9
ReDim ArraytoComp(UBound(Fertigb) + 2)
memCPY ArraytoComp(2), Fertigb(0), UBound(Fertigb) + 1
ArraytoComp(0) = 120
ArraytoComp(1) = 156
Case 0 'Dll fast
Compress_with_zLib Fertigb, ArraytoComp, 9
End Select
End Sub

Private Function Create_Codes(tree() As CodesType, Lenghts() As Long, NumCodes As Long) As Long
 Dim bits(16) As Long
 Dim next_code(16) As Long
 Dim Code As Long
 Dim LN As Long
 Dim x As Long
 Dim MaxBits As Long
 Dim Minbits As Long
 Minbits = 16
 For x = 0 To NumCodes
 bits(Lenghts(x)) = bits(Lenghts(x)) + 1
 If Lenghts(x) > MaxBits Then MaxBits = Lenghts(x)
 If Lenghts(x) < Minbits And Lenghts(x) > 0 Then Minbits = Lenghts(x)
If Lenghts(x) < 7 Then
 x = x
End If
 Next
 LN = 1
 For x = 1 To MaxBits
 LN = LN + LN
 LN = LN - bits(x)
 If LN < 0 Then Create_Codes = LN: Exit Function
 Next
 Create_Codes = LN
 Code = 0
 bits(0) = 0
 For x = 1 To MaxBits
 Code = (Code + bits(x - 1)) * 2
 next_code(x) = Code
 Next
 For x = 0 To NumCodes
 LN = Lenghts(x)
 If LN <> 0 Then
 tree(x).Length = LN
 tree(x).Code = next_code(LN)
 next_code(LN) = next_code(LN) + 1
 End If
 Next
End Function

Private Sub FindLongestMatch( _
    Winpos As Long, _
    FindMatch As Long, _
    NewSearch As Boolean)
    
    Static obj     As IUnknown
    Static asm(37) As Long
    Static tmp(11) As Long
        
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H5304528B:  asm(5) = &H5A8B5756
        asm(6) = &HC4B8B0C:   asm(7) = &H738B73E3:  asm(8) = &H4E6C108
        asm(9) = &H72036B7E:  asm(10) = &H1C428B10: asm(11) = &H3B08462B
        asm(12) = &H5D7F2042: asm(13) = &H4389068B: asm(14) = &H768B5108
        asm(15) = &H72035608: asm(16) = &H72035614: asm(17) = &H187A8B24
        asm(18) = &HC12C4A8B: asm(19) = &H167402E9: asm(20) = &H73B068B
        asm(21) = &HC6830975: asm(22) = &H4C78304:  asm(23) = &HB9EFEB49
        asm(24) = &H3:        asm(25) = &H4A8B08EB: asm(26) = &H3E1832C
        asm(27) = &H68A0B74:  asm(28) = &H575073A:  asm(29) = &H75494746
        asm(30) = &HF02B58F5: asm(31) = &H723B5958: asm(32) = &H89067E04
        asm(33) = &H42890472: asm(34) = &H28723B08: asm(35) = &H7549037D
        asm(36) = &H5B5E5F8D: asm(37) = &H4C2
    End If
    
    ' **************************************************************************
    
    Dim hvl As Long
    Dim max As Long
    
    MatchL = MinMatch - 1
    hvl = HashValue(Winpos)
    
    If hvl Then
        With Hash(hvl)
            If NewSearch Then .Pointer = .FirstSlot
            
            max = FileLength - Winpos
            If max > MaxMatch Then max = MaxMatch
            max = max - MinMatch
           
            tmp(1) = MatchL
            tmp(2) = MatchP
            tmp(3) = VarPtr(Hash(hvl))
            tmp(4) = VarPtr(Slot(0))
            tmp(5) = VarPtr(IData.bar(0))
            tmp(6) = VarPtr(IData.bar(Winpos + MinMatch))
            tmp(7) = Winpos
            tmp(8) = MaxHistory
            tmp(9) = MinMatch
            tmp(10) = FindMatch
            tmp(11) = max
            
            memCPY obj, asm(0), 4
            Set obj = Nothing
        
            MatchL = tmp(1)
            MatchP = tmp(2)
        End With
    End If
End Sub

Private Sub FindLongestMatch1(Winpos As Long, FindMatch As Long, NewSearch As Boolean)
    Dim hvl As Long
    Dim lpc As Long
    Dim pos As Long
    Dim sln As Long
    Dim tmp As Long
    Dim max As Long
    Dim dmy As Long
    
    MatchL = MinMatch - 1
    hvl = HashValue(Winpos)
    
    If hvl Then
        With Hash(hvl)
            If NewSearch Then .Pointer = .FirstSlot
            
            max = FileLength - Winpos
            If max > MaxMatch Then max = MaxMatch
            max = max - MinMatch
           
            For lpc = 1 To .NumSlots
                If .Pointer < 1 Then Exit For
                pos = Slot(.Pointer).Winpos
                If Winpos - pos > MaxHistory Then Exit For
            
                sln = pos + MinMatch
                tmp = Winpos + MinMatch
                dmy = max
                
                Do While IData.bar(sln) = IData.bar(tmp)
                    sln = sln + 1
                    tmp = tmp + 1
                    dmy = dmy - 1
                    If dmy = 0 Then Exit Do
                Loop
    
                sln = sln - pos
                .Pointer = Slot(.Pointer).Next
                
                If MatchL < sln Then
                    MatchL = sln
                    MatchP = pos
                End If
            
                If MatchL >= FindMatch Then Exit For
            Next
        End With
    End If
End Sub

Private Sub Flush_Buffer()
    If Udata.bit Then Call PutByte(Udata.val And 255)
End Sub

Private Function GetFreeHash()
 NxtFreeHash = NxtFreeHash + 1
 If NxtFreeHash > UBound(Hash) Then ReDim Preserve Hash(NxtFreeHash + 1000)
 GetFreeHash = NxtFreeHash
End Function

Private Function GetFreeHashSlot1()
 NxtFreeHashSlot1 = NxtFreeHashSlot1 + 1
 If NxtFreeHashSlot1 > UBound(HashSlot1) Then ReDim Preserve HashSlot1(NxtFreeHashSlot1 + 1000)
 GetFreeHashSlot1 = NxtFreeHashSlot1
End Function

Private Function GetFreeHashSlot2()
 NxtFreeHashSlot2 = NxtFreeHashSlot2 + 1
 If NxtFreeHashSlot2 > UBound(hashSlot2) Then ReDim Preserve hashSlot2(NxtFreeHashSlot2 + 1000)
 GetFreeHashSlot2 = NxtFreeHashSlot2
End Function

Private Function GetFreeSlot()
 NxtFreeSlot = NxtFreeSlot + 1
 If NxtFreeSlot > UBound(Slot) Then ReDim Preserve Slot(NxtFreeSlot + 1000)
 GetFreeSlot = NxtFreeSlot
End Function

Private Sub GetIconBits(bBits() As Byte, bpp As Long, nDC As Long, nBitmap As Long, CopyArr() As RGBQUAD, Height As Long, Breite As Long)
 Dim bi As BITMAPINFO8Bit
 Dim Test As Long
 If bpp > BMPBIT Then
 bpp = bpp - BMPBIT
 bi.bmiHeader.biHeight = 0 - Height
 Else
 bi.bmiHeader.biHeight = Height
 End If
 bi.bmiHeader.biBitCount = bpp
 bi.bmiHeader.biCompression = BI_RGB
 bi.bmiHeader.biPlanes = 1
 bi.bmiHeader.biWidth = Breite
 bi.bmiHeader.biSize = Len(bi.bmiHeader)
 Test = GetDIBits(nDC, nBitmap, 0, Height, bBits(0), bi, DIB_RGB_COLORS)
 If bpp = 1 Then
 memCPY CopyArr(0), bi.bmiColors(0), Len(CopyArr(0)) * 2
 ElseIf bpp = 4 Then
 memCPY CopyArr(0), bi.bmiColors(0), Len(CopyArr(0)) * 16
 ElseIf bpp = 8 Then
 memCPY CopyArr(0), bi.bmiColors(0), Len(CopyArr(0)) * 256
 End If
End Sub

Private Function HashValue(Winpos As Long) As Long
 Dim HVal As Long
 Dim NFS As Long
 HVal = Hash1(IData.bar(Winpos))
 If HVal = 0 Then Exit Function
 Do
 If HashSlot1(HVal).ToHash = 0 Then Exit Function
 If HashSlot1(HVal).Char = IData.bar(Winpos + 1) Then
 HVal = HashSlot1(HVal).ToSlot
 Exit Do
 End If
 HVal = HashSlot1(HVal).ToHash
 Loop
 Do
 If hashSlot2(HVal).ToHash = 0 Then Exit Function
 If hashSlot2(HVal).Char = IData.bar(Winpos + 2) Then
 HVal = hashSlot2(HVal).ToSlot
 Exit Do
 End If
 HVal = hashSlot2(HVal).ToHash
 Loop
 HashValue = HVal
End Function

Private Sub INIT(Level As Long, Deflate64 As Boolean)
 Dim Temp()
 Temp = Array(0, 0, 0, 0, _
 4, 4, 8, 4, _
 4, 5, 16, 8, _
 4, 6, 32, 32, _
 4, 4, 16, 16, _
 8, 16, 32, 32, _
 8, 16, 128, 128, _
 8, 32, 128, 256, _
 32, 128, 258, 1024, _
 32, 258, 258, 4096)
 GoodMatch = Temp(Level * 4)
 LazyMatch = Temp(Level * 4 + 1)
 NiceMatch = Temp(Level * 4 + 2)
 MaxChain = Temp(Level * 4 + 3)
 MaxHistory = 32767
 If Deflate64 Then MaxHistory = 65535
 MaxMatch = 258
 If Deflate64 Then MaxMatch = 65538
 NumDist = 29
 If Deflate64 Then NumDist = 31
 ReDim Slot(5 * MaxChain)
 ReDim Hash(100000)
 ReDim Hash1(255)
 ReDim HashSlot1(2000)
 ReDim hashSlot2(2000)
 Erase IData.bar
 ReDim Udata.bar(4096)
 IData.bit = 0
 IData.pos = 0
 Udata.bit = 0
 Udata.pos = 0
 NxtFreeSlot = 0
 NxtFreeHash = 0
 NxtFreeHashSlot1 = 0
 NxtFreeHashSlot2 = 0
End Sub

Public Sub init_Trees()
 Dim x As Long
 Dim Temp()
 Erase LitLen
 Erase Distance
 Erase Bl_Tree
 Erase DC
 Erase LL
 Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
 For x = 0 To UBound(Temp): Border(x) = Temp(x): Next
 Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
 For x = 0 To UBound(Temp): Lens(x) = Temp(x): Next
 Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
 For x = 0 To UBound(Temp): Lext(x) = Temp(x): Next
 Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
 For x = 0 To UBound(Temp): Dist(x) = Temp(x): Next
 Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
 For x = 0 To UBound(Temp): Dext(x) = Temp(x): Next

    Udata.bit = 0
    Udata.val = 0
End Sub

Private Sub InsertSlot(Winpos As Long)
 Dim HVal As Long
 Dim NFS As Long
 Dim NuSlot As Long
 HVal = HashValue(Winpos)
 If HVal = 0 Then
 HVal = Hash1(IData.bar(Winpos))
 If HVal = 0 Then
 HVal = GetFreeHashSlot1
 Hash1(IData.bar(Winpos)) = HVal
 End If
 Do
 If HashSlot1(HVal).ToHash = 0 Then
 NFS = GetFreeHashSlot1
 HashSlot1(HVal).ToHash = NFS
 NFS = GetFreeHashSlot2
 HashSlot1(HVal).ToSlot = NFS
 HashSlot1(HVal).Char = IData.bar(Winpos + 1)
 End If
 If HashSlot1(HVal).Char = IData.bar(Winpos + 1) Then
 HVal = HashSlot1(HVal).ToSlot
 Exit Do
 End If
 HVal = HashSlot1(HVal).ToHash
 Loop
 Do
 If hashSlot2(HVal).ToHash = 0 Then
 NFS = GetFreeHashSlot2
 hashSlot2(HVal).ToHash = NFS
 NFS = GetFreeHash
 hashSlot2(HVal).ToSlot = NFS
 hashSlot2(HVal).Char = IData.bar(Winpos + 2)
 End If
 If hashSlot2(HVal).Char = IData.bar(Winpos + 2) Then
 HVal = hashSlot2(HVal).ToSlot
 Exit Do
 End If
 HVal = hashSlot2(HVal).ToHash
 Loop
 End If
 If Hash(HVal).NumSlots = 0 Then
 NFS = GetFreeSlot
 Hash(HVal).FirstSlot = NFS
 Hash(HVal).LastSlot = NFS
 Hash(HVal).NumSlots = 1
 Slot(NFS).Winpos = Winpos
 Slot(NFS).Next = -1
 Exit Sub
 End If
 If Hash(HVal).NumSlots = MaxChain Then
 NFS = Hash(HVal).LastSlot
 Slot(Slot(NFS).Prev).Next = -1
 NuSlot = Hash(HVal).FirstSlot
 Hash(HVal).FirstSlot = NFS
 Slot(NFS).Next = NuSlot
 Slot(NFS).Winpos = Winpos
 Else
 NuSlot = Hash(HVal).FirstSlot
 NFS = GetFreeSlot
 Hash(HVal).FirstSlot = NFS
 Slot(NFS).Next = NuSlot
 Slot(NFS).Winpos = Winpos
 Slot(NuSlot).Prev = NFS
 Hash(HVal).NumSlots = Hash(HVal).NumSlots + 1
 End If
End Sub

Private Function MakeIDATs(PNGMap() As Byte, fno As Long, withVB As Boolean) As Long
Dim AdlerCRC As Long
Dim Gr As Long
Dim Teilgr As Long
Dim Teilanz As Long
Dim TeilRest As Long
Dim i As Long
Dim Bytearray() As Byte
Dim Stand As Long

Teilgr = 8192
 AdlerCRC = x_Adler32(PNGMap)  'IDAT
 AdlerCRC = x_lngSWP(AdlerCRC)
 CompressPNG PNGMap, withVB
 ReDim Preserve PNGMap(UBound(PNGMap) + 4)
 memCPY PNGMap(UBound(PNGMap) - 3), ByVal VarPtr(AdlerCRC), 4
 Gr = UBound(PNGMap) + 1
 If Gr > Teilgr Then
Teilanz = Gr \ Teilgr
TeilRest = Gr - (Teilgr * Teilanz)
For i = 1 To Teilanz
ReDim Bytearray(Teilgr - 1)
memCPY Bytearray(0), PNGMap(Stand), Teilgr
Stand = Stand + Teilgr
x_putSCT fno, "IDAT", VarPtr(Bytearray(0)), UBound(Bytearray) + 1
Next i
If TeilRest > 0 Then
ReDim Bytearray(TeilRest - 1)
memCPY Bytearray(0), PNGMap(Stand), TeilRest
x_putSCT fno, "IDAT", VarPtr(Bytearray(0)), UBound(Bytearray) + 1
End If
Else
x_putSCT fno, "IDAT", VarPtr(PNGMap(0)), UBound(PNGMap) + 1
End If

End Function

Private Sub MakePNG24(Feld() As Byte, PNGMap() As Byte, nWidth As Long, nHeight As Long, Bytesperline As Long)
 Dim x As Long, y As Long, nStartPos As Long
 Dim i As Long
 Dim High As Long
 Dim Linebytes As Long
 Dim Roh As Long
 Dim Fertig As Long
 Dim tt As Long
 Dim Ausgleich As Long
 Dim Übergabe() As Byte
 ReDim Übergabe(UBound(Feld))
Dim Standfertig As Long
Dim Standroh As Long
Dim Alphanummer As Long
Dim z As Long
memCPY Übergabe(0), Feld(0), UBound(Feld) + 1
If nWidth Mod 2 <> 0 Then
Ausgleich = 3
End If
 For i = 0 To UBound(Feld) Step 3 'RGB tauschen
 memCPY Übergabe(i), Feld(i + 2), 1
 memCPY Übergabe(i + 2), Feld(i), 1
 Next i
 memCPY Feld(0), Übergabe(0), UBound(Feld) + 1
 Linebytes = Bytesperline * 3
 Fertig = 1
 Roh = 0
 ReDim Übergabe(UBound(Feld) + nHeight)
 For i = 0 To nHeight - 1
 memCPY Übergabe(Fertig), Feld(Roh), Linebytes
 Roh = Roh + Linebytes + Ausgleich
 Fertig = Fertig + Linebytes + 1
 Next i
 If HasAlpha And m_AlphaSet Then
  If UBound(AlphaArray) = ((nWidth * nHeight) - 1) Then
  ReDim PNGMap((nHeight * nWidth * 4) - 1 + nHeight)
 For i = 0 To nHeight - 1
 Standfertig = Standfertig + 1 'Filter (immer 0)
 Standroh = Standroh + 1
 For z = 0 To nWidth - 1
 memCPY PNGMap(Standfertig), Übergabe(Standroh), 3
 Standfertig = Standfertig + 3
 PNGMap(Standfertig) = AlphaArray(Alphanummer)
 Alphanummer = Alphanummer + 1
 Standfertig = Standfertig + 1
 Standroh = Standroh + 3
 Next z
 Next i
 Else
  m_AlphaSet = 0
 PNGMap = Übergabe
End If
 Else
 m_AlphaSet = 0
 PNGMap = Übergabe
End If
End Sub

Private Sub MakePNGPalette(Bitmappalette() As RGBQUAD, PNGPalette() As Byte, nColors As Long)
 Dim cnt As Long
 
 For cnt = 0 To nColors - 1
 PNGPalette(cnt * 3) = Bitmappalette(cnt).rgbRed
 PNGPalette(cnt * 3 + 1) = Bitmappalette(cnt).rgbGreen
 PNGPalette(cnt * 3 + 2) = Bitmappalette(cnt).rgbBlue
 Next
End Sub

Private Sub OptimizeHuffman(Codes() As CodesType, Freq() As Long, NumCodes As Long, MaxBits As Long)
 Dim i As Long
 Dim j As Long
 Dim k As Long
 Dim N As Long
 Dim V1 As Long
 Dim v2 As Long
 Dim others() As Long
 Dim codesize() As Long
 Dim bits() As Long
 Dim swp As Long
 Dim swp2 As Long
 Dim freq2() As Long
 Dim Code As Long
 Dim next_code() As Long
 Dim LN As Long
 Dim CodLenPos() As Long
 Dim OverFlow As Boolean
 ReDim next_code(MaxBits)
 ReDim others(NumCodes)
 ReDim codesize(NumCodes)
 ReDim bits(MaxBits + 5)
 ReDim CodLenPos(MaxBits + 5, NumCodes + 1)
 
 freq2 = Freq
 v2 = 0
 For i = 0 To NumCodes
 others(i) = -1
 If Freq(i) <> 0 Then v2 = v2 + 1
 Next i
 If v2 = 1 Then
 For i = 0 To NumCodes
 If freq2(i) = 0 Then
 freq2(i) = 1
 Exit For
 End If
 Next
 End If

 Do
 V1 = -1
 v2 = -1
 swp = 2147483647
 swp2 = 2147483647
 For i = 0 To NumCodes
 If freq2(i) <> 0 Then
 If (freq2(i) <= swp2) Then
 If (freq2(i) <= swp) Then
 swp2 = swp
 v2 = V1
 swp = freq2(i)
 V1 = i
 Else
 swp2 = freq2(i)
 v2 = i
 End If
 End If
 End If
 Next i
 If v2 = -1 Then
 freq2(V1) = 0
 Exit Do
 End If
 freq2(V1) = freq2(V1) + freq2(v2)
 freq2(v2) = 0
 codesize(V1) = codesize(V1) + 1
 While (others(V1) >= 0)
 V1 = others(V1)
 codesize(V1) = codesize(V1) + 1
 Wend
 others(V1) = v2
 codesize(v2) = codesize(v2) + 1
 While (others(v2) >= 0)
 v2 = others(v2)
 codesize(v2) = codesize(v2) + 1
 Wend
 Loop

 N = 0
 For i = 0 To NumCodes
 If codesize(i) <> 0 Then
 bits(codesize(i)) = bits(codesize(i)) + 1
 CodLenPos(codesize(i), bits(codesize(i))) = i
 CodLenPos(codesize(i), 0) = bits(codesize(i))
 If N < codesize(i) Then N = codesize(i)
 End If
 Next i

 i = N
 While i > MaxBits
 While bits(i) > 0
 For j = i - 2 To 1 Step -1
 If bits(j) > 0 Then Exit For
 Next j
 bits(i) = bits(i) - 2
 bits(i - 1) = bits(i - 1) + 1
 bits(j + 1) = bits(j + 1) + 2
 bits(j) = bits(j) - 1
 Wend
 i = i - 1
 OverFlow = True
 Wend

 If OverFlow Then
 V1 = 0
 j = 0
 Do While bits(V1) = 0
 V1 = V1 + 1
 Loop
 v2 = bits(V1)
 Do While V1 <= MaxBits
 For i = 1 To CodLenPos(j, 0)
 If codesize(CodLenPos(j, i)) <> 0 Then
 codesize(CodLenPos(j, i)) = V1
 v2 = v2 - 1
 If v2 = 0 Then
 V1 = V1 + 1
 Do While bits(V1) = 0 And V1 <= MaxBits
 V1 = V1 + 1
 Loop
 v2 = bits(V1)
 End If
 End If
 Next
 j = j + 1
 Loop
 End If

 Code = 0
 For i = 1 To MaxBits
 Code = (Code + bits(i - 1)) * 2
 next_code(i) = Code
 Next
 For i = 0 To NumCodes
 LN = codesize(i)
 If LN <> 0 Then
 Codes(i).Length = LN
 Codes(i).Code = next_code(LN)
 next_code(LN) = next_code(LN) + 1
 End If
 Next
End Sub

Private Sub putbits(Value As Long, Numbits As Long)
    With Udata
        .val = .val + Value * 2 ^ .bit
        .bit = .bit + Numbits
    
        Do While .bit > 7
            Call PutByte(.val And 255)
            .val = Fix(.val / 256)
            .bit = .bit - 8
        Loop
    End With
End Sub

Private Sub putbits1(lng As Long, cnt As Long)
    Static obj     As IUnknown
    Static asm(21) As Long
    Static tmp(3)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H5604528B:  asm(5) = &H8B0C728B
        asm(6) = &H4E8B0442:  asm(7) = &H74C90A08:  asm(8) = &H3E0D302
        asm(9) = &H4A020C46:  asm(10) = &H8F98008:  asm(11) = &H53571E72
        asm(12) = &H7F8B3E8B: asm(13) = &H45E8B0C:  asm(14) = &H433B0488
        asm(15) = &H8008E8C1: asm(16) = &HF47308E9: asm(17) = &H80045E89
        asm(18) = &H5F5B08C1: asm(19) = &H89084E88: asm(20) = &HC25E0C46
        asm(21) = &H4
    End If
    
    If (Udata.pos And 1023) > 1019 Then
        ReDim Preserve Udata.bar(Udata.pos + 1024)
    End If
    
    tmp(1) = lng
    tmp(2) = cnt
    tmp(3) = VarPtr(Udata)
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
End Sub

Private Sub PutBitsRev(Value As Long, Numbits As Long)
    putbits x_lngRVS(Value, Numbits), Numbits
End Sub

Private Function PutByte(Char As Long) As Long
    If (Udata.pos And 1023) = 0 Then ReDim Preserve Udata.bar(Udata.pos + 1024)
    Udata.bar(Udata.pos) = Char
    PutByte = Udata.pos
    Udata.pos = Udata.pos + 1
End Function

Private Sub SavePNG( _
    fsc As String, _
    bar() As Byte, _
    wid As Long, _
    hgt As Long, _
    bpp As Long, _
    pal() As RGBQUAD, _
    withVB As Boolean)
    
    Dim Hilfslong As Long
    Dim hdr     As IHDR
    Dim fno     As Long
    Dim png()   As Byte
    Dim tmp()   As Byte
    Dim PNGPalette(0 To 767) As Byte
    Dim BkgdNumber As Byte
    
    fno = FreeFile
    If Dir$(fsc) <> "" Then Kill fsc
    
    If Dir$(fsc) <> "" Then
        MsgBox "could not delete file"
        
    ElseIf fno = 0 Then
        MsgBox "could not create filehandle"
        
    ElseIf bpp <> 24 And bpp <> 8 Then
        MsgBox "BitsPerPixel setting not supported"
        
    Else
    Select Case bpp
    Case 24
        hdr.Width = x_lngSWP(wid)
        hdr.Height = x_lngSWP(hgt)
        hdr.bpp = 8     ' ??
        hdr.ColorType = IIf(HasAlpha And m_AlphaSet, 6, 2)
        
        MakePNG24 bar, png, wid, hgt, wid
        
        Open fsc For Binary Access Read Write As #fno
        ' check errors!
        
        Put #fno, , Chr$(&H89) & "PNG"
        Put #fno, , &HA1A0A0D
        
        x_putSCT fno, "IHDR", VarPtr(hdr), Len(hdr)
 
        If HasbkgdColor Then x_putSCT fno, "bKGD", bkgdColorLong, -1
        If HasTrans Then x_putSCT fno, "tRNS", Transparent24, -1

        MakeIDATs png, fno, withVB

        tmp = StrConv("Comment" & vbNullChar & "Saved with AlKos SavePng.cls", vbFromUnicode)
        x_putSCT fno, "tEXt", VarPtr(tmp(0)), UBound(tmp) + 1
        x_putSCT fno, "IEND", 0, 0
    
    Case 8
        hdr.Width = x_lngSWP(wid)
        hdr.Height = x_lngSWP(hgt)
        hdr.bpp = 8     ' ??
        hdr.ColorType = 3
        MakePNGPalette pal, PNGPalette, 256
                Open fsc For Binary Access Read Write As #fno
        ' check errors!
        MakePNG8 bar, png, wid, hgt
        Put #fno, , Chr$(&H89) & "PNG"
        Put #fno, , &HA1A0A0D
        
        x_putSCT fno, "IHDR", VarPtr(hdr), Len(hdr)
        x_putSCT fno, "PLTE", VarPtr(PNGPalette(0)), UBound(PNGPalette) + 1
        If HasbkgdColor Then
        Hilfslong = FindPalNumber(PNGPalette, bkgdColorLong)
        If Hilfslong <> -1 Then
        BkgdNumber = CByte(Hilfslong)
        x_putSCT fno, "bKGD", VarPtr(BkgdNumber), 1
        End If
        End If
        If HasTrans Then x_putSCT fno, "tRNS", VarPtr(TransArray8(0)), UBound(TransArray8) + 1
        MakeIDATs png, fno, withVB

        tmp = StrConv("Comment" & vbNullChar & "Saved with AlKos SavePng.cls", vbFromUnicode)
        x_putSCT fno, "tEXt", VarPtr(tmp(0)), UBound(tmp) + 1
        x_putSCT fno, "IEND", 0, 0

        End Select
    End If
    If fno Then Close #fno
End Sub

Public Sub SavePNGinFile(Filename As String, pic As PictureBox, BitsPPix As Long, Optional withVB As Boolean = False)
 Dim Übergabe() As Byte
 Dim Temppal(0 To 255) As RGBQUAD
 GetBitmapDataTemp pic, BitsPPix, Übergabe, Temppal
 SavePNG Filename, Übergabe(), BildbreiteRight, Bildhöhe, BitCnt, Temppal(), withVB
End Sub

Private Sub Scan_Tree(tree() As CodesType, maxcode As Long)
 Dim N As Long
 Dim Curlen As Long, Nextlen As Long, PrevLen As Long
 Dim Count As Long
 Dim Max_Count As Long
 Dim Min_Count As Long
 Do While tree(maxcode).Length = 0
 maxcode = maxcode - 1
 Loop
 PrevLen = -1
 Nextlen = tree(0).Length
 Max_Count = 7
 Min_Count = 4
 If Nextlen = 0 Then Max_Count = 138: Min_Count = 3

 For N = 0 To maxcode
 Curlen = Nextlen
 If N <> maxcode Then Nextlen = tree(N + 1).Length Else Nextlen = -1
 Count = Count + 1
 If (Count < Max_Count) And (Curlen = Nextlen) Then
'Do nothing
 Else
 If Count < Min_Count Then
 Bl_Tree(Curlen).Code = Bl_Tree(Curlen).Code + Count
 ElseIf Curlen <> 0 Then
 If Curlen <> PrevLen Then Bl_Tree(Curlen).Code = Bl_Tree(Curlen).Code + 1
 Bl_Tree(16).Code = Bl_Tree(16).Code + 1
 ElseIf Count <= 10 Then
 Bl_Tree(17).Code = Bl_Tree(17).Code + 1
 Else
 Bl_Tree(18).Code = Bl_Tree(18).Code + 1
 End If
 Count = 0
 PrevLen = Curlen
 If Nextlen = 0 Then
 Max_Count = 138
 Min_Count = 3
 ElseIf Curlen = Nextlen Then
 Max_Count = 6
 Min_Count = 3
 Else
 Max_Count = 7
 Min_Count = 4
 End If
 End If
 Next
End Sub

Private Sub Send_Tree(tree() As CodesType, maxcode As Long)
    Dim lpc As Long
    Dim act As Long
    Dim nxt As Long
    Dim pvs As Long
    Dim cnt As Long
    Dim max As Long
    Dim min As Long
    
    Do While tree(maxcode).Length = 0
        maxcode = maxcode - 1
    Loop
 
    pvs = -1
    nxt = tree(0).Length
    min = IIf(nxt, 4, 3)
    max = IIf(nxt, 7, 138)

    For lpc = 0 To maxcode
        act = nxt
        cnt = cnt + 1
        
        If lpc < maxcode Then
            nxt = tree(lpc + 1).Length
        Else
            nxt = -1
        End If
        
        If (cnt >= max) Or (act <> nxt) Then
            If cnt < min Then
                Do While cnt
                    Call PutBitsRev(Bl_Tree(act).Code, Bl_Tree(act).Length)
                    cnt = cnt - 1
                Loop
 
            ElseIf act Then
                If act <> pvs Then
                    Call PutBitsRev(Bl_Tree(act).Code, Bl_Tree(act).Length)
                    cnt = cnt - 1
                End If
 
                Call PutBitsRev(Bl_Tree(16).Code, Bl_Tree(16).Length)
                Call putbits(cnt - 3, 2)
 
            ElseIf cnt < 11 Then
                Call PutBitsRev(Bl_Tree(17).Code, Bl_Tree(17).Length)
                Call putbits(cnt - 3, 3)
 
            Else
                Call PutBitsRev(Bl_Tree(18).Code, Bl_Tree(18).Length)
                Call putbits(cnt - 11, 7)
            End If
 
            cnt = 0
            pvs = act
            
            Select Case nxt
            Case 0:     min = 3: max = 138
            Case act:   min = 3: max = 6
            Case Else:  min = 4: max = 7
            End Select
        End If
    Next
End Sub

Public Sub SetAlphaBytes(bar() As Byte)
    m_AlphaSet = UBound(bar) + 1
    ReDim AlphaArray(m_AlphaSet - 1)
    memCPY AlphaArray(0), bar(0), m_AlphaSet
End Sub

Function x_Adler32( _
    bar() As Byte, _
    Optional crc As Long = 1) As Long
    Static obj     As IUnknown
    Static asm(21) As Long
    Static tmp(4)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H55575653:  asm(4) = &H1424748B:  asm(5) = &HF04768B
        asm(6) = &HF0C5EB7:   asm(7) = &H8B0E56B7:  asm(8) = &H7E8B084E
        asm(9) = &H106E8B04:  asm(10) = &H99C03352: asm(11) = &HD303178A
        asm(12) = &H372D53B:  asm(13) = &H8BF5F792: asm(14) = &HD00358DA
        asm(15) = &H572D53B:  asm(16) = &HF792C033: asm(17) = &H794947F5
        asm(18) = &H10E2C1DF: asm(19) = &H1689D30B: asm(20) = &H5B5E5F5D
        asm(21) = &H4C2
    End If
    
    tmp(1) = VarPtr(bar(0))
    tmp(2) = UBound(bar)
    tmp(3) = crc
    tmp(4) = 65521      ' const
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_Adler32 = tmp(0)
End Function

Function x_lngCRC( _
    ptr As Long, _
    cnt As Long, _
    Optional crc As Long) As Long
    
    Static obj      As IUnknown
    Static asm(15)  As Long
    Static tmp(4)   As Long
    Static tbl(255) As Long
    
    If asm(0) = 0 Then
        tbl(0) = &H0:          tbl(1) = &H77073096:   tbl(2) = &HEE0E612C:    tbl(3) = &H990951BA
        tbl(4) = &H76DC419:    tbl(5) = &H706AF48F:   tbl(6) = &HE963A535:    tbl(7) = &H9E6495A3
        tbl(8) = &HEDB8832:    tbl(9) = &H79DCB8A4:   tbl(10) = &HE0D5E91E:   tbl(11) = &H97D2D988
        tbl(12) = &H9B64C2B:   tbl(13) = &H7EB17CBD:  tbl(14) = &HE7B82D07:   tbl(15) = &H90BF1D91
        tbl(16) = &H1DB71064:  tbl(17) = &H6AB020F2:  tbl(18) = &HF3B97148:   tbl(19) = &H84BE41DE
        tbl(20) = &H1ADAD47D:  tbl(21) = &H6DDDE4EB:  tbl(22) = &HF4D4B551:   tbl(23) = &H83D385C7
        tbl(24) = &H136C9856:  tbl(25) = &H646BA8C0:  tbl(26) = &HFD62F97A:   tbl(27) = &H8A65C9EC
        tbl(28) = &H14015C4F:  tbl(29) = &H63066CD9:  tbl(30) = &HFA0F3D63:   tbl(31) = &H8D080DF5
        tbl(32) = &H3B6E20C8:  tbl(33) = &H4C69105E:  tbl(34) = &HD56041E4:   tbl(35) = &HA2677172
        tbl(36) = &H3C03E4D1:  tbl(37) = &H4B04D447:  tbl(38) = &HD20D85FD:   tbl(39) = &HA50AB56B
        tbl(40) = &H35B5A8FA:  tbl(41) = &H42B2986C:  tbl(42) = &HDBBBC9D6:   tbl(43) = &HACBCF940
        tbl(44) = &H32D86CE3:  tbl(45) = &H45DF5C75:  tbl(46) = &HDCD60DCF:   tbl(47) = &HABD13D59
        tbl(48) = &H26D930AC:  tbl(49) = &H51DE003A:  tbl(50) = &HC8D75180:   tbl(51) = &HBFD06116
        tbl(52) = &H21B4F4B5:  tbl(53) = &H56B3C423:  tbl(54) = &HCFBA9599:   tbl(55) = &HB8BDA50F
        tbl(56) = &H2802B89E:  tbl(57) = &H5F058808:  tbl(58) = &HC60CD9B2:   tbl(59) = &HB10BE924
        tbl(60) = &H2F6F7C87:  tbl(61) = &H58684C11:  tbl(62) = &HC1611DAB:   tbl(63) = &HB6662D3D
        tbl(64) = &H76DC4190:  tbl(65) = &H1DB7106:   tbl(66) = &H98D220BC:   tbl(67) = &HEFD5102A
        tbl(68) = &H71B18589:  tbl(69) = &H6B6B51F:   tbl(70) = &H9FBFE4A5:   tbl(71) = &HE8B8D433
        tbl(72) = &H7807C9A2:  tbl(73) = &HF00F934:   tbl(74) = &H9609A88E:   tbl(75) = &HE10E9818
        tbl(76) = &H7F6A0DBB:  tbl(77) = &H86D3D2D:   tbl(78) = &H91646C97:   tbl(79) = &HE6635C01
        tbl(80) = &H6B6B51F4:  tbl(81) = &H1C6C6162:  tbl(82) = &H856530D8:   tbl(83) = &HF262004E
        tbl(84) = &H6C0695ED:  tbl(85) = &H1B01A57B:  tbl(86) = &H8208F4C1:   tbl(87) = &HF50FC457
        tbl(88) = &H65B0D9C6:  tbl(89) = &H12B7E950:  tbl(90) = &H8BBEB8EA:   tbl(91) = &HFCB9887C
        tbl(92) = &H62DD1DDF:  tbl(93) = &H15DA2D49:  tbl(94) = &H8CD37CF3:   tbl(95) = &HFBD44C65
        tbl(96) = &H4DB26158:  tbl(97) = &H3AB551CE:  tbl(98) = &HA3BC0074:   tbl(99) = &HD4BB30E2
        tbl(100) = &H4ADFA541: tbl(101) = &H3DD895D7: tbl(102) = &HA4D1C46D:  tbl(103) = &HD3D6F4FB
        tbl(104) = &H4369E96A: tbl(105) = &H346ED9FC: tbl(106) = &HAD678846:  tbl(107) = &HDA60B8D0
        tbl(108) = &H44042D73: tbl(109) = &H33031DE5: tbl(110) = &HAA0A4C5F:  tbl(111) = &HDD0D7CC9
        tbl(112) = &H5005713C: tbl(113) = &H270241AA: tbl(114) = &HBE0B1010:  tbl(115) = &HC90C2086
        tbl(116) = &H5768B525: tbl(117) = &H206F85B3: tbl(118) = &HB966D409:  tbl(119) = &HCE61E49F
        tbl(120) = &H5EDEF90E: tbl(121) = &H29D9C998: tbl(122) = &HB0D09822:  tbl(123) = &HC7D7A8B4
        tbl(124) = &H59B33D17: tbl(125) = &H2EB40D81: tbl(126) = &HB7BD5C3B:  tbl(127) = &HC0BA6CAD
        tbl(128) = &HEDB88320: tbl(129) = &H9ABFB3B6: tbl(130) = &H3B6E20C:  tbl(131) = &H74B1D29A
        tbl(132) = &HEAD54739: tbl(133) = &H9DD277AF: tbl(134) = &H4DB2615:  tbl(135) = &H73DC1683
        tbl(136) = &HE3630B12: tbl(137) = &H94643B84: tbl(138) = &HD6D6A3E:  tbl(139) = &H7A6A5AA8
        tbl(140) = &HE40ECF0B: tbl(141) = &H9309FF9D: tbl(142) = &HA00AE27:  tbl(143) = &H7D079EB1
        tbl(144) = &HF00F9344: tbl(145) = &H8708A3D2: tbl(146) = &H1E01F268: tbl(147) = &H6906C2FE
        tbl(148) = &HF762575D: tbl(149) = &H806567CB: tbl(150) = &H196C3671: tbl(151) = &H6E6B06E7
        tbl(152) = &HFED41B76: tbl(153) = &H89D32BE0: tbl(154) = &H10DA7A5A: tbl(155) = &H67DD4ACC
        tbl(156) = &HF9B9DF6F: tbl(157) = &H8EBEEFF9: tbl(158) = &H17B7BE43: tbl(159) = &H60B08ED5
        tbl(160) = &HD6D6A3E8: tbl(161) = &HA1D1937E: tbl(162) = &H38D8C2C4: tbl(163) = &H4FDFF252
        tbl(164) = &HD1BB67F1: tbl(165) = &HA6BC5767: tbl(166) = &H3FB506DD: tbl(167) = &H48B2364B
        tbl(168) = &HD80D2BDA: tbl(169) = &HAF0A1B4C: tbl(170) = &H36034AF6: tbl(171) = &H41047A60
        tbl(172) = &HDF60EFC3: tbl(173) = &HA867DF55: tbl(174) = &H316E8EEF: tbl(175) = &H4669BE79
        tbl(176) = &HCB61B38C: tbl(177) = &HBC66831A: tbl(178) = &H256FD2A0: tbl(179) = &H5268E236
        tbl(180) = &HCC0C7795: tbl(181) = &HBB0B4703: tbl(182) = &H220216B9: tbl(183) = &H5505262F
        tbl(184) = &HC5BA3BBE: tbl(185) = &HB2BD0B28: tbl(186) = &H2BB45A92: tbl(187) = &H5CB36A04
        tbl(188) = &HC2D7FFA7: tbl(189) = &HB5D0CF31: tbl(190) = &H2CD99E8B: tbl(191) = &H5BDEAE1D
        tbl(192) = &H9B64C2B0: tbl(193) = &HEC63F226: tbl(194) = &H756AA39C: tbl(195) = &H26D930A
        tbl(196) = &H9C0906A9: tbl(197) = &HEB0E363F: tbl(198) = &H72076785: tbl(199) = &H5005713
        tbl(200) = &H95BF4A82: tbl(201) = &HE2B87A14: tbl(202) = &H7BB12BAE: tbl(203) = &HCB61B38
        tbl(204) = &H92D28E9B: tbl(205) = &HE5D5BE0D: tbl(206) = &H7CDCEFB7: tbl(207) = &HBDBDF21
        tbl(208) = &H86D3D2D4: tbl(209) = &HF1D4E242: tbl(210) = &H68DDB3F8: tbl(211) = &H1FDA836E
        tbl(212) = &H81BE16CD: tbl(213) = &HF6B9265B: tbl(214) = &H6FB077E1: tbl(215) = &H18B74777
        tbl(216) = &H88085AE6: tbl(217) = &HFF0F6A70: tbl(218) = &H66063BCA: tbl(219) = &H11010B5C
        tbl(220) = &H8F659EFF: tbl(221) = &HF862AE69: tbl(222) = &H616BFFD3: tbl(223) = &H166CCF45
        tbl(224) = &HA00AE278: tbl(225) = &HD70DD2EE: tbl(226) = &H4E048354: tbl(227) = &H3903B3C2
        tbl(228) = &HA7672661: tbl(229) = &HD06016F7: tbl(230) = &H4969474D: tbl(231) = &H3E6E77DB
        tbl(232) = &HAED16A4A: tbl(233) = &HD9D65ADC: tbl(234) = &H40DF0B66: tbl(235) = &H37D83BF0
        tbl(236) = &HA9BCAE53: tbl(237) = &HDEBB9EC5: tbl(238) = &H47B2CF7F: tbl(239) = &H30B5FFE9
        tbl(240) = &HBDBDF21C: tbl(241) = &HCABAC28A: tbl(242) = &H53B39330: tbl(243) = &H24B4A3A6
        tbl(244) = &HBAD03605: tbl(245) = &HCDD70693: tbl(246) = &H54DE5729: tbl(247) = &H23D967BF
        tbl(248) = &HB3667A2E: tbl(249) = &HC4614AB8: tbl(250) = &H5D681B02: tbl(251) = &H2A6F2B94
        tbl(252) = &HB40BBE37: tbl(253) = &HC30C8EA1: tbl(254) = &H5A05DF1B: tbl(255) = &H2D02EF8D
        
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H5304528B:  asm(5) = &H728B5652
        asm(6) = &H84A8B04:   asm(7) = &H8B105A8B:  asm(8) = &HD0F70C42
        asm(9) = &HD23310E3:  asm(10) = &H1632D08A: asm(11) = &H3308E8C1
        asm(12) = &H49469304: asm(13) = &H5A5EF275: asm(14) = &H89D0F75B
        asm(15) = &H4C202
    End If
    
    tmp(1) = ptr
    tmp(2) = cnt
    tmp(3) = crc
    tmp(4) = VarPtr(tbl(0))
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_lngCRC = tmp(0)
End Function

Function x_lngINC( _
    scn As Long, _
    cnt As Long, _
    src() As Long, _
    tgt() As Long, _
    Optional off As Long) As Long
    
    Static obj     As IUnknown
    Static asm(13) As Long
    Static tmp(5)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H5604528B:  asm(5) = &H8B04428B
        asm(6) = &H728B084A:  asm(7) = &H7849410C:  asm(8) = &H8E04390E
        asm(9) = &H4A03F87F:  asm(10) = &H10728B14: asm(11) = &H5E8E04FF
        asm(12) = &H4C20A89:  asm(13) = &H0
    End If
    
    tmp(1) = scn
    tmp(2) = cnt
    tmp(3) = VarPtr(src(0))
    tmp(4) = VarPtr(tgt(0))
    tmp(5) = off
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_lngINC = tmp(0)
End Function

Function x_lngRVS( _
    lng As Long, _
    Optional cnt As Long = 32) As Long
    
    Static obj     As IUnknown
    Static asm(10) As Long
    Static tmp(2)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))
    
        asm(3) = &H424548B:   asm(4) = &H5204528B:  asm(5) = &H8B084A8B
        asm(6) = &HC0330452:  asm(7) = &HD0D1EAD1:  asm(8) = &H5AF97F49
        asm(9) = &H4C20289:   asm(10) = &H0
    End If
    
    tmp(1) = lng
    tmp(2) = cnt
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_lngRVS = tmp(0)
End Function

Function x_lngSHR(lng As Long, Optional cnt As Long = 8) As Long
    Static obj     As IUnknown
    Static asm(7)  As Long
    Static tmp(2)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H8B04528B:  asm(5) = &H4A8B0442
        asm(6) = &H89E8D308:  asm(7) = &H4C202
    End If
    
    tmp(1) = lng
    tmp(2) = cnt
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_lngSHR = tmp(0)
End Function

Function x_lngSWP(lng As Long) As Long
    Static obj     As IUnknown
    Static asm(7)  As Long
    Static tmp(1)  As Long
    
    If asm(0) = 0 Then
        asm(0) = VarPtr(asm(0))
        asm(1) = VarPtr(tmp(0))
        asm(2) = VarPtr(asm(3))

        asm(3) = &H424548B:   asm(4) = &H8B04528B:  asm(5) = &HC80F0442
        asm(6) = &H4C20289:   asm(7) = &H0
    End If
    
    tmp(1) = lng
    
    memCPY obj, asm(0), 4
    Set obj = Nothing
    x_lngSWP = tmp(0)
End Function

Private Function x_putSCT( _
    fno As Long, _
    trv As String, _
    ptr As Long, _
    cnt As Long) As Long

    Dim crc     As Long
    Dim bar()   As Byte
    Dim tmp(5)  As Byte
    Dim adr     As Long
    Dim svd(1)  As Long
    
    bar = StrConv(trv, vbFromUnicode)
    crc = x_lngCRC(VarPtr(bar(0)), UBound(bar) + 1)

    Select Case cnt
    Case -1  ' color
        memCPY bar(0), ptr, 4
        tmp(1) = bar(0)
        tmp(3) = bar(1)
        tmp(5) = bar(2)
        
        ptr = VarPtr(tmp(0))
        cnt = 6
    End Select

    Put #fno, , x_lngSWP(cnt)
    Put #fno, , trv
    
    If cnt Then
        crc = x_lngCRC(ptr, cnt, crc)
        adr = 0# + Not (Not bar)
        
        memCPY svd(0), ByVal adr + 12, 8
        memCPY ByVal adr + 12, ptr, 4
        memCPY ByVal adr + 16, cnt, 4
        Put #fno, , bar
        memCPY ByVal adr + 12, svd(0), 8
    End If
    
    Put #fno, , x_lngSWP(crc)
End Function


Private Sub MakePNG8(Roh() As Byte, Fertig() As Byte, Width As Long, Height As Long)
'Filterbyte einfügen
Dim Bytesperline As Long
Dim BMPLinewidth As Long
Dim Gesamtgröße As Long
Dim i As Long
Dim h As Long
Dim j As Long

BMPLinewidth = (UBound(Roh) + 1) / Height 'wenn ungerade dann 1 mehr als Width
Gesamtgröße = Height * Width
Bytesperline = Width
ReDim Fertig(Gesamtgröße + Height - 1)
h = 1
For i = 0 To Height - 1
CopyMemory Fertig(h), Roh(j), Bytesperline
h = h + Bytesperline + 1
j = j + BMPLinewidth
Next i
End Sub

Private Function FindPalNumber(Palett() As Byte, ColorLong As Long) As Long
Dim i As Long
Dim Stand As Long
Dim Zähler As Long

FindPalNumber = -1 'Falls nicht gefunden
For i = 0 To 255
If RGB(Palett(Zähler), Palett(Zähler + 1), Palett(Zähler + 2)) = ColorLong Then
FindPalNumber = i
Exit For
End If
Zähler = Zähler + 3
Next i

End Function

Public Sub SetTransBytes8(Transarray() As Byte)
On Error GoTo Fehler
ReDim TransArray8(UBound(Transarray))
CopyMemory TransArray8(0), Transarray(0), UBound(Transarray) + 1
Exit Sub
Fehler:
HasTrans = False
End Sub

Private Sub Class_Initialize()
ReDim AlphaArray(0)
ReDim TransArray8(0)
End Sub

Private Sub GetBitmapDataTemp(pic As PictureBox, BitsPPix As Long, Bitmaparray() As Byte, Temppal() As RGBQUAD)
 Dim a As Long
 Dim i As Long
 Dim iBitmap As Long, iDC As Long
 Dim b As Long
 Dim bBits() As Byte
 Dim Dibbreite As Long
 Dim Groesse As Long
 Dim Style As Long
 Dim BMP8 As New Palette8Bit
 Dim aDC As Long
 Dim aPrevBmp As Long
 Dim bi24BitInfo As BITMAPINFO
 BitCnt = BitsPPix
 a = pic.Parent.ScaleMode
 b = pic.ScaleMode
 pic.Parent.ScaleMode = 3
 pic.ScaleMode = 3
 Bildhöhe = pic.ScaleHeight
 Bildbreite = pic.ScaleWidth
 pic.Parent.ScaleMode = a
 pic.ScaleMode = b
 BildbreiteRight = Bildbreite
 Select Case BitsPPix
 Case 8
 If Bildbreite Mod 2 <> 0 Then
 Bildbreite = Bildbreite + 1
 End If

 Bitcbr = Bildbreite

 BitCnt = 8
 If Bildbreite Mod 4 = 0 Then '8 bit
 Dibbreite = Bildbreite
 Else
 Dibbreite = Bildbreite + 2
 End If
BMP8.Save8Bpp pic, App.Path & "\test.bmp"
 Case 24
 If Bildbreite Mod 2 <> 0 Then
 Bildbreite = Bildbreite + 1
 End If
  Bitcbr = Bildbreite * 3
 BitCnt = 24
 If Bildbreite * 3 Mod 4 = 0 Then
 Dibbreite = Bildbreite * 3
 Else
 Dibbreite = Bildbreite * 3 + 2
 End If
 ReDim bBits((Dibbreite) * Bildhöhe - 1)
 With bi24BitInfo.bmiHeader
 .biBitCount = 24
 .biCompression = BI_RGB
 .biPlanes = 1
 .biSize = Len(bi24BitInfo.bmiHeader)
 .biWidth = Bildbreite
 .biHeight = Bildhöhe
 End With

 iDC = CreateCompatibleDC(pic.hdc)
 iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
 SelectObject iDC, iBitmap
 
 BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, pic.hdc, 0, 0, vbSrcCopy

 GetIconBits bBits(), BitCnt + BMPBIT, iDC, iBitmap, Temppal(), Bildhöhe, Bildbreite
 End Select
 
 ReDim Bitmaparray(Bitcbr * Bildhöhe - 1)
 If Bitcbr <> Dibbreite Then
 For i = 0 To Bildhöhe - 1
 memCPY Bitmaparray(i * Bitcbr), bBits(i * Dibbreite), Bitcbr
 Next i
 Else
 memCPY Bitmaparray(0), bBits(0), UBound(bBits) + 1
 End If
 DeleteDC iDC
 DeleteObject iBitmap
End Sub


Public Sub GetBitmapData(PicBox As PictureBox, BitsPerPix As Long, Bitmaparray() As Byte, Palettenarray() As Byte)
  Dim i As Long
  Dim Temppal(0 To 255) As RGBQUAD
Dim Zähler As Long
 GetBitmapDataTemp PicBox, BitsPerPix, Bitmaparray, Temppal
If BitsPerPix = 8 Then
ReDim Palettenarray(767)
For i = 0 To 255
Palettenarray(Zähler) = Temppal(i).rgbRed
Palettenarray(Zähler + 1) = Temppal(i).rgbGreen
Palettenarray(Zähler + 2) = Temppal(i).rgbBlue
Zähler = Zähler + 3
Next i
End If
End Sub
